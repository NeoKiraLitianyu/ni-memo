"""Deterministic schema discovery. Ambiguity is data, never a guess."""
from dataclasses import dataclass
import re
import unicodedata

from ni_memo.evidence import EvidenceItem


@dataclass(frozen=True)
class Binding:
    key: str
    evidence: EvidenceItem
    rule: str
    confidence: str
    score: int

    def to_dict(self):
        return {
            "key": self.key, "source_uri": self.evidence.source_uri,
            "source_location": self.evidence.source_location, "rule": self.rule,
            "confidence": self.confidence,
        }


@dataclass(frozen=True)
class DiscoveryResult:
    bindings: dict[str, Binding]
    missing: tuple[str, ...]
    mapping_conflicts: dict[str, tuple[Binding, ...]]

    def to_dict(self):
        return {
            "bindings": {key: self.bindings[key].to_dict() for key in sorted(self.bindings)},
            "missing": list(self.missing),
            "mapping_conflicts": {
                key: [item.to_dict() for item in self.mapping_conflicts[key]]
                for key in sorted(self.mapping_conflicts)
            },
        }


def discover(schema, evidence):
    evidence = sorted(evidence, key=lambda item: (item.source_uri.lower(), item.source_location))
    # 跨文件身份隔离: 当同时存在数据文件 (xlsx/csv/json/pdf) 与文档文件 (docx) 时,
    # 身份类字段 (company./business./team.) 的事实只能来自数据文件.
    # 历史 memo DOCX 是叙述/视觉参考, 不是新项目的公司身份事实源 —
    # 否则 "GL xlsx + 海古德 memo" 会把海古德的静电卡盘故事编造进 GL memo (T8 违例).
    _has_data_file = any(
        item.source_type in {"xlsx", "csv", "json", "pdf"} for item in evidence
    )
    _has_docx = any(item.source_type == "docx" for item in evidence)
    _isolation_active = _has_data_file and _has_docx

    candidates = {}
    for field in schema.fields:
        for item in evidence:
            if not _source_compatible(field, item):
                continue
            if _isolation_active and field.key.startswith(
                ("company.", "business.", "team.", "investment.", "risk.")
            ):
                if item.source_type == "docx":
                    # 历史 memo 叙述 (公司身份/投资亮点/风险) 不得充当数据文件项目的
                    # 事实源 — 否则 "GL xlsx + 海古德 memo" 会把海古德的静电卡盘故事
                    # 编造进 GL memo (T8 违例). 这些字段保持 missing → pending 显式标注.
                    continue
            score, rule = _match(field, item)
            if (not score or not _candidate_usable(item)
                    or not _value_compatible(field.value_kind, item.value)
                    or not _unit_compatible(field.unit_family, item.unit)):
                continue
            match_score = score
            score = match_score * 100 + _source_preference(item)
            key = _binding_key(field, item)
            if key:
                candidates.setdefault(key, []).append(Binding(
                    key, item, rule, "high" if match_score >= 95 else "medium", score
                ))
    bindings, conflicts = {}, {}
    for key, values in sorted(candidates.items()):
        ranked = sorted(values, key=lambda value: (-value.score, value.evidence.source_uri.lower(),
                                                    value.evidence.source_location))
        best = [value for value in ranked if value.score == ranked[0].score]
        unique_locations = {(value.evidence.source_uri, value.evidence.source_location) for value in best}
        if len(unique_locations) == 1:
            bindings[key] = best[0]
        elif _same_value(best):
            bindings[key] = best[0]
        else:
            conflicts[key] = tuple(best)
    present = set(bindings) | set(conflicts)
    missing = tuple(sorted(field.key for field in schema.fields
                           if not any(key == field.key or key.startswith(field.key + ".") for key in present)))
    return DiscoveryResult(bindings, missing, conflicts)


def _match(field, item):
    label = _norm(item.label)
    contexts = [_norm(value) for value in item.context]
    composite_contexts = contexts if _generic_label(label) else contexts[:1]
    composites = ({context + label for context in composite_contexts}
                  | {label + context for context in composite_contexts})
    best = (0, "")
    for alias in field.aliases:
        target = _norm(alias)
        if not target:
            continue
        context_contains_target = any(target in context for context in contexts)
        if label == target:
            best = max(best, (100, "exact_alias"))
        elif target not in label and not context_contains_target and target in composites:
            best = max(best, (105, "context_label_alias"))
        elif (target not in label and not context_contains_target and len(target) >= 3
              and any(target in composite for composite in composites)):
            best = max(best, (103, "contained_context_label_alias"))
        elif len(target) >= 3 and target in label:
            best = max(best, (80, "contained_alias"))
        elif len(target) >= 2 and any(target in context for context in contexts):
            # 2 字符中文 alias (投前/投后/估值) 在列头 context 中的命中同样有效;
            # 3 字符门槛对 CJK 双字词不公平 (T8: 投前估值缺失).
            best = max(best, (75, "contained_context_alias"))
    return best


def _same_value(bindings):
    values = [binding.evidence.value for binding in bindings]
    first = values[0]
    return all(type(value) is type(first) and value == first for value in values[1:])


def _binding_key(field, item):
    if field.value_kind in {"list", "person_list"}:
        suffix = _norm(item.source_location) or "item"
        return f"{field.key}.{suffix}"
    if field.value_kind != "series":
        return field.key
    period = _period(item)
    if not period:
        return None
    # Row-level breakdown (e.g. 高铁/城轨/合计市场规模, 敏感性行) must not
    # collapse into one series key: keep the row label as a suffix so distinct
    # granularities coexist as separate keys instead of fabricated conflicts.
    row_dim = _row_dimension(item)
    return f"{field.key}.{period}" + (f"·{row_dim}" if row_dim else "")


def _row_dimension(item):
    """Extract a granularity token only when the row label names a real breakdown.

    Generic row labels (毛利率/净利润/合计/市场规模...) are dimensions of the
    field itself, not sub-buckets: splitting them fragments the year series and
    fabricates parent conflicts. Only explicit breakdown tokens (高铁/城轨/能源/
    轨交SiC/无锡/东台...) become a key suffix.
    """
    label = _norm(item.label)
    generic = {"margin", "rate", "ratio", "growth", "amount", "value", "比例", "增速",
               "增长率", "金额", "规模", "市场规模", "合计", "总量", "净利润", "毛利润",
               "毛利率", "净利率", "资产负债率", "货币资金", "总资产", "归母净利润", "ebitda"}
    if label in generic:
        return ""
    # 剥离常见前缀(中国/全球)与单位后缀(市场规模/万元/亿美元...), 保留细分口径词
    cleaned = label
    for prefix in ("中国", "全球"):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):]
    cleaned = re.sub(r"(万元|亿元|亿美元|人民币|元|片个|个片|片|个)$", "", cleaned)
    for suffix in ("市场规模合计", "规模测算", "市场规模", "规模", "总量", "合计", "需求", "数量"):
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)]
            break
    cleaned = re.sub(r"(用|及|和|与)$", "", cleaned)
    if not cleaned or cleaned in generic or len(cleaned) > 14:
        return ""
    return cleaned


def _period(item):
    text = " ".join((*item.context, item.label))
    interim = re.search(
        r"\b(20\d{2})\s*(?:1\s*[-–—]\s*)?(\d{1,2})\s*[Mm]\b",
        text,
        re.IGNORECASE,
    )
    if interim:
        month = int(interim.group(2))
        if 1 <= month <= 12:
            return f"{interim.group(1)}_{month}m"
    annual = re.search(r"\b(20\d{2})\s*([AaEe])?\b", text, re.IGNORECASE)
    if not annual:
        return None
    period = f"{annual.group(1)}{(annual.group(2) or '').lower()}"
    # Multi-row headers (e.g. "2022A·东台") carry an entity/scenario dimension.
    # Keep it in the key so distinct columns do not collapse into one series key.
    # Entity is read from the header path context only — never from the row label.
    entity = _header_entity(" ".join(item.context))
    return f"{period}·{entity}" if entity else period


def _header_entity(text):
    """Extract a short entity/scenario token from a joined header path."""
    for token in re.split(r"[·|/]", text):
        token = token.strip()
        if re.fullmatch(r"20\d{2}[AaEe]?", token, re.IGNORECASE):
            continue
        if token and len(token) <= 8 and not re.search(r"\d", token):
            return token
    return ""


def _period_suffix(key):
    return bool(re.search(r"\.20\d{2}(?:[ae]|_\d{1,2}m)?$", key))


def _norm(value):
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return "".join(char for char in text if char.isalnum() or "\u4e00" <= char <= "\u9fff")


def _unit_compatible(family, unit):
    if not unit or family in {"none", "date"}:
        return True
    normalized = _norm(unit)
    if family == "percent":
        return "%" in str(unit) or any(marker in normalized for marker in ("percent", "pct", "百分比"))
    if family == "capacity":
        return any(marker in normalized for marker in ("片", "件", "台", "吨", "平方米", "mw", "gw", "capacity"))
    markers = {
        "currency": ("元", "rmb", "cny", "usd", "dollar", "million", "billion"),
        "multiple": ("x", "倍"),
        "price": ("元股", "pershare", "price"),
    }
    return any(_norm(marker) in normalized for marker in markers.get(family, ()))


def _source_preference(item):
    sheet = _norm(item.source_location.split("!", 1)[0])
    # 纯单元格 (模型输入事实) 优先于公式单元格 (模型推导结果); 公式推导值
    # 的重算覆盖由 merge_recalculated_values 完成, 此处不重复加分.
    score = 2 if not item.formula else 0
    if item.source_type == "docx":
        match = re.search(r"paragraph=(\d+)", item.source_location)
        if match and "entity=company" in item.source_location:
            paragraph = int(match.group(1))
            score += 15 if paragraph <= 10 else 10 if paragraph <= 50 else 5 if paragraph <= 200 else 0
        return score
    if sheet in {"returnsummary", "summary", "摘要", "汇总"}:
        score += 10
        match = re.search(r"![A-Z]+(\d+)$", item.source_location, re.IGNORECASE)
        if match and int(match.group(1)) <= 60:
            score += 3
        return score
    if any(marker in sheet for marker in ("summary", "汇总", "摘要")):
        return score + 5
    return score


def _generic_label(label):
    return label in {"margin", "rate", "ratio", "growth", "amount", "value", "比例", "增速", "增长率", "金额"}


def _candidate_usable(item):
    if not item.formula:
        return True
    # 公式候选可用性: 缓存值 (可能已被 LibreOffice 重算覆盖) 必须可用.
    cached = str(item.cached_value).upper()
    return item.cached_value is not None and cached not in {
        "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!"
    }


def _value_compatible(value_kind, value):
    if value_kind in {"text", "list", "person_list"}:
        if not isinstance(value, str) or not value.strip():
            return False
        normalized_value = _norm(value)
        if value_kind == "list" and normalized_value in {
            "营业收入", "合同金额万元", "行业", "简介", "股东背景", "收入万元",
            "客户", "公司", "企业", "厂商", "编号", "序号", "金额", "占比",
        }:
            return False
        if value_kind == "person_list" and (
            value.lstrip().startswith(("<<", ">>", "←", "→"))
            or any(marker in value for marker in ("受让", "股本", "出资金额", "持股比例"))
        ):
            return False
        return isinstance(value, str) and bool(value.strip())
    if value_kind in {"number", "series"}:
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    if value_kind == "date":
        return isinstance(value, str) and bool(re.search(r"(?:19|20)\d{2}", value))
    if value_kind == "table":
        return isinstance(value, (dict, list, tuple))
    return value is not None


def _source_compatible(field, item):
    sheet = _norm(item.source_location.split("!", 1)[0])
    is_competition = any(marker in sheet for marker in ("竞争格局", "竞争对手", "competitor", "competition"))
    if is_competition and field.key.startswith(("company.", "business.", "team.")):
        return False
    # A person list cannot carry a currency/capacity/ratio unit. Spreadsheet
    # summary matrices often repeat a management/ESOP row label across nearby
    # accounting cells; accepting those cells fabricates a management team.
    if field.value_kind == "person_list" and item.unit:
        return False
    label = _norm(item.label)
    if field.key == "financial.net_profit":
        # 净利润行的 label 必须是净利润本体; 所得税/利息费/少数股东损益/IRR/MOIC
        # 敏感性行不是净利润, 绑定会污染财务系列 (T8 违例).
        if not any(marker in label for marker in ("净利润", "netprofit", "netincome")):
            return False
    if field.key == "financial.revenue":
        # 营业收入行的 label 必须是收入本体; 营业成本/毛利/费用行不是收入.
        if not any(marker in label for marker in ("营业收入", "收入", "revenue", "sales")):
            return False
        if any(marker in label for marker in ("成本", "毛利", "费用", "净利")):
            return False
    if field.key == "financial.gross_margin":
        # 毛利率行的 label 必须是毛利率/利润率本体; EBITDA 利润率/净利率不是毛利率.
        if not any(marker in label for marker in ("毛利率", "利润率", "grossmargin", "grossmargin率")):
            return False
        if any(marker in label for marker in ("ebitda", "净利率", "净利润")):
            return False
    if field.key in {"deal.pre_money", "deal.post_money"}:
        # 投前/投后估值是标量交易事实; 退出估值/换股收购前估值不是本轮投前/投后,
        # 绑定会把回报假设污染进交易参数 (T8 违例). 检查 label + 列头 context.
        full = label + " " + " ".join(_norm(value) for value in item.context)
        marker = "投后" if field.key == "deal.post_money" else "投前"
        if marker not in full and "估值" not in full:
            return False
        if any(word in full for word in ("退出", "换股", "收购前", "投后估值考虑")):
            return False
    if field.key == "company.registered_capital":
        if "注册资本" not in label and "registeredcapital" not in label:
            return False
        if any(marker in label for marker in ("剩余", "受让", "对应", "转让")):
            return False
    if field.key == "team.founders" and not any(
        marker in label for marker in ("创始", "团队", "管理层", "founder", "managementteam")
    ):
        return False
    if field.key == "market.competitors" and label in {"公司", "企业", "厂商", "company"}:
        return is_competition
    if field.key == "market.growth":
        # 市场增速/CAGR 必须与市场/行业语义绑定; 财务行与产品行的增长率 (收入/
        # 净利/毛利 YOY, 产品线 CAGR) 不是市场规模增速 (T8 违例). 只有 label 直接
        # 表达市场/行业增速语义才接受.
        full = label + " " + " ".join(_norm(value) for value in item.context)
        if "市场" not in label and "行业" not in label and "规模" not in label:
            return False
        if any(marker in label for marker in ("营业收入", "净利", "毛利", "revenue", "净利率",
                                              "毛利率", "yoy", "成本", "费用", "税金", "营业利润",
                                              "利润总额", "所得税", "ebitda", "利息", "资产",
                                              "其他收益", "营业外")):
            return False
    if field.key == "business.capacity":
        # 产能是标量事实; 敏感性分析表的产能是不同情景假设, 不是单一产能.
        if any(marker in sheet for marker in ("sensitive", "敏感性")):
            return False
    if field.key == "business.customers":
        return any(marker in sheet for marker in ("客户", "customer", "订单"))
    if field.key == "business.products" and is_competition:
        return False
    return True
