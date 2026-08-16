"""Deterministic, source-bound investment analysis over a frozen fact snapshot."""
from dataclasses import dataclass
import math
import re


@dataclass(frozen=True)
class AnalysisClaim:
    block_key: str
    title: str
    body: str
    sources: tuple[str, ...]
    inference: str = "model_derived"
    topic: str = ""


@dataclass(frozen=True)
class AnalysisBundle:
    items: dict[str, tuple[AnalysisClaim, ...]]

    def get(self, block_key):
        return self.items.get(block_key, ())

    def metrics(self):
        claims = [claim for values in self.items.values() for claim in values]
        return {
            "generated": bool(claims),
            "block_count": len(self.items),
            "claim_count": len(claims),
            "body_characters": sum(len(claim.body) for claim in claims),
            "long_paragraphs": sum(len(claim.body) >= 100 for claim in claims),
            "all_claims_sourced": bool(claims) and all(claim.sources for claim in claims),
        }


def compose_model_analysis(schema, snapshot, workbook_profiles=()):
    """Create IC analysis without asserting facts absent from the frozen snapshot."""
    del schema  # The interface accepts the versioned schema; recipes use stable fact prefixes.
    items = {}

    def add(block, title, body, *keys, topic=""):
        facts = [snapshot.selected[key] for key in keys if key in snapshot.selected]
        if not facts:
            return
        claim = AnalysisClaim(
            block_key=block,
            title=title,
            body=body.strip(),
            sources=tuple(dict.fromkeys(
                f"{fact.source_uri}#{fact.source_location}" for fact in facts
            )),
            topic=topic,
        )
        items.setdefault(block, []).append(claim)

    revenue_a = _series(snapshot, "financial.revenue", "A")
    revenue_e = _series(snapshot, "financial.revenue", "E")
    gross_margin_a = _series(snapshot, "financial.gross_margin", "A")
    gross_margin_e = _series(snapshot, "financial.gross_margin", "E")
    net_profit_a = _series(snapshot, "financial.net_profit", "A")
    net_profit_e = _series(snapshot, "financial.net_profit", "E")
    cash_e = _series(snapshot, "financial.cash", "E")
    assets_e = _series(snapshot, "financial.total_assets", "E")
    debt_a = _series(snapshot, "financial.debt_ratio", "A")
    debt_e = _series(snapshot, "financial.debt_ratio", "E")
    ebitda_e = _series(snapshot, "financial.ebitda", "E")

    if revenue_a and revenue_e:
        actual_first, actual_last = revenue_a[0], revenue_a[-1]
        forecast_first, forecast_last = revenue_e[0], revenue_e[-1]
        forecast_cagr = _cagr(forecast_first[1], forecast_last[1], forecast_last[0] - forecast_first[0])
        add(
            "summary.narrative", "模型可见的经营轮廓",
            f"基于模型已绑定数据，历史营业收入由{_money(actual_first[1], actual_first[2])}增至"
            f"{_money(actual_last[1], actual_last[2])}，预测期则由{_money(forecast_first[1], forecast_first[2])}"
            f"增至{_money(forecast_last[1], forecast_last[2])}，对应预测复合增速约{_pct(forecast_cagr)}。"
            "这说明估值与回报并非建立在平稳延续情景上，而是显著依赖收入平台跨越式扩张。该结论只描述模型隐含路径，"
            "不能替代对订单、产能、价格和客户兑现节奏的业务尽调。",
            actual_first[3], actual_last[3], forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )
        add(
            "case.highlights", "预测收入进入加速扩张阶段",
            f"基于模型，{forecast_first[0]}E至{forecast_last[0]}E营业收入从"
            f"{_money(forecast_first[1], forecast_first[2])}增长至{_money(forecast_last[1], forecast_last[2])}，"
            f"复合增速约{_pct(forecast_cagr)}。若订单、产线爬坡与价格假设能够同步兑现，规模放大将为固定成本摊薄和利润释放提供基础；"
            "反之，任何一项关键假设延迟都会使后续年度形成更陡的追赶压力，因此该亮点应被理解为需要验证的模型上行路径。",
            forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )
        add(
            "case.risks", "高增长路径的执行风险",
            f"基于模型，预测期营业收入需要从{_money(forecast_first[1], forecast_first[2])}提升至"
            f"{_money(forecast_last[1], forecast_last[2])}，累计扩张约{_multiple(forecast_last[1] / forecast_first[1])}。"
            "模型没有同时提供可核验的订单覆盖、客户集中度、产能利用率及价格桥接，因此收入增长目前仍是财务假设而非已验证经营事实。"
            "投资决策应把年度订单锁定率、产能投放节点和回款条件设为分期交割或否决条件。",
            forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )
        add(
            "industry.narrative", "模型隐含的需求扩张假设",
            f"基于模型，预测收入在{forecast_first[0]}E至{forecast_last[0]}E期间保持高增，说明底层估值逻辑隐含了"
            "赛道需求扩张、份额提升或并购整合中的至少一种驱动。但当前输入没有提供外部市场规模、行业增速和竞争份额证据，"
            f"因此不能把模型中的{_pct(forecast_cagr)}复合增速直接表述为行业增长。后续应将公司增长、行业增长和并表因素拆分验证。",
            forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )
        company = snapshot.selected.get("company.name")
        if _available(company):
            add(
                "company.basic", "实名主体与模型经营规模已建立证据连接",
                f"公开证据已将本项目主体绑定为{company.value}；模型显示其历史期形成约"
                f"{_money(actual_last[1], actual_last[2])}收入规模，并在预测期被假设扩张至"
                f"{_money(forecast_last[1], forecast_last[2])}。主体名称的绑定消除了匿名项目识别缺口，"
                "但业务品类、客户、团队和市场地位仍必须分别依据各自来源判断，不能从财务曲线反推。",
                "company.name", actual_last[3], forecast_last[3],
                topic="company_scale",
            )
        else:
            add(
                "company.basic", "模型呈现的业务规模而非身份结论",
                f"基于模型，标的在历史期形成约{_money(actual_last[1], actual_last[2])}收入规模，并在预测期被假设扩张至"
                f"{_money(forecast_last[1], forecast_last[2])}。这些数字可以支持对经营体量和资本需求的分析，但不能反推出具体主体、"
                "业务品类或市场地位。正文因此采用“Project GL（匿名标的）”作为项目标识，所有主体身份与组织信息继续列入尽调缺口。",
                actual_last[3], forecast_last[3],
                topic="company_scale",
            )
        add(
            "financial.forecast", "预测收入增长的斜率与可兑现性",
            f"基于模型，收入由{forecast_first[0]}E的{_money(forecast_first[1], forecast_first[2])}增至"
            f"{forecast_last[0]}E的{_money(forecast_last[1], forecast_last[2])}，复合增速约{_pct(forecast_cagr)}。"
            "该路径要求商业需求、供应能力和营运资金同步扩张，不能仅用终值年度解释。投委会应逐年核对收入桥，包括存量业务自然增长、"
            "新增产能、并购并表、单价变化和销量贡献，并把未有订单或合同支撑的部分单独折价。",
            forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )
        add(
            "financial.forecast", "预测期跨越历史规模的幅度",
            f"基于模型，{forecast_first[0]}E收入已达到最近历史年度的"
            f"{_multiple(forecast_first[1] / actual_last[1])}，至{forecast_last[0]}E进一步达到"
            f"{_multiple(forecast_last[1] / actual_last[1])}。如此明显的规模跃迁通常意味着模型包含新产线、并购整合或商业化放量等结构性事件。"
            "在事件性质和执行条件尚未由业务材料解释前，预测不能按普通有机增长给予同等置信权重。",
            actual_last[3], forecast_first[3], forecast_last[3],
            topic="revenue_growth",
        )

    if net_profit_a and net_profit_e:
        hist_last, forecast_first, forecast_last = net_profit_a[-1], net_profit_e[0], net_profit_e[-1]
        add(
            "summary.narrative", "盈利拐点来自预测而非历史验证",
            f"基于模型，最近历史年度净利润为{_money(hist_last[1], hist_last[2])}，"
            f"而{forecast_first[0]}E转为{_money(forecast_first[1], forecast_first[2])}，至{forecast_last[0]}E增至"
            f"{_money(forecast_last[1], forecast_last[2])}。该变化构成投资逻辑中的核心盈利拐点，但其性质仍是预测。"
            "应通过毛利改善、费用率下降、并购协同及非经常性损益四条桥逐项复核，避免把会计口径变化误读为经营质量提升。",
            hist_last[3], forecast_first[3], forecast_last[3],
            topic="profit_turnaround",
        )
        add(
            "case.highlights", "模型显示由亏损进入持续盈利",
            f"基于模型，净利润从历史期的{_money(hist_last[1], hist_last[2])}切换到{forecast_first[0]}E的"
            f"{_money(forecast_first[1], forecast_first[2])}，并在{forecast_last[0]}E达到"
            f"{_money(forecast_last[1], forecast_last[2])}。若盈利改善主要来自可持续毛利和经营杠杆，而非一次性项目，"
            "这一拐点将显著改善退出估值基础；验证重点是利润与现金流的一致性以及费用投入是否被低估。",
            hist_last[3], forecast_first[3], forecast_last[3],
            topic="profit_turnaround",
        )
        add(
            "financial.historical", "历史盈利能力尚未形成稳定基线",
            f"基于模型，最近历史年度净利润仍为{_money(hist_last[1], hist_last[2])}，说明预测期的正向利润并非简单延续历史趋势。"
            "历史亏损意味着成本结构、资产效率或业务组合至少有一项尚未达到稳定状态。投委会需要取得经审计报表、分业务利润表和异常项目明细，"
            "以确认亏损来源、可逆程度及改善所需现金投入。",
            hist_last[3],
            topic="profit_turnaround",
        )
        add(
            "financial.forecast", "净利润释放速度高于历史可见基础",
            f"基于模型，净利润在{forecast_first[0]}E达到{_money(forecast_first[1], forecast_first[2])}，"
            f"在{forecast_last[0]}E进一步达到{_money(forecast_last[1], forecast_last[2])}。盈利持续增长对收入兑现、毛利水平和费用控制均有要求。"
            "建议使用收入、毛利率和费用率三因素拆解净利润，并设置低于预算时的资金拨付、估值调整或治理权触发机制。",
            forecast_first[3], forecast_last[3],
            topic="profit_turnaround",
        )

    if gross_margin_a and gross_margin_e:
        hist_last, peak, forecast_last = gross_margin_a[-1], max(gross_margin_e, key=lambda item: item[1]), gross_margin_e[-1]
        add(
            "case.highlights", "预测毛利率较历史低点显著修复",
            f"基于模型，毛利率由最近历史年度的{_pct(hist_last[1])}修复至预测期高点{_pct(peak[1])}。"
            "如果修复来自产品结构升级、产能利用率提升或采购成本改善，盈利弹性可能高于收入增速；但当前模型未说明改善来源，"
            "因此应通过单价、单位成本、良率和产品组合的量价桥确认其可持续性。",
            hist_last[3], peak[3],
            topic="gross_margin",
        )
        add(
            "case.risks", "预测后段毛利率回落",
            f"基于模型，毛利率在{peak[0]}E达到{_pct(peak[1])}后，于{forecast_last[0]}E回落至{_pct(forecast_last[1])}，"
            f"下降约{_percentage_points(peak[1] - forecast_last[1])}。这表明模型自身已经包含竞争、价格、爬坡成本或低毛利业务占比上升的压力。"
            "若收入增速同时放缓，利润端可能承受双重挤压，需对终值年度毛利率和估值倍数进行联动压力测试。",
            peak[3], forecast_last[3],
            topic="gross_margin",
        )
        add(
            "financial.historical", "历史毛利率波动提示经营质量风险",
            f"基于模型，最近历史年度毛利率为{_pct(hist_last[1])}，与预测期高点{_pct(peak[1])}存在显著差距。"
            "如此幅度的变化不能仅用规模效应解释，需要核对收入确认、存货跌价、停工损失、并购会计及成本归集口径。"
            "在取得按产品和客户拆分的毛利资料前，模型中的利润改善应保持折价而不是直接资本化。",
            hist_last[3], peak[3],
            topic="gross_margin",
        )
        add(
            "financial.forecast", "毛利率先升后降决定利润质量",
            f"基于模型，预测毛利率在{peak[0]}E达到{_pct(peak[1])}，随后在{forecast_last[0]}E降至{_pct(forecast_last[1])}。"
            "这一路径意味着利润增长更多依赖收入规模，而非持续扩张的单位经济性。投资分析应把毛利高点视为需要证据支持的阶段性假设，"
            "并在估值中采用中枢毛利率而不是峰值毛利率。",
            peak[3], forecast_last[3],
            topic="gross_margin",
        )

    if cash_e:
        low, high, last = min(cash_e, key=lambda item: item[1]), max(cash_e, key=lambda item: item[1]), cash_e[-1]
        add(
            "case.risks", "现金余额大幅波动",
            f"基于模型，预测现金余额在{high[0]}E达到{_money(high[1], high[2])}，但在{low[0]}E降至"
            f"{_money(low[1], low[2])}，期末为{_money(last[1], last[2])}。波动幅度说明模型中可能存在投资、并购、偿债或营运资金集中占用。"
            "在现金流量表和资金用途未完成逐项勾稽前，账面盈利不能等同于可分配现金，流动性安全垫应作为交割条件。",
            high[3], low[3], last[3],
            topic="cash_volatility",
        )
        add(
            "company.operations", "资金投入节奏决定运营兑现",
            f"基于模型，预测现金从高点{_money(high[1], high[2])}下降至低点{_money(low[1], low[2])}，"
            "说明经营扩张伴随显著资金占用。该现象可能来自固定资产建设、存货和应收增加或交易性支出，但模型数字本身不能区分原因。"
            "应取得资本开支计划、营运资金周转和资金用途清单，以判断投入能否在预测收入兑现前保持充足流动性。",
            high[3], low[3],
            topic="cash_volatility",
        )
        add(
            "financial.forecast", "现金利润转化需要单独验证",
            f"基于模型，预测现金余额最低降至{_money(low[1], low[2])}，而期末回升至{_money(last[1], last[2])}。"
            "现金路径与利润增长并非自然同步，可能受到资本开支、并购付款、债务安排和营运资金的共同影响。"
            "建议以经营现金流、自由现金流和最低现金余额三项指标重新设置财务契约，避免只看净利润完成率。",
            low[3], last[3],
            topic="cash_volatility",
        )

    if assets_e:
        first, last = assets_e[0], assets_e[-1]
        add(
            "case.highlights", "资产平台同步扩张提供收入承载基础",
            f"基于模型，总资产由{first[0]}E的{_money(first[1], first[2])}增至{last[0]}E的"
            f"{_money(last[1], last[2])}，扩张至{_multiple(last[1] / first[1])}。如果新增资产对应可投产产能、可回收营运资金或可产生现金流的并购资产，"
            "其增长可以支撑收入放量；若主要由低效资产或高商誉构成，则规模增长反而会降低资本回报。",
            first[3], last[3],
            topic="asset_expansion",
        )
        add(
            "company.operations", "资产扩张需要与产能和回报率勾稽",
            f"基于模型，预测总资产在{first[0]}E至{last[0]}E期间增加"
            f"{_money(last[1] - first[1], last[2])}。资产增长明显快于静态经营基线，意味着运营体系需要承接建设、整合和爬坡任务。"
            "尽调应把新增资产拆分为固定资产、营运资金、商誉和其他项目，并计算投产时间、周转效率及投入资本回报率。",
            first[3], last[3],
            topic="asset_expansion",
        )
        add(
            "financial.forecast", "资产负担与收入增长应同步评估",
            f"基于模型，总资产从{_money(first[1], first[2])}增至{_money(last[1], last[2])}。"
            "这一扩张既可能是增长必要条件，也可能带来折旧、融资和整合成本。评价预测合理性时，不能只看收入和利润增速，"
            "还应比较资产周转率、资本开支回收期以及新增资产对自由现金流的影响。",
            first[3], last[3],
            topic="asset_expansion",
        )
        revenue_by_year = {item[0]: item for item in revenue_e}
        asset_by_year = {item[0]: item for item in assets_e}
        common_years = sorted(set(revenue_by_year) & set(asset_by_year))
        if common_years:
            first_year, last_year = common_years[0], common_years[-1]
            first_revenue, last_revenue = revenue_by_year[first_year], revenue_by_year[last_year]
            first_assets, last_assets = asset_by_year[first_year], asset_by_year[last_year]
            add(
                "company.operations", "资产周转改善是扩张回报的必要条件",
                f"基于模型，营业收入与总资产的简单比值由{first_year}E的"
                f"{_multiple(first_revenue[1] / first_assets[1])}提升至{last_year}E的"
                f"{_multiple(last_revenue[1] / last_assets[1])}。这意味着模型不仅要求资产规模扩大，还要求新增资产逐步形成更高收入产出。"
                "若建设、整合或客户导入延迟，资产周转改善将落空，折旧、融资和管理成本却可能先行发生；应按项目核对投产节点、收入贡献和爬坡达成率。",
                first_revenue[3], last_revenue[3], first_assets[3], last_assets[3],
                topic="asset_expansion",
            )
        if cash_e:
            cash_low = min(cash_e, key=lambda item: item[1])
            add(
                "company.operations", "资产扩张与现金安全垫需要联动管理",
                f"基于模型，总资产由{_money(first[1], first[2])}增至{_money(last[1], last[2])}，"
                f"同期预测现金余额一度降至{_money(cash_low[1], cash_low[2])}。这组路径表明增长计划可能在形成回报前先消耗流动性。"
                "运营预算应把资本开支、营运资金和债务偿还按月排布，并设置最低现金余额、备用融资额度和投资节奏调整机制，"
                "防止单纯追求收入目标导致资金链承压。",
                first[3], last[3], cash_low[3],
                topic="asset_expansion",
            )

    if debt_a and debt_e:
        hist_last, low, last = debt_a[-1], min(debt_e, key=lambda item: item[1]), debt_e[-1]
        add(
            "financial.historical", "历史杠杆水平偏高",
            f"基于模型，最近历史年度资产负债率为{_pct(hist_last[1])}。较高杠杆会放大盈利波动对偿债能力和再融资的影响，"
            "也使预测期的资金投入更依赖资本结构安排。应核对有息负债、到期结构、担保抵押、财务费用和契约限制，"
            "并确认模型中的负债率口径是否包含租赁负债及并购对价安排。",
            hist_last[3],
            topic="leverage",
        )
        add(
            "case.risks", "去杠杆后重新上升的资本结构风险",
            f"基于模型，资产负债率从历史期{_pct(hist_last[1])}降至预测低点{_pct(low[1])}，但在"
            f"{last[0]}E回升至{_pct(last[1])}。这说明资本结构改善并非单向趋势，后续扩张可能重新消耗资产负债表空间。"
            "应对新增债务、利率和偿债现金流进行压力测试，并设置最高杠杆与最低利息覆盖倍数。",
            hist_last[3], low[3], last[3],
            topic="leverage",
        )
        add(
            "financial.forecast", "资本结构改善存在阶段性",
            f"基于模型，预测资产负债率最低降至{_pct(low[1])}，期末回升至{_pct(last[1])}。"
            "低点可能受本轮股权资金或资产重估影响，不能直接理解为经营现金流驱动的永久去杠杆。"
            "投资人应把资金注入、债务偿还和新增融资拆开，确认期末杠杆回升是否与扩张回报相匹配。",
            low[3], last[3],
            topic="leverage",
        )

    if ebitda_e:
        first, last = ebitda_e[0], ebitda_e[-1]
        add(
            "financial.forecast", "EBITDA 增长提供估值承接但不等于现金流",
            f"基于模型，EBITDA由{first[0]}E的{_money(first[1], first[2])}增至{last[0]}E的"
            f"{_money(last[1], last[2])}。该指标可用于观察经营杠杆和偿债覆盖，但未扣除资本开支、营运资金和税费。"
            "若退出估值以EBITDA或利润倍数为基础，应同时验证调整项的一致性，避免把一次性收益或资本化费用纳入持续经营能力。",
            first[3], last[3],
            topic="ebitda",
        )

    _compose_transaction_claims(snapshot, add)
    _compose_return_claims(snapshot, add)
    _compose_scope_claims(snapshot, workbook_profiles, add)
    return AnalysisBundle(_dedupe_topics(items))


def _dedupe_topics(items):
    """Collapse duplicate analysis claims across blocks.

    The same fact topic (IRR/MOIC, pre/post money, exit PE, cash path, leverage,
    stake reconciliation) is often composed into several blocks with near-identical
    wording. Keep only the highest-priority occurrence per topic so the memo does
    not repeat the same number-and-judgement paragraph 3-4 times. Priority favors
    the most decision-relevant block for that topic.
    """
    priority = {
        "summary.narrative": 0,
        "transaction.terms": 1,
        "transaction.ownership": 2,
        "case.highlights": 3,
        "case.risks": 4,
        "financial.returns": 5,
        "financial.sensitivity": 6,
        "financial.forecast": 7,
        "industry.narrative": 8,
        "company.basic": 9,
    }
    best = {}
    for block, claims in items.items():
        for claim in claims:
            if not claim.topic:
                continue
            key = claim.topic
            cur = best.get(key)
            if cur is None or priority.get(claim.block_key, 99) < priority.get(cur.block_key, 99):
                best[key] = claim
    dropped = {claim for claim in (
        c for claims in items.values() for c in claims
    ) if claim.topic and best.get(claim.topic) is not claim}
    return {
        block: tuple(claim for claim in claims if claim not in dropped)
        for block, claims in items.items()
    }


def _compose_transaction_claims(snapshot, add):
    facts = snapshot.selected
    pre = facts.get("deal.pre_money")
    post = facts.get("deal.post_money")
    investment = facts.get("deal.equity_investment")
    stake = facts.get("deal.project_stake")
    if pre and post and investment:
        add(
            "transaction.terms", "交易估值与资金规模",
            f"基于模型，本轮投前估值为{_money(pre.value, pre.unit)}，投后估值为{_money(post.value, post.unit)}，"
            f"股权投资额为{_money(investment.value, investment.unit)}。投前、投资额与投后在数值上形成完整资金桥，"
            "但该桥仅说明模型口径自洽，仍需以正式交易文件核对币种、支付安排、老股与增资拆分、交割条件及估值调整机制。",
            "deal.pre_money", "deal.post_money", "deal.equity_investment",
            topic="deal_valuation",
        )
        identity_clause = (
            "主体身份已由公开证据绑定，但正式交易条款文件尚需补齐。"
            if _available(snapshot.selected.get("company.name")) else
            "主体身份和条款文件尚未补齐。"
        )
        add(
            "summary.narrative", "交易结构决定模型回报的可实现性",
            f"基于模型，投资额{_money(investment.value, investment.unit)}对应投后估值"
            f"{_money(post.value, post.unit)}。回报测算建立在这一进入成本、后续稀释和退出估值共同成立的基础上。"
            f"{identity_clause}本文只能给出条件性判断，不能把模型回报视为可签署的交易结论。",
            "deal.equity_investment", "deal.post_money",
            topic="deal_valuation",
        )
    if post and investment and stake:
        implied = float(investment.value) / float(post.value)
        reported = _as_fraction(stake.value)
        add(
            "transaction.ownership", "投资额与持股口径需要勾稽",
            f"基于模型，投资额除以投后估值得到的简单比例约为{_pct(implied)}，而模型另列项目持股比例为{_pct(reported)}。"
            "两者差异可能来自老股、分层载体、稀释、部分资金不计入股权或字段口径不同。该差异不应由引擎自行选择一个数字，"
            "必须通过资本化表、资金路径和正式条款逐项解释。",
            "deal.equity_investment", "deal.post_money", "deal.project_stake",
            topic="stake_reconcile",
        )
        add(
            "case.risks", "持股比例与资金桥存在口径差异",
            f"基于模型，简单投后比例约{_pct(implied)}，但单列持股比例为{_pct(reported)}。如果差异不能由交易结构解释，"
            "则IRR、MOIC和退出所得可能使用了不同分母。投委会应把资本化表完全摊薄口径、载体层级、老股比例和后续稀释列为签署前必核事项。",
            "deal.equity_investment", "deal.post_money", "deal.project_stake",
            topic="stake_reconcile",
        )


def _compose_return_claims(snapshot, add):
    facts = snapshot.selected
    irr = facts.get("return.irr")
    moic = facts.get("return.moic")
    exit_pe = facts.get("return.exit_pe")
    valuations = _series(snapshot, "return.exit_valuation", "E")
    if irr and moic:
        add(
            "summary.narrative", "模型回报达到积极区间但仍是条件性结果",
            f"基于模型，项目IRR为{_pct(irr.value)}、MOIC为{_multiple(moic.value)}。两项指标在模型层面具有吸引力，"
            "但其实现依赖进入估值、持股口径、盈利预测、退出时点与退出倍数同时成立。由于这些关键条件尚未全部获得外部证据，"
            "建议结论为“谨慎推进并补充关键尽调”，而不是直接形成无条件投资建议。",
            "return.irr", "return.moic",
            topic="return_profile",
        )
        add(
            "case.highlights", "模型回报具备投资吸引力",
            f"基于模型，IRR为{_pct(irr.value)}、MOIC为{_multiple(moic.value)}，说明在基准情景下投资收益能够覆盖"
            "较长持有期和经营改善的不确定性。该亮点的有效性取决于退出所得是否使用完全摊薄持股、退出估值是否与利润口径一致，"
            "以及中间融资是否造成额外稀释，需与资本化表和回报现金流逐项勾稽。",
            "return.irr", "return.moic",
            topic="return_profile",
        )
        add(
            "financial.returns", "基准情景回报",
            f"基于模型，基准IRR为{_pct(irr.value)}，投资倍数为{_multiple(moic.value)}。IRR反映时间价值，MOIC反映绝对回收倍数，"
            "二者组合比单一指标更适合判断退出质量。当前结果应作为情景输出而非承诺值，任何退出延迟、利润不达标、倍数下调或稀释增加都会同时压低两项指标。",
            "return.irr", "return.moic",
            topic="return_profile",
        )
        add(
            "financial.sensitivity", "回报指标的联合敏感性",
            f"基于模型，{_pct(irr.value)} IRR与{_multiple(moic.value)} MOIC由同一组进入和退出现金流产生。"
            "压力测试不应只改变退出倍数，还应同时考虑利润完成率、退出延迟和持股稀释。建议至少设置收入下调、毛利下调、退出倍数下调及退出延期四类情景，"
            "并观察IRR和MOIC是否仍满足基金底线。",
            "return.irr", "return.moic",
            topic="return_profile",
        )
    if exit_pe:
        add(
            "case.risks", "退出倍数假设偏强",
            f"基于模型，退出市盈率采用{_multiple(exit_pe.value)}。较高退出倍数意味着回报不仅依赖利润兑现，也依赖资本市场在退出时继续给予成长估值。"
            "如果终值年度增长放缓、毛利率下降或市场风险偏好收缩，估值倍数与利润可能同时承压。应以可比交易、上市公司前瞻倍数和下行情景验证终值。",
            "return.exit_pe",
            topic="exit_pe",
        )
        add(
            "financial.returns", "退出倍数是终值的核心杠杆",
            f"基于模型，退出估值使用{_multiple(exit_pe.value)}市盈率。该参数对终值具有线性影响，并会放大终值利润误差。"
            "投委会应区分标的成长溢价、流动性折价和退出路径差异，在没有可靠可比证据时，以更低倍数作为基准并将高倍数仅保留为上行情景。",
            "return.exit_pe",
            topic="exit_pe",
        )
        add(
            "financial.sensitivity", "退出倍数下调应作为第一压力情景",
            f"基于模型，基准退出倍数为{_multiple(exit_pe.value)}。建议至少测试倍数下调20%、30%和40%的情景，"
            "并与利润完成率同步变化，而不是保持其他条件不变。只有在倍数回归行业中枢后仍能达到基金要求，基准回报才具有足够安全边际。",
            "return.exit_pe",
            topic="exit_pe",
        )
    if valuations:
        first, last = valuations[0], valuations[-1]
        add(
            "financial.returns", "退出估值随盈利年度推进",
            f"基于模型，退出估值从{first[0]}E的{_money(first[1], first[2])}上升至{last[0]}E的"
            f"{_money(last[1], last[2])}。终值增长说明模型把后续盈利释放计入退出价值，但更晚退出也会降低IRR。"
            "应把不同退出年度的估值、持股比例、现金回收和持有期限放在同一张回报桥中，避免只比较名义终值。",
            first[3], last[3],
            topic="exit_valuation",
        )


def _compose_scope_claims(snapshot, workbook_profiles, add):
    anchors = tuple(sorted(snapshot.selected))[:4]
    if not anchors:
        return
    if not _available(snapshot.selected.get("market.competitors")):
        add(
            "industry.competition", "竞争信息尚未进入模型证据链",
            "基于模型可确认的内容主要是财务、交易和回报参数，现有冻结事实没有提供可独立识别的市场份额、主要竞争者、替代技术和价格比较。"
            "因此本文不生成竞争排名或领先地位判断。后续应补充可比企业、产品参数、客户选择标准和历史中标数据，再把竞争结论纳入估值与风险分析。",
            *anchors,
        )
    if not _available(snapshot.selected.get("team.founders")):
        add(
            "company.team", "治理与执行能力需要专项尽调",
            "基于模型能够看到高增长、盈利修复和资产扩张目标，但这些目标本身不能证明管理与执行能力。现有冻结事实没有形成可核验的治理结构、"
            "核心岗位履历、激励安排和关键人员稳定性证据。投委会应要求管理层访谈、组织架构、关键岗位分工和历史预算达成率，以验证模型是否具备执行主体。",
            *anchors,
        )
    if not (_available(snapshot.selected.get("business.products")) and _available(snapshot.selected.get("business.customers"))):
        add(
            "company.products", "收入预测需要产品与订单桥",
            "基于模型的收入预测尚未由产品组合、销量、单价、客户和订单数据解释。本文不会从财务曲线反推具体业务内容。"
            "后续材料应把每条收入曲线拆成产品、区域、渠道与客户贡献，并标注在手订单、框架协议、预测订单和管理层目标，才能判断增长的证据等级。",
            *anchors,
        )
    source_count = len({profile.source_uri for profile in workbook_profiles}) or len({
        fact.source_uri for fact in snapshot.selected.values()
    })
    company = snapshot.selected.get("company.name")
    title = "实名项目的模型证据边界" if _available(company) else "匿名模型的证据边界"
    identity = f"主体已绑定为{company.value}，" if _available(company) else "主体证照尚未绑定，"
    add(
        "company.basic", title,
        f"基于模型，本次分析使用{source_count}个已提供数据源形成冻结事实，{identity}能够支持财务趋势、交易结构和回报情景判断，"
        "但不能替代股权沿革、业务合同和团队材料。所有定性结论均以“模型显示”或“基于模型推断”表达，"
        "未披露信息集中进入尽调缺口，不因版面完整而被自动补写。",
        *(("company.name",) if _available(company) else ()), *anchors,
    )
    add(
        "industry.narrative", "行业验证应独立于公司预测",
        "基于模型得到的增长与利润路径属于公司层预测，不能直接代表行业规模或行业增速。行业验证应另行取得权威统计、上下游产量、"
        "渗透率、价格趋势和竞争供给，并将行业增长与份额提升分开。只有外部需求与内部份额两条证据链同时成立，模型的高增长假设才具备可投资性。",
        *anchors,
    )


def _series(snapshot, base, suffix):
    values = []
    # 兼容 2.5.1 引入的 row-breakdown 后缀 (financial.revenue.2021e·营业收入)
    # 与旧版纯 key (financial.revenue.2021e) 两种形态。优先裸 key；只有带后缀
    # 变体时取该时期的第一个匹配，避免同一时期产生重复行。
    bare = re.compile(rf"^{re.escape(base)}\.(20\d{{2}})([aAeE])$")
    suffixed = re.compile(rf"^{re.escape(base)}\.(20\d{{2}})([aAeE])(?:·.+)?$")
    for key, fact in snapshot.selected.items():
        if not isinstance(fact.value, (int, float)) or isinstance(fact.value, bool):
            continue
        m = bare.match(key)
        if m and m.group(2).upper() == suffix:
            values.append((int(m.group(1)), float(fact.value), fact.unit, key))
    if values:
        return sorted(values)
    seen = set()
    for key, fact in snapshot.selected.items():
        if not isinstance(fact.value, (int, float)) or isinstance(fact.value, bool):
            continue
        m = suffixed.match(key)
        if not m or m.group(2).upper() != suffix:
            continue
        period = int(m.group(1))
        if period in seen:
            continue
        seen.add(period)
        values.append((period, float(fact.value), fact.unit, key))
    return sorted(values)


def _available(fact):
    return fact is not None and str(getattr(fact.status, "value", fact.status)) not in {"missing", "conflict"}


def _cagr(first, last, years):
    if years <= 0 or first <= 0 or last <= 0:
        return 0.0
    return (last / first) ** (1 / years) - 1


def _as_fraction(value):
    value = float(value)
    return value / 100 if abs(value) > 1.5 else value


def _pct(value):
    return f"{_as_fraction(value) * 100:.2f}%"


def _percentage_points(value):
    return f"{abs(_as_fraction(value) * 100):.2f}个百分点"


def _multiple(value):
    return f"{float(value):.2f}x"


def _money(value, unit):
    number = float(value)
    text = str(unit or "model unit")
    if "百万元" in text:
        return f"{number / 100:,.2f}亿元"
    if text == "万元" or "万元人民币" in text:
        return f"{number / 10000:,.2f}亿元"
    if "亿元" in text:
        return f"{number:,.2f}亿元"
    if math.isfinite(number):
        return f"{number:,.2f}{text}"
    return f"{value}{text}"
