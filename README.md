# ni-memo — 可追溯投资备忘录生成引擎

从 XLSX 财务模型、历史 DOCX、PDF、CSV、JSON 等混合输入，生成**源可追溯**的 IC（投资委员会）备忘录 DOCX。每个事实都绑定到具体来源（单元格/段落/URL），绝不编造。

> 本包是 WorkBuddy 环境下的 skill 目录（`~/.workbuddy/skills/ni-memo/`），可独立复制到任意 Codex/Claude 环境使用。核心是 `ni_memo` Python 包，不依赖 WorkBuddy 运行时。

---

## 环境要求

- Python ≥ 3.10（在 3.13 实测通过）
- Windows / macOS / Linux 均可（LibreOffice 字段更新仅 Windows 自动启用）

## 安装

```bash
# 1. 进入 skill 目录
cd ni-memo

# 2. 安装运行时依赖
pip install -r requirements.txt

# 3. 验证安装（可选）
python -m ni_memo.cli --help
```

> **可选依赖**：LibreOffice（含 `uno`）用于更新 DOCX 的 TOC/页码域。不装也能运行——引擎自动降级为 PowerShell COM（Windows）或跳过字段更新，不影响 memo 内容生成。

## 一条命令跑通

```powershell
python -m ni_memo.cli `
  --inputs "D:\path\model.xlsx" "D:\path\historical-memo.docx" `
  --out "D:\Codex\Outputs\ni-memo\project-name" `
  --work-dir "D:\WorkBuddy\NeoStar\一级市场机会\_runs\ni-memo\project-name-YYYYMMDD" `
  --narrative "D:\path\narrative.json" `
  --truth-overrides "D:\path\truth_overrides.json" `
  --project-as-of "2021-09-30" `
  --snapshot-id "project-name-YYYYMMDD"
```

- `--inputs`：任意组合的 XLSX / DOCX / PDF / CSV / JSON 源文件（不要求有 `Database` sheet）
- `--narrative`：投资叙事素材 JSON（补充模型无法表达的行业/业务/团队故事）
- `--truth-overrides`：调用方确认的真值 JSON（如投前估值、交易条款）
- `--formal-template`：正式模板 DOCX，校验生成 memo 与模板的版式一致性

**产出**：`--out` 目录下唯一的用户交付物 `memo.docx`（七章节：项目概要/交易条款/投资亮点与风险/行业概况/公司概况/财务情况/附录），含封面、可更新的 TOC、页眉页脚、页码。

## 验收证据（--work-dir 内）

| 文件 | 内容 |
|---|---|
| `facts.json` | 冻结的事实快照与每个候选，含精确来源 |
| `mapping.json` | 字段到来源的绑定规则与证据 |
| `pending.md` | 缺失事实、冲突、公式门禁 |
| `acceptance_report.json` | 内容/完整度/公式/复算/模板/视觉验收 |
| `run_summary.json` | 顶层结果，含 grade |
| `visual/` | PDF 与页面 PNG 渲染证据（QA 用，非交付物） |

## 质量等级

- `PASS`：所有必需事实完整 + 100% 独立验证/佐证覆盖
- `PASS_WITH_NOTES`：可用但不可签署（可能有缺失事实或公式警告）
- `FAIL`：仍有可用草稿，但存在硬性缺口

引擎**永远不会**把缺失的定性字段编造成"听起来像事实"的叙述——缺失就明确列为 `pending`。

## 常见问题

**Q: 没有 LibreOffice 会怎样？**
A: 不崩。字段更新降级为跳过/PowerShell，memo 内容照常生成。

**Q: 我的模型文件没有 `Database` sheet？**
A: 不需要。引擎自动语义发现（`discover.py`），不依赖固定 sheet 名。

**Q: 金额显示为放大 10 倍？**
A: 若遇到，检查 `ni_memo/analysis.py` 的 `_money` 函数中 `百万元 → 亿元` 换算应为 `/100`（非 `/10`）。此 bug 已修复并有回归测试锁定。

## 开发与测试

```bash
pip install pytest
python -m pytest tests/ -q
```

测试覆盖四个关键防线：
- `_money` 单位换算（防 10 倍回归）
- `_dedupe_topics` 跨区块去重
- `_series` 序列正则后缀兼容
- 硬编码绝对路径扫描

## 目录结构

```
ni-memo/
├── SKILL.md                    # skill 主文档（产品契约/工作流/命令）
├── requirements.txt            # 运行时依赖
├── README.md                   # 本文件
├── schema/
│   └── standard_ic_memo.json   # 版式契约（7章节/语义槽/别名/要求）
├── ni_memo/
│   ├── cli.py                  # 命令行入口（组装产物）
│   ├── ingest.py               # 源无关的稀疏抽取
│   ├── discover.py             # 确定性语义绑定
│   ├── reconcile.py            # 来源/日期感知选择与冲突保留
│   ├── analysis.py             # 模型驱动投资分析（含去重）
│   ├── content.py              # schema + 冻结事实 → 类型化内容块
│   ├── render.py               # 干净 DOCX 渲染（A4/仿宋/TOC/页脚）
│   ├── formula_audit.py        # 公式审计与缓存门禁
│   ├── recalculate.py          # 非破坏式 LibreOffice 复算
│   ├── chart.py                # 图表 PNG 渲染
│   ├── narratives.py           # 叙事素材加载
│   ├── fields.py               # TOC/页码域更新（LibreOffice/PowerShell）
│   ├── visual.py               # PDF/页面 PNG/联系表视觉门禁
│   └── ...                     # 其余支撑模块
└── tests/                      # 回归测试
```
