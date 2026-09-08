# resume-stats-table-skill

校招简历信息统计表生成 **agent 技能（skill）**。批量读取应届生简历（PDF），按「校招人员信息统计表」固定模板整理成 Excel，支持新一届建表与后续增量追加。

适用于 DSH / Claude Code 等支持 skill 的 agent 工具。脚本**只用 Python 标准库**，Windows / Linux / macOS 通用。

> ⚠️ 本仓库**不含**内部招聘模板（如 `固定模板.xlsx`，含具体表头）。使用时请按你单位自己的模板调整列映射（见 `skills/resume-stats-table/references/sheet_columns.md`）。

---

## 功能

- **新一届建表**：以固定模板为底，保留所有 sheet 的表头/格式/合并单元格，清空人员数据，填入本批候选人。
- **按岗位地点重命名**：主 sheet → `{地点}岗`，面试名单 sheet → `{地点}面试名单`（如深圳）。
- **面试安排**：按用户时间填「面试安排」表头 D2/E2；不确定则填「待定」。
- **增量追加**：后续新增简历**只追加**到主表末尾，按姓名/联系方式去重、序号自动续号、只动主表、**另存新副本、绝不覆盖原文件**。
- **保留静态字典**：如「211&985院校名单」等 sheet 原样保留。

---

## 安装（skill）

把 `skills/resume-stats-table/`（或可下载的 `resume-stats-table.skill` 分发包）放入 agent 技能目录：

| 工具 | 目录 |
|---|---|
| DSH / Claude Code（Windows） | `C:\Users\<用户名>\.agents\skills\` |
| DSH / Claude Code（Linux/macOS） | `~/.agents/skills/` |
| 其它工具 | 按各自 skill 安装方式放置 |

安装后，在 agent 里说明你的意图即可，例如：

> 「按固定模板把这批简历整理成 2027 届表格，地点深圳，面试安排 D2/E2 填 9.19（周五）/9.20（周六）。」
> 「又有新简历了，填进已生成的表里。」

---

## 直接使用脚本

技能脚本位于 `skills/resume-stats-table/scripts/`，可独立调用（**仅 Python 标准库**，除第 1 步需 PyMuPDF）：

```bash
# 1) PDF → 文本（需安装 pymupdf）
pip install pymupdf
python extract_pdfs.py <简历目录> <输出目录>

# 2) 读取模板结构
python read_template.py <模板.xlsx>

# 3) 建新一届表格
python build_next_year.py <固定模板.xlsx> <输出.xlsx> <data.json> [config.json]

# 4) 后续增量追加（绝不覆盖原文件，另存 _新增 副本）
python append_candidates.py <已生成.xlsx> <新增data.json>
```

`config.json` 示例：
```json
{ "location": "深圳", "interview_schedule": { "D": "9.19（周五）", "E": "9.20（周六）" } }
```

---

## 目录结构

```
.
├── skills/resume-stats-table/
│   ├── SKILL.md                       # 技能主说明（agent 读取）
│   ├── references/
│   │   ├── data_schema.md             # 数据 JSON 结构
│   │   └── sheet_columns.md           # 主表列与取值规范
│   └── scripts/
│       ├── extract_pdfs.py            # PDF→文本（需 pymupdf）
│       ├── read_template.py           # 读模板结构
│       ├── build_next_year.py         # 建新一届表格
│       └── append_candidates.py       # 增量追加新增简历
└── resume-stats-table.skill           # 可分发的 .skill 归档
```

> 数据 JSON 结构见 `skills/resume-stats-table/references/data_schema.md`。

---

## License

按需添加（如 MIT 等）。本仓库作者：`curry@30`。
