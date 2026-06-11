# vipthink-renewal-intention

> VIPTHINK LP 续费意向盘全流水线自动化工具
>
> BI数据导出 → 意向分析 → 质量诊断 → 池子统计 → LP绩效卡 → 小组统计 → 过程报告 → 飞书上表 → 钉钉推送

---

## 1. 工具用途

本工具用于 LP（学习规划师）团队续费意向盘的自动化分析。从 SmartBI 导出包含 LP-家长对话的原始 Excel 数据，自动识别续费意向信号，将学员分为四类（高意向/中意向/观望/风险），生成多维度统计报告并推送至钉钉群。

**核心能力：**

- 基于 NLP 关键词匹配的 33 类续费信号自动识别
- 方案E 分类上限法综合评分（HIGH ≥ 3 / MEDIUM ≥ 1.5 / RISK ≤ -1）
- 10 步全自动流水线（BI 导出 → 分析 → 推送）
- 周报/日报双模式，支持飞书表格同步
- 多维度分析：学员级明细、LP 绩效卡、池子拆分、小组排名、低意向归因

## 2. 适用场景

| 场景 | 说明 |
|------|------|
| 每周续费意向周报 | 完整流水线，含 8 组 × 7 池交叉分析，钉钉推 4 条消息 |
| 每日续费变化日报 | 仅推送昨日变化的学员，无变化静默，排除周五 |
| 手动分析单份数据 | 只跑 analyze_renewal.py，输出 Excel 明细 + MD 报告 |

## 3. 安装

### 环境要求

- Python ≥ 3.10
- Node.js ≥ 18（用于 Playwright 浏览器自动化）
- 依赖包：`openpyxl`, `playwright`

```bash
# 安装 Python 依赖
pip install openpyxl playwright

# 安装 Playwright 浏览器
python -m playwright install chromium
```

### WorkBuddy Skill 安装

1. WorkBuddy 左侧「技能」→「导入技能」
2. 选择 `vipthink-renewal-intention.zip` 上传

### 手动安装（从源码）

```bash
# Windows
Expand-Archive vipthink-renewal-intention.zip -DestinationPath "$env:USERPROFILE\.workbuddy\skills\vipthink-renewal-intention" -Force

# macOS / Linux
mkdir -p ~/.workbuddy/skills/vipthink-renewal-intention
unzip -o vipthink-renewal-intention.zip -d ~/.workbuddy/skills/vipthink-renewal-intention
```

## 4. 配置

复制 `config.example.json` 为 `config.json`，填入实际值：

```json
{
  "paths": {
    "python_exe": "/path/to/python",
    "output_base": "/path/to/output",
    "smartbi_scripts": "/path/to/smartbi-cli/scripts",
    "lark_cli": "/path/to/lark-cli"
  },
  "task_name": "haiwaisiwei_yuyifenxi",
  "report_id": "<YOUR_SMARTBI_REPORT_ID>",
  "dingtalk_webhook_key": "<YOUR_DINGTALK_WEBHOOK_KEY>"
}
```

云端部署时，设置环境变量 `VIPTHINK_CONFIG=/opt/config.json` 指向自定义配置。

## 5. 运行方式

### 命令行

```bash
# 完整周报流水线
python scripts/pipeline.py --mode weekly

# 日报流水线（比较昨日变化）
python scripts/pipeline.py --mode daily

# 跳过 BI 导出（复用已有文件）
python scripts/pipeline.py --skip-export --mode weekly

# 跳过推送（仅分析）
python scripts/pipeline.py --skip-push --mode weekly

# 仅运行核心分析脚本
python scripts/analyze_renewal.py <输入Excel> --output-dir ./output
```

### WorkBuddy 对话触发

- "帮我做续费意向分析"
- "盘点续费池"
- "跑一下流水线"
- "发周报"

## 6. 输入要求

### Step 1 输入（BI 导出）

Excel 文件，由 SmartBI `SIMPLE_REPORT` 接口导出，必须包含以下列：

| 列名 | 说明 | 必填 |
|------|------|------|
| `学员ID` | 学员唯一标识 | ✅ |
| `材料全对话` | LP-家长完整对话历史 | ✅ |
| `LP` | 归属学习规划师 | 可选 |
| `池子` | 学员所在续费池 | 可选 |
| `小组` | 所属业务小组 | 可选 |

> **取数范围**：全月全量（本月1号 → 昨天），增量模式存在脏数据残留风险。

## 7. 输出说明

| 产物 | 格式 | 说明 |
|------|------|------|
| `renewal_intention_{date}.xlsx` | Excel (3 Sheet) | 意向明细 + 统计汇总 + LP 维度 |
| `renewal_report_{date}.md` | Markdown | 管理报告：Top10 + 风险信号 + LP对比 + 行动清单 |
| `renewal_quality_{date}.xlsx` | Excel (4 Sheet) | LP 销售过程质量诊断 |
| `pool_stats_{date}.xlsx` | Excel | 池子拆分统计 |
| `lp_performance_{date}.xlsx` | Excel | LP 续费绩效卡 |
| `process_report_{date}.xlsx` | Excel | 过程报告 (8组×7池×14列交叉表) |
| `low_intention_{date}.xlsx` | Excel | 低意向原因分类 |
| `group_stats_{date}.xlsx` | Excel | 小组维度统计 |
| `group_report_{date}.md` | Markdown | 小组汇总报告 |
| 钉钉推送 (周报) | Markdown×4 | 简报+小组排名+过程分析+LP诊断 |
| 钉钉推送 (日报) | Markdown×1 | 变化简报，无变化不推送 |
| 飞书表格 | 在线Sheet (3个) | 意向明细 + 过程分析 + 低意向分析 |

## 8. 意向分类标准

| 等级 | 含义 | 跟进建议 |
|------|------|---------|
| 🔥 高意向 | 主动询价、价格敏感、明确表态续费 | 48h 内推动成交 |
| 🟡 中意向 | 积极回应但未下决心 | 3 天内深度培育 |
| 🔵 观望 | 敷衍或沉默 | 保持常规触达 |
| 🔴 风险 | 明确拒绝、投诉、询问退款 | 升级主管介入 |

> **核心判断规则**：价格敏感/反复比价 = 高意向。认真算账说明客户在做购买决策。仅当同时表达"太贵不续"才降级为风险。

## 9. 流水线架构

```
Step1    BI 导出 (Playwright)              → 海外思维学员语义分析明细.xlsx
Step2    续费意向分析 (方案E评分)           → renewal_intention_{date}.xlsx
Step2.5  LP 销售质量分析 (非阻塞)           → renewal_quality_{date}.xlsx
Step2.6  池子拆分统计 (非阻塞)              → pool_stats_{date}.xlsx
Step2.7  LP 续费绩效卡 (非阻塞)             → lp_performance_{date}.xlsx
Step2.8  低意向原因分析 (非阻塞)            → low_intention_{date}.xlsx
Step3    小组统计 + 过程报告               → group_{stats,report}_{date}
Step3.5  飞书三Sheet上传                   → 飞书在线表格
Step4a   周报推送 (4条钉钉消息)
Step4b   日报推送 (1条，无变化静默)
```

## 10. 已知限制

| 限制 | 影响 | 缓解措施 |
|------|------|---------|
| BI 导出 5万+行，Playwright 常超时 | Step1 偶发失败 | `--skip-export` 复用已有文件 |
| 日报首日无对比数据 | 无法比较变化 | 降级为"首日简报" |
| lark-cli 无法设置公开访问 | 飞书链接仅组织内可访问 | 暂无缓解 |
| 两个评分体系不一致 | 意向分 vs 质量分 | 方案E(分类上限) vs 简单累加，已注明 |

## 11. 文件结构

```
vipthink-renewal-intention/
├── README.md                     # 本文件 — 安装/运行/输入输出
├── SKILL.md                      # WorkBuddy Skill 触发定义
├── project-meta.json             # 工具元数据（注册表用）
├── PUBLIC_SAFE_CHECKLIST.md      # 安全清单
├── CHANGELOG.md                  # 版本变化
├── REPLAY.md                     # 复核指南
├── TOOL_REVIEW.md                # 复用评审
├── config.example.json           # 配置模板（脱敏）
├── references/
│   └── signal_taxonomy.md        # 33 类续费信号分类体系
├── rules/
│   └── scoring_rules.yaml        # 加权评分规则 (v3.0)
├── templates/
│   ├── renewal_report_template.md
│   └── quality_report_template.md
├── scripts/
│   ├── pipeline.py               # 主流水线编排器
│   ├── config.py                 # 配置加载器
│   ├── analyze_renewal.py        # Step2 意向分析
│   ├── analyze_sales_quality.py  # Step2.5 质量诊断
│   ├── pool_stats.py             # Step2.6 池子统计
│   ├── lp_performance.py         # Step2.7 绩效卡
│   ├── low_intention_analysis.py # Step2.8 低意向归因
│   ├── group_stats.py            # Step3 小组统计
│   ├── process_report.py         # Step3 过程报告
│   ├── upload_feishu_sheet.py    # Step3.5 飞书上表
│   ├── dingtalk_push.py          # Step4a 周报推送
│   ├── daily_report.py           # Step4b 日报推送
│   └── run_export.py             # BI 导出辅助
├── samples/
│   ├── input/                    # 输入样例
│   └── output/                   # 输出样例
├── docs/
│   ├── business-glossary.md      # 业务口径与字段定义
│   ├── scoring-design.md         # 评分方案设计说明
│   └── pipeline-flow.md          # 流水线流程详解
└── .gitignore
```

---

*更新时间：2026-06-10 | 作者：VIPTHINK LP 效能团队*
