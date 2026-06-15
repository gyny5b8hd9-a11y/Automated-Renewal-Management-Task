---
name: vipthink-renewal-intention
description: VIPTHINK LP 续费意向盘全流水线 — BI导出→意向分析→质量诊断→池子统计→LP绩效→低意向归因→小组统计→过程报告→飞书上表→钉钉双群推送
agent_created: true
---

# VIPTHINK 续费意向盘分析

## 触发词
续费意向、意向盘、续费池、流水线、周报、日报、pipeline、推送、跑流水线

## 流水线
`BI导出 → 意向分析(Step2) → 质量分析(Step2.5) → 池子统计(Step2.6) → LP绩效(Step2.7) → 低意向分析(Step2.8) → 小组统计(Step3) → 过程报告 → 飞书上表(Step3.5) → 钉钉双群推送(Step4)`

## 运行

### 自动化调度（已配置）
- **周报**: 每周五 09:00 → `pipeline.py --mode weekly --skip-export`
- **日报**: 周二~周日 09:00 → `pipeline.py --mode daily --skip-export`

### 手动运行
```bash
cd C:\Users\admin\.workbuddy\skills\vipthink-renewal-intention
C:\Users\admin\.workbuddy\binaries\python\envs\default\Scripts\python.exe scripts/pipeline.py --mode daily --skip-export
```

| 参数 | 说明 |
|------|------|
| `--mode weekly` | 周报，1条简报（总览+排名+过程表+飞书链接） |
| `--mode daily` | 日报，比较昨日变化，无变化静默不推送 |
| `--skip-export` | 复用已有 BI 导出文件（日常必加） |
| `--skip-push` | 跳过推送（仅分析不推送） |

### 强制推送（即使无变化）
```bash
python scripts/daily_report.py --today YYYY-MM-DD --yesterday YYYY-MM-DD --key KEY1 KEY2 --feishu-url URL
```

## 推送配置
- **双群同时推送**: `config.json` → `dingtalk_webhook_keys` (list)
- **周报输出**: 1条消息（简报）→ 意向总览 + 小组排名 + 过程分析表(8组×7池) + 飞书链接
- **日报输出**: 1条消息 → 变化总览 + 当前分布 + 升温/风险明细 + 飞书链接
- **周报 `--msg 1`**: 仅简报，不对分组推明细清单

## 评分体系
方案E 分类上限法：HIGH≥3 / MEDIUM≥1.5 / RISK≤-1（完整规则见 `references/signal_taxonomy.md`）

## 关键规则
- 价格敏感=高意向；已退款=直接风险
- 取数：全月全量（本月1号→昨天），`--skip-export` 复用已有导出
- 输出目录：`C:\Users\admin\Desktop\BI自动文档下载\haiwaisiwei_yuyifenxi\{date}\`
- 飞书三Sheet：续费意向盘(21列) / 过程分析(14列) / 低意向归因(11列)
- 脚本：`analyze_renewal.py` / `analyze_sales_quality.py` / `group_stats.py` / `pool_stats.py` / `lp_performance.py` / `low_intention_analysis.py` / `process_report.py` / `upload_feishu_sheet.py` / `dingtalk_push.py` / `daily_report.py`
