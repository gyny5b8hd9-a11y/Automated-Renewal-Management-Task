# REPLAY.md

> 复核指南：让团队其他成员能验证工具是否能跑通

---

## 一、环境准备

### 1.1 前置条件

- [ ] Python ≥ 3.10 已安装
- [ ] Node.js ≥ 18 已安装（Playwright）
- [ ] 有权访问 SmartBI 报表系统
- [ ] 已配置钉钉 Webhook（测试群）

### 1.2 安装依赖

```bash
# Python 依赖
pip install openpyxl playwright

# Playwright 浏览器
python -m playwright install chromium

# lark-cli（飞书上表需要）
npm install -g @anthropic/lark-cli
```

### 1.3 配置

```bash
# 复制配置模板
cp config.example.json config.json

# 编辑 config.json，填入：
#   - smartbi_scripts: SmartBI CLI 脚本路径
#   - report_id: SmartBI 报表 ID
#   - dingtalk_webhook_key: 测试群 Webhook
#   - output_base: 输出目录
#   - lark_cli: lark-cli 安装路径
```

---

## 二、验证步骤

### Step 1: 核心分析脚本

**目的：** 验证信号识别和评分逻辑正确。

```bash
# 使用样例数据运行
python scripts/analyze_renewal.py samples/input/renewal_chat_sample.xlsx --output-dir ./test_output

# 预期结果：
#   - test_output/renewal_intention_YYYY-MM-DD.xlsx 生成
#   - test_output/renewal_report_YYYY-MM-DD.md 生成
#   - 控制台输出各等级学员数量汇总
```

**验证点：**

| 项目 | 检查方式 |
|------|---------|
| 高意向学员 | 打开 Excel 明细 Sheet，确认"主动询价""价格敏感"等标签正确激活 |
| 风险学员 | 确认含"退款"关键词的学员被标记为 风险 |
| 价格算账 | 含"比价""算下来"的对话应标记为高意向 |
| 输出格式 | 3 Sheet（明细 + 汇总 + LP维度），列名完整 |

### Step 2: 子脚本独立测试

```bash
# 池子统计
python scripts/pool_stats.py test_output/renewal_intention_YYYY-MM-DD.xlsx --output-dir ./test_output

# LP 绩效卡
python scripts/lp_performance.py test_output/renewal_intention_YYYY-MM-DD.xlsx --output-dir ./test_output

# 小组统计
python scripts/group_stats.py test_output/renewal_intention_YYYY-MM-DD.xlsx --output-dir ./test_output

# 过程报告
python scripts/process_report.py test_output/renewal_intention_YYYY-MM-DD.xlsx --output-dir ./test_output
```

### Step 3: 完整流水线（跳过 BI 导出）

```bash
# 用样例数据模拟完整流程
python scripts/pipeline.py --skip-export --skip-push --mode weekly

# 预期：
#   - 所有中间产物生成（7 个 Excel + 3 个 MD）
#   - pipeline_YYYY-MM-DD.log 记录完整流程
#   - 无报错、无异常退出
```

### Step 4: 飞书上表测试（可选）

```bash
python scripts/upload_feishu_sheet.py \
  test_output/renewal_intention_YYYY-MM-DD.xlsx \
  --raw-excel test_output/海外思维学员语义分析明细.xlsx \
  --process-excel test_output/process_report_YYYY-MM-DD.xlsx \
  --low-excel test_output/low_intention_YYYY-MM-DD.xlsx \
  --json
```

**预期：** 返回飞书表格 URL。

### Step 5: 钉钉推送测试（可选）

```bash
# 用测试群 Webhook 验证
python scripts/dingtalk_push.py \
  --intent-excel test_output/renewal_intention_YYYY-MM-DD.xlsx \
  --quality-excel test_output/renewal_quality_YYYY-MM-DD.xlsx \
  --raw-excel test_output/海外思维学员语义分析明细.xlsx \
  --date-start 2026-06-01 --date-end 2026-06-09 \
  --key <TEST_WEBHOOK_KEY> \
  --msg 0
```

---

## 三、预期输出对照

### 3.1 意向分布参考范围

基于历史数据，正常业务状态下意向分布如下：

| 等级 | 合理范围 |
|------|---------|
| 高意向 | 5% - 15% |
| 中意向 | 15% - 30% |
| 观望 | 40% - 60% |
| 风险 | 5% - 15% |

> 如分布严重偏离（如高意向 < 2% 或 > 30%），需检查：
> - 取数范围是否正确
> - 对话数据是否完整
> - 信号关键词是否覆盖当前业务场景

### 3.2 输出文件结构

```
test_output/
├── renewal_intention_YYYY-MM-DD.xlsx   # 3 Sheet
├── renewal_report_YYYY-MM-DD.md        # 管理报告
├── renewal_quality_YYYY-MM-DD.xlsx     # 4 Sheet
├── pool_stats_YYYY-MM-DD.xlsx          # 池子统计
├── lp_performance_YYYY-MM-DD.xlsx      # 绩效卡
├── process_report_YYYY-MM-DD.xlsx      # 过程报告
├── low_intention_YYYY-MM-DD.xlsx       # 低意向归因
├── group_stats_YYYY-MM-DD.xlsx         # 小组统计
├── group_report_YYYY-MM-DD.md          # 小组报告
└── pipeline_YYYY-MM-DD.log             # 运行日志
```

---

## 四、常见问题排查

| 症状 | 可能原因 | 解决方式 |
|------|---------|---------|
| 意向分析返回空 | 对话列名不匹配 | 检查 Excel 是否包含 `材料全对话` 列 |
| 信号识别率异常低 | 对话格式不标准 | 检查对话是否含 `员工:` `客户:` 标识 |
| pool_stats 统计不全 | 池子字段映射缺失 | 确认池子值为标准名称 |
| Playwright 超时 | 网络慢或报表过大 | 设置更长的超时或使用 `--skip-export` |
| 飞书上表失败 | lark-cli 登录过期 | 重新执行 `lark-cli login` |
| 钉钉推送无响应 | Webhook 配置错误 | 用 curl 测试 Webhook 连通性 |

---

## 五、环境清理

```bash
# 清理测试产物
rm -rf test_output/
```

---

*最后验证：2026-06-10 | 验证人：VIPTHINK LP 效能团队*
