# 输出样例说明

## 预期输出文件清单

完整的流水线运行后，预期生成以下文件：

| 文件 | 格式 | 说明 |
|------|------|------|
| `renewal_intention_日期.xlsx` | Excel (3 Sheet) | 意向明细 + 统计汇总 + LP维度 |
| `renewal_report_日期.md` | Markdown | 管理报告（参见 `renewal_report_sample.md`） |
| `renewal_quality_日期.xlsx` | Excel (4 Sheet) | 销售质量诊断 |
| `pool_stats_日期.xlsx` | Excel | 池子拆分统计 |
| `lp_performance_日期.xlsx` | Excel | LP续费绩效卡 |
| `process_report_日期.xlsx` | Excel | 过程报告（8组×7池×14列） |
| `low_intention_日期.xlsx` | Excel | 低意向原因分类 |
| `group_stats_日期.xlsx` | Excel | 小组统计 |
| `group_report_日期.md` | Markdown | 小组汇总报告 |

## renewal_intention_日期.xlsx 结构

### Sheet1: 意向明细
21列：序号 / 学员ID / 归属LP / 小组 / 区域 / 池子 / 续费金额 / 语种 / 失联天数 / 近7天动作数 / 是否失联超过7天 / 意向等级 / 综合得分 / 正向信号数 / 风险信号数 / 总消息数 / 正向信号 / 风险信号 / 最近消息摘要 / 跟进行动建议

### Sheet2: 统计汇总
按等级汇总：人数、占比、信心度

### Sheet3: LP维度
按LP汇总：学员数、高意向率、风险率、质量评分、建议

## renewal_quality_日期.xlsx 结构

### Sheet1: 学员分析
每位学员的销售问题标签列表

### Sheet2: 意向统计
按意向等级的分布和问题类型交叉

### Sheet3: LP问题统计
每位LP的问题发生频次

### Sheet4: 问题类型汇总
各问题类型的人数和占比

## process_report_日期.xlsx 结构

两行表头，8组×7池×14列交叉表：
- 行：港澳1组 / 港澳2组 / 港澳3组 / 美澳1组 / 美澳2组 / 美澳3组 / 美澳4组 / 美澳5组
- 列：学员 / 签单率 / 未续 / 覆盖率 / H% / M% / L% / R% / >7天 / 断联率 / 退费占比 / 结课占比
- 核心池包括：执行中 / 停课 / 等班

> **注意：以上所有样例均为虚构数据，不含任何真实学员/家长/LP信息。**
