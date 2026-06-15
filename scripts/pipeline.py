#!/usr/bin/env python3
"""BI导出 → 续费意向分析 → 小组统计 → 钉钉推送 自动化流水线。

取数策略: 全月全量（本月1号 → 昨天），每次从头跑，无状态、幂等。

用法:
  python pipeline.py [--date-start 2026-05-01 --date-end 2026-05-24] [--dingtalk-key KEY] [--test]
"""

import argparse
import asyncio
import json
import os
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

import config

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
SMARTBI_SCRIPTS = config.path("smartbi_scripts")
PYTHON = str(config.path("python_exe"))

OUTPUT_BASE = config.path("output_base")
TASK_NAME = config.get("task_name")
REPORT_ID = config.get("report_id")

DEFAULT_WEBHOOK_KEYS = config.get("dingtalk_webhook_keys")
DEFAULT_SIGN_KEY = ""


class PipelineLogger:
    def __init__(self, log_path: Path):
        self.path = log_path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def log(self, msg: str):
        ts = date.today().isoformat()
        line = f"[{ts}] {msg}"
        print(line)
        with self.path.open("a", encoding="utf-8") as f:
            f.write(line + "\n")


def month_date_range() -> tuple[date, date]:
    """全月全量: date_start = 本月1号, date_end = 昨天

    设计决策 (2026-05-25):
    续费意向分析取数范围经四角色多轮讨论，决议采用全月全量方案。
    LP 对话存在大量事后修改/删除，增量模式无法检测删除会导致脏数据残留。
    瓶颈在 SmartBI 导出（3-8min），不在分析计算（30-90s）。
    全量方案保持无状态、幂等、可审计。
    """
    today = date.today()
    end = today - timedelta(days=1)
    start = today.replace(day=1)
    return start, end


def resolve_out_dir(run_date: date) -> Path:
    return OUTPUT_BASE / TASK_NAME / run_date.isoformat()


def step_export(start: date, end: date, out_dir: Path, logger: PipelineLogger) -> Path | None:
    """Step 1: 导出 SmartBI SIMPLE_REPORT"""
    logger.log(f"Step 1: 导出报表 {start}→{end}")

    # Append export script dir to sys.path and import
    sys.path.insert(0, str(SMARTBI_SCRIPTS))
    from export_simple_report import export_simple_report

    out_path = out_dir / "海外思维学员语义分析明细.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    filters = [
        ["做工起始时间", start.isoformat(), start.isoformat()],
        ["做工结束时间", end.isoformat(), end.isoformat()],
        ["池子", "", ""],
    ]

    result = asyncio.run(export_simple_report(
        report_id=REPORT_ID,
        output_path=str(out_path),
        filters=filters,
        max_rows=60000,
    ))

    logger.log(f"  导出完成: {result['output']} ({result['bytes']} bytes, {result['rowCount']} rows)")
    return out_path


def step_analyze(raw_excel: Path, out_dir: Path, logger: PipelineLogger) -> Path:
    """Step 2: 续费意向分析"""
    logger.log(f"Step 2: 续费意向分析")

    cmd = [
        PYTHON,
        str(SCRIPTS / "analyze_renewal.py"),
        str(raw_excel),
        "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})

    if result.returncode != 0:
        err_detail = result.stderr[-500:] if result.stderr else result.stdout[-500:]
        logger.log(f"  ERROR: {err_detail}")
        raise RuntimeError(f"续费意向分析失败(rc={result.returncode})")
    
    # Show last meaningful line from stdout (the summary line)
    out_lines = [l for l in result.stdout.strip().split('\n') if l.strip() and '进度' not in l]
    if out_lines:
        logger.log(f"  {out_lines[-1][:200]}")

    # Find output files
    today = date.today()
    intent_xlsx = out_dir / f"renewal_intention_{today}.xlsx"
    intent_md = out_dir / f"renewal_report_{today}.md"

    logger.log(f"  完成: {intent_xlsx.name}, {intent_md.name}")
    return intent_md


def step_quality(raw_excel: Path, out_dir: Path, logger: PipelineLogger) -> Path | None:
    """Step 2.5: LP 销售过程质量分析"""
    logger.log("Step 2.5: 销售质量分析")
    
    cmd = [
        PYTHON,
        str(SCRIPTS / "analyze_sales_quality.py"),
        str(raw_excel),
        "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    
    if result.returncode != 0:
        err_detail = result.stderr[-500:] if result.stderr else result.stdout[-500:]
        logger.log(f"  WARNING: 质量分析失败 (非阻塞): {err_detail}")
        return None
    
    # Show last meaningful line
    out_lines = [l for l in result.stdout.strip().split('\n') if l.strip()]
    if out_lines:
        logger.log(f"  {out_lines[-1][:200]}")
    
    today = date.today()
    quality_xlsx = out_dir / f"renewal_quality_{today}.xlsx"
    if quality_xlsx.exists():
        logger.log(f"  完成: {quality_xlsx.name}")
        return quality_xlsx
    return None


def step_pool_stats(raw_excel: Path, out_dir: Path, logger: PipelineLogger) -> Path | None:
    """Step 2.6: 池子拆分统计"""
    logger.log("Step 2.6: 池子统计")
    cmd = [
        PYTHON, str(SCRIPTS / "pool_stats.py"),
        str(raw_excel), "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    if result.returncode != 0:
        logger.log(f"  WARNING: 池子统计失败")
        return None
    for line in result.stdout.strip().split('\n'):
        if line.strip():
            logger.log(f"  {line.strip()[:200]}")
    today = date.today()
    p = out_dir / f"pool_stats_{today}.xlsx"
    return p if p.exists() else None


def step_lp_performance(raw_excel: Path, out_dir: Path, logger: PipelineLogger) -> Path | None:
    """Step 2.7: LP续费绩效卡"""
    logger.log("Step 2.7: LP绩效卡")
    cmd = [
        PYTHON, str(SCRIPTS / "lp_performance.py"),
        str(raw_excel), "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    if result.returncode != 0:
        logger.log(f"  WARNING: LP绩效卡失败")
        return None
    for line in result.stdout.strip().split('\n'):
        if line.strip():
            logger.log(f"  {line.strip()[:200]}")
    today = date.today()
    p = out_dir / f"lp_performance_{today}.xlsx"
    return p if p.exists() else None


def step_low_intention(intent_xlsx: Path, raw_excel: Path, out_dir: Path, logger: PipelineLogger) -> Path | None:
    """Step 2.8: 低意向客户原因分类"""
    logger.log("Step 2.8: 低意向原因分析")
    cmd = [
        PYTHON, str(SCRIPTS / "low_intention_analysis.py"),
        str(intent_xlsx), str(raw_excel),
        "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    if result.returncode != 0:
        logger.log(f"  WARNING: 低意向分析失败")
        return None
    for line in result.stdout.strip().split('\n'):
        if line.strip():
            logger.log(f"  {line.strip()[:200]}")
    today = date.today()
    p = out_dir / f"low_intention_{today}.xlsx"
    return p if p.exists() else None


def step_group_stats(raw_excel: Path, intent_xlsx: Path, out_dir: Path, logger: PipelineLogger) -> Path:
    """Step 3: 小组维度统计"""
    logger.log("Step 3: 小组统计")

    cmd = [
        PYTHON,
        str(SCRIPTS / "group_stats.py"),
        str(raw_excel),
        str(intent_xlsx),
        "--output-dir", str(out_dir),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})

    if result.returncode != 0:
        err_detail = result.stderr[-500:] if result.stderr else result.stdout[-500:]
        logger.log(f"  ERROR: {err_detail}")
        raise RuntimeError(f"小组统计失败(rc={result.returncode})")

    today = date.today()
    report_md = out_dir / f"group_report_{today}.md"
    logger.log(f"  完成: {report_md.name}")
    return report_md


def step_push(raw_excel: Path, intent_xlsx: Path, quality_xlsx: Path,
               webhook_keys: list, date_start: str, date_end: str,
               feishu_url: str, logger: PipelineLogger, out_dir: Path = None):
    """Step 4: 钉钉推送 — 委托 dingtalk_push.py 子进程，支持多群"""
    logger.log("Step 4: 钉钉推送")

    run_date = date.today()
    pool_excel = out_dir / f"process_report_{run_date}.xlsx" if out_dir else None

    cmd = [
        PYTHON, str(SCRIPTS / "dingtalk_push.py"),
        "--intent-excel", str(intent_xlsx),
        "--quality-excel", str(quality_xlsx),
        "--raw-excel", str(raw_excel),
        "--date-start", date_start,
        "--date-end", date_end,
        "--key",
    ]
    cmd.extend(webhook_keys)
    cmd.extend(["--msg", "1"])  # 1=仅简报（总览+排名+过程表+飞书链接）
    if feishu_url:
        cmd.extend(["--feishu-url", feishu_url])
    if pool_excel and pool_excel.exists():
        cmd.extend(["--pool-excel", str(pool_excel)])
    
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    
    if result.returncode != 0:
        err_detail = result.stderr[-500:] if result.stderr else result.stdout[-500:]
        logger.log(f"  ERROR: {err_detail}")
        raise RuntimeError(f"钉钉推送失败(rc={result.returncode})")
    
    for line in result.stdout.strip().split('\n'):
        if line.strip():
            logger.log(f"  {line.strip()[:200]}")
    logger.log(f"  推送完成")


def step_push_daily(intent_xlsx: Path, webhook_keys: list, feishu_url: str, logger: PipelineLogger):
    """Step 4 (日版): 日报比较 → 钉钉推送，支持多群"""
    logger.log("Step 4 (日报): 比较昨日变化并推送")

    import subprocess
    cmd = [
        PYTHON,
        str(SCRIPTS / "daily_report.py"),
        "--today", date.today().isoformat(),
        "--key",
    ]
    cmd.extend(webhook_keys)
    cmd.extend(["--silent-on-no-change"])
    if feishu_url:
        cmd.extend(["--feishu-url", feishu_url])
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})

    # daily_report.py 无变化时 exit 0 但输出 "无变化，跳过推送"
    stdout_text = result.stdout.strip() if result.stdout else ""
    if result.returncode != 0:
        err_detail = result.stderr[-500:] if result.stderr else stdout_text[-500:]
        logger.log(f"  ERROR: {err_detail}")
        raise RuntimeError(f"日报推送失败(rc={result.returncode})")
    
    for line in stdout_text.split("\n"):
        if line.strip():
            logger.log(f"  {line.strip()[:200]}")
    
    if "无变化" in stdout_text:
        logger.log("  日报无变化，未推送")
    else:
        logger.log("  日报推送完成")


def step_upload_feishu(intent_xlsx: Path, raw_xlsx: Path, process_xlsx: Path, low_xlsx: Path, logger: PipelineLogger) -> str:
    """Step 3.5: 上传意向 Excel + 过程分析表 + 低意向分析到飞书"""
    logger.log("Step 3.5: 上传飞书表格")
    
    cmd = [
        PYTHON,
        str(SCRIPTS / "upload_feishu_sheet.py"),
        str(intent_xlsx),
        "--raw-excel", str(raw_xlsx),
        "--process-excel", str(process_xlsx),
        "--low-excel", str(low_xlsx),
        "--json",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    
    if result.returncode != 0:
        err = result.stderr[-300:] if result.stderr else result.stdout[-300:]
        logger.log(f"  WARNING: 飞书上传失败: {err}")
        return ""
    
    try:
        import json
        # 过滤出纯JSON行（忽略stderr warn混入stdout的情况）
        stdout_text = result.stdout.strip() if result.stdout else ""
        json_line = ""
        for line in stdout_text.split("\n"):
            line = line.strip()
            if line.startswith("{") and line.endswith("}"):
                json_line = line
                break
        if not json_line:
            logger.log(f"  WARNING: 飞书上传返回解析失败")
            return ""
        data = json.loads(json_line)
        url = data.get("url", "")
        logger.log(f"  飞书表格: {url}")
        return url
    except:
        logger.log(f"  WARNING: 飞书上传返回解析失败")
        return ""


def main():
    parser = argparse.ArgumentParser(description="BI导出→意向分析→小组统计→钉钉推送 流水线")
    parser.add_argument("--date-start", help="做工起始时间 (YYYY-MM-DD), 默认本月1号")
    parser.add_argument("--date-end", help="做工结束时间 (YYYY-MM-DD), 默认昨天")
    parser.add_argument("--mode", choices=["weekly", "daily"], default="weekly",
                        help="周报(weekly)或日报(daily), 均使用全月全量取数, 默认weekly")
    parser.add_argument("--dingtalk-key", nargs="*", default=DEFAULT_WEBHOOK_KEYS,
                        help="钉钉 Webhook Key（可多个），默认从 config 读取")
    parser.add_argument("--sign-key", default=DEFAULT_SIGN_KEY, help="钉钉签名密钥")
    parser.add_argument("--skip-export", action="store_true", help="跳过导出步骤")
    parser.add_argument("--skip-analyze", action="store_true", help="跳过分析步骤")
    parser.add_argument("--skip-push", action="store_true", help="跳过推送步骤")
    parser.add_argument("--test", action="store_true", help="测试模式: 推送时仅发500字符")
    args = parser.parse_args()

    start = date.fromisoformat(args.date_start) if args.date_start else None
    end = date.fromisoformat(args.date_end) if args.date_end else None

    if start is None or end is None:
        start, end = month_date_range()

    run_date = date.today()
    out_dir = resolve_out_dir(run_date)
    log_path = out_dir / f"pipeline_{run_date}.log"
    logger = PipelineLogger(log_path)

    logger.log(f"=== 流水线启动 ({args.mode}) ===")
    logger.log(f"日期范围: {start} → {end}")
    logger.log(f"输出目录: {out_dir}")

    try:
        raw_excel = None
        if not args.skip_export:
            raw_excel = step_export(start, end, out_dir, logger)
        else:
            # 跳过导出时，找最新可用的数据文件
            candidates = sorted(OUTPUT_BASE.rglob(f"{TASK_NAME}/*/海外思维学员语义分析明细.xlsx"), reverse=True)
            if candidates:
                raw_excel = candidates[0]
                out_dir = raw_excel.parent  # 使用已有数据的目录
                logger.log(f"跳过导出，使用已有文件: {raw_excel}")
            else:
                raise FileNotFoundError(f"跳过导出但未找到任何已有数据文件")

        if not args.skip_analyze:
            intent_md = step_analyze(raw_excel, out_dir, logger)
        else:
            intent_md = out_dir / f"renewal_report_{run_date}.md"
            if not intent_md.exists():
                raise FileNotFoundError(f"跳过分析但报告不存在: {intent_md}")
            logger.log(f"跳过分析，使用已有报告: {intent_md}")

        # Step 2.5: Sales quality analysis (非阻塞，失败不影响后续)
        quality_xlsx = None
        if not args.skip_analyze:
            quality_xlsx = step_quality(raw_excel, out_dir, logger)

        # Run group stats
        intent_xlsx = out_dir / f"renewal_intention_{run_date}.xlsx"
        if not args.skip_analyze:
            group_md = step_group_stats(raw_excel, intent_xlsx, out_dir, logger)
            step_pool_stats(raw_excel, out_dir, logger)
            step_lp_performance(raw_excel, out_dir, logger)
            low_xlsx = step_low_intention(intent_xlsx, raw_excel, out_dir, logger)
            # Generate process report for push
            subprocess.run(
                [PYTHON, str(SCRIPTS / "process_report.py"), str(raw_excel), "--output-dir", str(out_dir)],
                capture_output=True, text=True, encoding="utf-8",
                cwd=str(ROOT), env={**os.environ, "PYTHONIOENCODING": "utf-8"})
            logger.log("  过程报告已生成")
        else:
            group_md = out_dir / f"group_report_{run_date}.md"
            if not group_md.exists():
                # Run it anyway
                group_md = step_group_stats(raw_excel, intent_xlsx, out_dir, logger)

        # Step 3.5: Upload to Feishu
        feishu_url = ""
        if not args.skip_push:
            feishu_url = step_upload_feishu(intent_xlsx, raw_excel, 
                out_dir / f"process_report_{run_date}.xlsx",
                out_dir / f"low_intention_{run_date}.xlsx", logger)
        
        if not args.skip_push:
            if args.mode == "daily":
                step_push_daily(intent_xlsx, args.dingtalk_key, feishu_url, logger)
            else:
                # Use quality_xlsx from step_quality, or find latest
                if not quality_xlsx:
                    candidates = sorted(out_dir.glob("renewal_quality_*.xlsx"), reverse=True)
                    quality_xlsx = candidates[0] if candidates else None
                if quality_xlsx:
                    step_push(raw_excel, intent_xlsx, quality_xlsx,
                              args.dingtalk_key, start.isoformat(), end.isoformat(),
                              feishu_url, logger, out_dir)
                else:
                    logger.log("跳过推送: 未找到质量分析文件")
        else:
            logger.log("跳过推送")

        logger.log("=== 流水线完成 ===")

        # Print summary
        print(json.dumps({
            "status": "ok",
            "date_range": f"{start} → {end}",
            "outputs": {
                "raw_excel": str(raw_excel),
                "intent_md": str(intent_md),
                "group_md": str(group_md),
                "log": str(log_path),
                "feishu_url": feishu_url,
            },
        }, ensure_ascii=False, indent=2))

    except Exception as e:
        logger.log(f"!!! 流水线失败: {e}")
        print(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
