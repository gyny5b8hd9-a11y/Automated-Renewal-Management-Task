#!/usr/bin/env python3
"""Run the fixed export script for 2026-05-18 → 2026-05-22"""
import asyncio, sys
from pathlib import Path

import config

sys.path.insert(0, str(Path(__file__).resolve().parent))
from export_simple_report_fixed import export_simple_report

report_id = config.get("report_id")
out_dir = config.path("output_base") / config.get("task_name") / "2026-05-23"
out_path = out_dir / "海外思维学员语义分析明细.xlsx"

filters = [
    ["做工起始时间", "2026-05-18", "2026-05-18"],
    ["做工结束时间", "2026-05-22", "2026-05-22"],
    ["池子", "", ""],
]

asyncio.run(export_simple_report(
    report_id=report_id,
    output_path=str(out_path),
    filters=filters,
))
