"""统一配置加载器。
优先从环境变量 VIPTHINK_CONFIG 读取，否则从同目录 config.json 读取。
云端部署时设置 VIPTHINK_CONFIG=/opt/config.json 即可。
"""
import json, os
from pathlib import Path

_cfg = None

def _load():
    global _cfg
    if _cfg is not None:
        return _cfg
    path = os.environ.get("VIPTHINK_CONFIG", str(Path(__file__).parent / "config.json"))
    with open(path, "r", encoding="utf-8") as f:
        _cfg = json.load(f)
    return _cfg

def path(key: str) -> Path:
    return Path(_load()["paths"][key])

def get(key: str):
    return _load()[key]
