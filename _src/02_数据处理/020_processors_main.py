# -*- coding: utf-8 -*-
# ============================================================
# 020_processors_main.py
# 功能：02_数据处理 模块的聚合入口
#   1) 通过 _bootstrap 加载兄弟模块 021/022/023/024
#      1.1) 021_vlookup     → vlookup_合并
#      1.2) 022_regex_extract → 正则提取
#      1.3) 023_replace     → 查找替换 / 对照表替换
#      1.4) 024_split_columns → 按分隔符分列 / 按正则分列
#   2) 对外统一暴露的函数（GUI / 主入口直接 import 这里）
#   3) 不在本文件实现具体逻辑，避免与子文件重复（每个具体子文件 ≤ 600 行）
# ============================================================

from __future__ import annotations

import sys
from pathlib import Path

# ---- 1) 把 src 加入 sys.path，再用 _bootstrap 加载兄弟模块 ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent)
                )     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402

_vlookup = 加载("02_数据处理", "021_vlookup")
_regex_extract = 加载("02_数据处理", "022_regex_extract")
_replace = 加载("02_数据处理", "023_replace")
_split = 加载("02_数据处理", "024_split_columns")
_param_scanner = 加载("02_数据处理", "025_param_scanner")


# ---- 2) 对外统一函数 ----
vlookup_合并 = _vlookup.vlookup_合并
正则提取 = _regex_extract.正则提取
查找替换 = _replace.查找替换
对照表替换 = _replace.对照表替换
按分隔符分列 = _split.按分隔符分列
按正则分列 = _split.按正则分列
扫描参数 = _param_scanner.扫描参数
按参数提取 = _param_scanner.按参数提取
参数模式列表 = _param_scanner.全部模式列表
模式_单个 = _param_scanner.模式_单个
模式_任一 = _param_scanner.模式_任一
模式_全部 = _param_scanner.模式_全部
模式_所有 = _param_scanner.模式_所有
模式_字段及值 = _param_scanner.模式_字段及值


__all__ = [
    "vlookup_合并",
    "正则提取",
    "查找替换",
    "对照表替换",
    "按分隔符分列",
    "按正则分列",
    "扫描参数",
    "按参数提取",
    "参数模式列表",
    "模式_单个",
    "模式_任一",
    "模式_全部",
    "模式_所有",
    "模式_字段及值",
]
