# -*- coding: utf-8 -*-
# ============================================================
# _bootstrap.py
# 功能：
#   1) 模块加载器（解决"数字开头文件名"无法直接 import 的问题）
#      1.1) 提供 加载(目录, 文件名) 接口：基于 importlib 按物理路径加载
#      1.2) 模块缓存：同名只加载一次，避免重复初始化
#      1.3) 自动定位 src 根目录：从 _bootstrap.py 自身位置反推
#   2) sys.path 注入
#      - 把 src/ 加入 sys.path，便于其他普通命名模块直接 import
# ============================================================

from __future__ import annotations

import sys
import importlib.util
from pathlib import Path
from types import ModuleType
from typing import Dict


# ---- 1) 全局缓存与路径常量 ----
_模块缓存: Dict[str, ModuleType] = {}
SRC根目录: Path = Path(__file__).resolve().parent

if str(SRC根目录) not in sys.path:
    # 让 _bootstrap 自身可被 import
    sys.path.insert(0, str(SRC根目录))


# ---- 1.1) 加载函数：按"目录名 + 文件名（不含 .py）"加载兄弟模块 ----
def 加载(目录名: str, 文件名: str) -> ModuleType:
    """
    按物理路径加载 src/目录名/文件名.py，并返回模块对象。

    :param 目录名: 编号子目录名，如 "01_数据读写"。空字符串 = src 根目录
    :param 文件名: 文件名（不含 .py），如 "011_excel_reader"
    :return: 加载完成的 ModuleType 对象
    """
    缓存键 = f"{目录名}/{文件名}"
    if 缓存键 in _模块缓存:
        return _模块缓存[缓存键]

    路径 = (SRC根目录 / 目录名 / f"{文件名}.py") if 目录名 else (SRC根目录 / f"{文件名}.py")
    if not 路径.exists():
        raise FileNotFoundError(f"模块不存在：{路径}")

    # 合法标识符模块名（防止 Qt 等 C 扩展反射时因数字开头崩溃）
    模块名 = f"sap_{文件名}" if 文件名[:1].isdigit() else 文件名

    spec = importlib.util.spec_from_file_location(模块名, str(路径))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法构造模块规范：{路径}")
    模块 = importlib.util.module_from_spec(spec)
    # 注册到 sys.modules，便于子模块互相 import
    sys.modules[模块名] = 模块
    _模块缓存[缓存键] = 模块
    spec.loader.exec_module(模块)
    return 模块
