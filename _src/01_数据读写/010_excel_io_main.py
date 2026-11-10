# -*- coding: utf-8 -*-
# ============================================================
# 010_excel_io_main.py
# 功能：
#   1) 模块对外聚合入口（01_数据读写）
#      1.1) 通过 _bootstrap.加载 拉起 011_excel_reader、012_excel_writer
#      1.2) 暴露统一接口：数据集 / 解析列引用 / 读取_文件 / 读取_激活簿 /
#                       列出文件表名 / 列出激活簿表名 /
#                       写出_新建文件 / 写出_当前激活簿
#   2) 命令行自检入口
#      - 直接运行可打印模块加载状态，便于排查路径问题
# ============================================================

from __future__ import annotations

import sys
from pathlib import Path

# ---- 1.1) 通过 _bootstrap 加载兄弟模块 ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent)
                )     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402

_reader = 加载("01_数据读写", "011_excel_reader")
_writer = 加载("01_数据读写", "012_excel_writer")

# ---- 1.2) 对外暴露统一接口 ----
数据集 = _reader.数据集
解析列引用 = _reader.解析列引用
读取_文件 = _reader.读取_文件
读取_激活簿 = _reader.读取_激活簿
列出文件表名 = _reader.列出文件表名
列出激活簿表名 = _reader.列出激活簿表名
获取激活簿标识 = _reader.获取激活簿标识

写出_新建文件 = _writer.写出_新建文件
写出_当前激活簿 = _writer.写出_当前激活簿


def main() -> None:
    """命令行自检：打印 reader/writer 暴露的对外名字。"""
    print("01_数据读写 模块加载成功")
    print("  数据集类       :", 数据集)
    print("  读取函数       : 读取_文件 / 读取_激活簿")
    print("  写出函数       : 写出_新建文件 / 写出_当前激活簿")
    print("  Sheet 列表函数 : 列出文件表名 / 列出激活簿表名")


if __name__ == "__main__":
    main()
