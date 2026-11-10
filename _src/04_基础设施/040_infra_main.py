# -*- coding: utf-8 -*-
# ============================================================
# 040_infra_main.py
# 功能：
#   1) 04_基础设施 模块对外聚合入口
#      1.1) 统一对外暴露：debug_print / 记录错误 / UI日志桥 / UI信息 / UI警告 / UI错误
#      1.2) 通过 _bootstrap.加载 拉起兄弟文件 041_log.py
#   2) 命令行自检入口
#      - 直接运行本文件可触发一次 debug_print 与 记录错误 演示日志生效
# ============================================================

from __future__ import annotations

import sys
from pathlib import Path

# ---- 1.2) 通过 _bootstrap 加载兄弟模块 ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402

_log = 加载("04_基础设施", "041_log")

# ---- 1.1) 对外暴露的统一接口 ----
debug_print = _log.debug_print
记录错误 = _log.记录错误
UI日志桥 = _log.UI日志桥
UI信息 = _log.UI信息
UI警告 = _log.UI警告
UI错误 = _log.UI错误
获取日志根目录 = _log.获取日志根目录


def main() -> None:
    """命令行自检：触发一次调试输出与错误记录。"""
    debug_print("基础设施自检：debug_print 已生效")
    记录错误("基础设施自检：这是一条测试错误（请到 项目名_log/error/LOG.txt 查看）")
    print(f"日志根目录：{获取日志根目录()}")


if __name__ == "__main__":
    main()
