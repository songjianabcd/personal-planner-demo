# -*- coding: utf-8 -*-
# ============================================================
# 041_log.py
# 功能：
#   1) 日志根目录管理
#      1.1) 获取日志根目录()：返回 src 同级 项目名_log/，并确保 error/ 与 debug/ 子目录存在
#      1.2) 生成日志行(描述, 文件名, 行号)：统一格式 "-----yyyy.mm.dd hh:mm:ss-------xxx.py 文件，xxx行   ***内容"
#      1.3) 检查并瘦身日志(文件, 阈值, 保留)：超阈值归档到 汇总_log.txt，仅留尾部
#   2) 调试输出
#      2.1) debug_print(消息)：[DEBUG][时间][文件:行号] 消息 → stdout + debug/debug_log.txt
#      2.2) DEBUG_MODE=0 时关闭，瘦身阈值 5000/500
#   3) 错误日志
#      3.1) 记录错误(描述)：写 stderr + error/LOG.txt（禁止弹框），瘦身阈值 10000/500
#   4) UI 日志桥接
#      4.1) UI日志桥.注册(回调)：GUI 启动时注册接收回调
#      4.2) UI信息/UI警告/UI错误：业务侧调用，同时推送 GUI 与底层日志
# ============================================================

from __future__ import annotations

import os
import sys
import inspect
import datetime
from pathlib import Path
from typing import Callable, Optional


# ---- 1.1) 获取日志根目录：定位到 src/ 的父目录下 项目名_log/ ----
def 获取日志根目录() -> Path:
    """
    返回 src 同级的 项目名_log/ 目录，并确保 error/ 与 debug/ 子目录存在。

    :return: 日志根目录 Path 对象
    """
    src目录 = Path(__file__).resolve().parent.parent                  # 当前文件位于 src/04_基础设施/，向上两级到 src
    项目根 = src目录.parent
    日志根 = 项目根 / f"{项目根.name}_log"
    (日志根 / "error").mkdir(parents=True, exist_ok=True)
    (日志根 / "debug").mkdir(parents=True, exist_ok=True)
    return 日志根


# ---- 1.2) 生成日志行：统一文本格式 ----
def 生成日志行(错误描述: str, 文件名: str = "", 行号: int = 0) -> str:
    """
    按规范生成日志一行。当文件名为空时，自动从调用栈第二层取。

    :param 错误描述: 错误内容文本
    :param 文件名: 调用方文件名（不含路径），留空 = 自动取
    :param 行号:   调用方行号，留空 = 自动取
    :return: 单行日志文本（无换行）
    """
    现在 = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")
    if not 文件名:
        try:
            上层 = inspect.stack()[2]
            文件名 = Path(上层.filename).name
            行号 = 上层.lineno
        except Exception:
            文件名, 行号 = "?", 0
    return f"-----{现在}-------{文件名} 文件，{行号}行   {错误描述}"


# ---- 1.3) 检查并瘦身日志：超阈值时把旧内容归档到 汇总_log.txt ----
def 检查并瘦身日志(日志文件: Path, 最大行数: int = 10000, 保留行数: int = 500) -> None:
    """
    超过最大行数时，把开头部分归档到 汇总_log.txt，只保留尾部 保留行数 行。

    :param 日志文件: 待瘦身的日志文件
    :param 最大行数: 触发瘦身的阈值
    :param 保留行数: 瘦身后保留的尾部行数
    """
    try:
        if not 日志文件.exists():
            return
        with 日志文件.open("r", encoding="utf-8", errors="ignore") as f:
            行列表 = f.readlines()
        if len(行列表) <= 最大行数:
            return
        归档行 = 行列表[:-保留行数]
        保留 = 行列表[-保留行数:]
        汇总文件 = 获取日志根目录() / "汇总_log.txt"
        with 汇总文件.open("a", encoding="utf-8") as f:
            f.write(f"\n===== 归档自 {日志文件.name} {datetime.datetime.now()} =====\n")
            f.writelines(归档行)
        with 日志文件.open("w", encoding="utf-8") as f:
            f.writelines(保留)
    except Exception:
        pass                                                          # 瘦身失败不应影响主流程


# ---- 2.1) debug_print：调试输出，写 stdout + debug/debug_log.txt ----
def _debug_开关() -> bool:
    """读取环境变量 DEBUG_MODE，默认开启。"""
    return os.environ.get("DEBUG_MODE", "1") != "0"


def debug_print(消息: str, 文件名: str = "", 行号: int = 0) -> None:
    """
    输出格式 [DEBUG][时间][文件:行号] 消息 到 stdout 与 debug/debug_log.txt。
    DEBUG_MODE=0 时关闭。

    :param 消息: 调试文本
    :param 文件名: 调用方文件名，留空 = 自动取
    :param 行号:   调用方行号，留空 = 自动取
    """
    if not _debug_开关():
        return
    try:
        if not 文件名:
            上层 = inspect.stack()[1]
            文件名 = Path(上层.filename).name
            行号 = 上层.lineno
        现在 = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")
        行 = f"[DEBUG][{现在}][{文件名}:{行号}] {消息}"
        print(行, file=sys.stdout, flush=True)
        日志文件 = 获取日志根目录() / "debug" / "debug_log.txt"
        with 日志文件.open("a", encoding="utf-8") as f:
            f.write(行 + "\n")
        检查并瘦身日志(日志文件, 最大行数=5000, 保留行数=500)
    except Exception:
        pass


# ---- 3.1) 记录错误：错误日志，写 stderr + error/LOG.txt（禁止弹框） ----
def 记录错误(错误描述: str) -> None:
    """
    把错误同时写到 stderr 与 error/LOG.txt，禁止弹框。

    :param 错误描述: 错误内容文本
    """
    try:
        上层 = inspect.stack()[1]
        文件名 = Path(上层.filename).name
        行号 = 上层.lineno
        行 = 生成日志行(错误描述, 文件名, 行号)
        print(行, file=sys.stderr, flush=True)
        日志文件 = 获取日志根目录() / "error" / "LOG.txt"
        with 日志文件.open("a", encoding="utf-8") as f:
            f.write(行 + "\n")
        检查并瘦身日志(日志文件, 最大行数=10000, 保留行数=500)
    except Exception:
        pass


# ---- 4) UI 日志桥接：把消息同步到 GUI 日志面板 ----
class UI日志桥:
    """让 GUI 文本框接收 debug_print/记录错误 内容（仅推送，不写文件）。"""

    _回调: Optional[Callable[[str, str], None]] = None

    @classmethod
    def 注册(cls, 回调: Callable[[str, str], None]) -> None:
        """注册 GUI 接收回调，签名 (消息, 级别) -> None。"""
        cls._回调 = 回调

    @classmethod
    def 推送(cls, 消息: str, 级别: str = "INFO") -> None:
        """业务侧调用：把消息推送到 GUI（不直接写日志文件）。"""
        if cls._回调:
            try:
                cls._回调(消息, 级别)
            except Exception:
                pass


# ---- 4.2) UI 便捷函数：同时写日志 + 推送 GUI ----
def UI信息(消息: str) -> None:
    """普通信息：推送 GUI + 调试日志。"""
    UI日志桥.推送(消息, "INFO")
    debug_print(消息)


def UI警告(消息: str) -> None:
    """警告信息：推送 GUI + 调试日志（带 [WARN] 前缀）。"""
    UI日志桥.推送(消息, "WARN")
    debug_print(f"[WARN] {消息}")


def UI错误(消息: str) -> None:
    """错误信息：推送 GUI + 写入 error/LOG.txt。"""
    UI日志桥.推送(消息, "ERROR")
    记录错误(消息)
