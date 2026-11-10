# -*- coding: utf-8 -*-
# ============================================================
# 023_replace.py
# 功能：
#   1) 单对查找替换
#      1.1) 查找替换(数据集对象, 列列表, 查找文本, 替换文本, 使用正则=False, 区分大小写=True, 全词匹配=False)
#           - 列列表 = [] / None → 对所有文本类列执行
#           - 使用正则=True：用 re.sub；False：纯字符串替换
#           - 全词匹配=True：自动加 \b...\b 边界
#   2) 对照表批量替换（旧值→新值 字典）
#      2.1) 对照表替换(数据集对象, 列列表, 对照表: dict, 整格匹配=True, 使用正则=False)
#           - 整格匹配=True：单元格全文等于"旧值"才替换为"新值"（默认）
#           - 整格匹配=False：在单元格内做子串替换；正则模式下旧值即正则
#   3) 通用工具
#      3.1) _应用到列列表()：解析列引用列表，没指定就遍历所有 object/字符串列
#      3.2) _复制数据集()：基于原数据集元信息构造新数据集
# ============================================================

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd

# ---- 拉起兄弟模块 ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _bootstrap import 加载                                          # noqa: E402

_log = 加载("04_基础设施", "041_log")
debug_print = _log.debug_print
记录错误 = _log.记录错误

_reader = 加载("01_数据读写", "011_excel_reader")
数据集 = _reader.数据集
解析列引用 = _reader.解析列引用
_安全取列 = _reader._安全取列


# ---- 1.1) 单对查找替换 ----
def 查找替换(数据集对象,
         列列表: Optional[Sequence[str]],
         查找文本: str,
         替换文本: str,
         使用正则: bool = False,
         区分大小写: bool = True,
         全词匹配: bool = False) -> "数据集":
    """
    在指定列上执行查找→替换；支持普通文本和正则两种。

    :param 数据集对象: 数据集
    :param 列列表: 列引用列表；None/空 = 全部文本列
    :param 查找文本: 要查找的内容
    :param 替换文本: 替换为
    :param 使用正则: True=按正则；False=按纯文本
    :param 区分大小写: 默认 True
    :param 全词匹配: True=加 \\b 边界（仅对正则/英文有效）
    :return: 新数据集
    """
    df = 数据集对象.数据框.copy()
    实际列列表 = _应用到列列表(df, 列列表)
    if not 实际列列表:
        记录错误("查找替换：没有可用列，原样返回")
        return _复制数据集(数据集对象, df)

    模式 = _构造正则模式(查找文本, 使用正则, 区分大小写, 全词匹配)
    替换计数 = 0
    for 列 in 实际列列表:
        序列 = _安全取列(df, 列).astype("object").map(_转文本_保留None)
        替换后 = 序列.map(lambda 文本, 模式=模式, 替换文本=替换文本:
                     _执行替换(文本, 模式, 替换文本))
        替换计数 += int((序列 != 替换后).sum())
        df[列] = 替换后

    debug_print(f"查找替换完成：列={实际列列表} | 命中替换={替换计数} | 正则={使用正则}")
    return _复制数据集(数据集对象, df)


# ---- 2.1) 对照表批量替换 ----
def 对照表替换(数据集对象,
          列列表: Optional[Sequence[str]],
          对照表: Dict[str, str],
          整格匹配: bool = True,
          使用正则: bool = False,
          区分大小写: bool = True) -> "数据集":
    """
    用 旧值→新值 字典批量替换；可选整格匹配 / 子串替换 / 正则。

    :param 数据集对象: 数据集
    :param 列列表: 列引用列表；None/空 = 全部文本列
    :param 对照表: {旧值: 新值, ...}
    :param 整格匹配: True=单元格等于旧值才替换；False=子串替换
    :param 使用正则: True=按正则；False=按纯文本（仅子串模式有效）
    :param 区分大小写: 仅子串模式有效
    :return: 新数据集
    """
    if not 对照表:
        记录错误("对照表为空，原样返回")
        return _复制数据集(数据集对象, 数据集对象.数据框.copy())

    df = 数据集对象.数据框.copy()
    实际列列表 = _应用到列列表(df, 列列表)
    if not 实际列列表:
        记录错误("对照表替换：没有可用列，原样返回")
        return _复制数据集(数据集对象, df)

    替换计数 = 0
    if 整格匹配:
        # ---- 整格匹配：序列级 replace（一次性） ----
        for 列 in 实际列列表:
            原 = _安全取列(df, 列)
            df[列] = 原.replace(对照表)
            替换计数 += int((原.astype("object") !=
                        _安全取列(df, 列).astype("object")).sum())
    else:
        # ---- 子串替换：逐对 旧→新 顺序应用 ----
        for 旧值, 新值 in 对照表.items():
            模式 = _构造正则模式(str(旧值), 使用正则, 区分大小写, 全词匹配=False)
            for 列 in 实际列列表:
                序列 = _安全取列(df, 列).astype("object").map(_转文本_保留None)
                替换后 = 序列.map(lambda 文本, 模式=模式, 新值=str(新值):
                             _执行替换(文本, 模式, 新值))
                替换计数 += int((序列 != 替换后).sum())
                df[列] = 替换后

    debug_print(f"对照表替换完成：列={实际列列表} | 条目={len(对照表)} | "
                f"整格匹配={整格匹配} | 命中替换={替换计数}")
    return _复制数据集(数据集对象, df)


# ---- 3.1) 列列表展开：未指定 → 所有 object/字符串列 ----
def _应用到列列表(数据框: pd.DataFrame,
            列列表: Optional[Sequence[str]]) -> List[str]:
    """解析列引用列表；空 → 所有 object 列。"""
    if not 列列表:
        return [c for c in 数据框.columns if 数据框[c].dtype == object]
    结果: List[str] = []
    for 引用 in 列列表:
        try:
            结果.append(解析列引用(引用, 数据框))
        except Exception as 异常:
            记录错误(f"替换跳过列 {引用}：{异常}")
    return 结果


# ---- 3.2) 转文本（None 保留） ----
def _转文本_保留None(值):
    """None/NaN 保留为 None；其它转 str。"""
    if 值 is None:
        return None
    try:
        if pd.isna(值):
            return None
    except Exception:
        pass
    return str(值)


# ---- 3.3) 构造正则模式 ----
def _构造正则模式(查找文本: str, 使用正则: bool,
            区分大小写: bool, 全词匹配: bool):
    """
    构造编译后的正则对象。
    使用正则=False 时把查找文本 re.escape 后再编译。
    """
    if not 查找文本:
        return None
    模式文本 = 查找文本 if 使用正则 else re.escape(查找文本)
    if 全词匹配:
        模式文本 = rf"\b(?:{模式文本})\b"
    标志 = 0 if 区分大小写 else re.IGNORECASE
    try:
        return re.compile(模式文本, 标志)
    except re.error as 异常:
        记录错误(f"替换正则非法：{查找文本} | {异常}")
        return None


# ---- 3.4) 执行替换：单格文本 ----
def _执行替换(文本, 模式, 替换文本: str):
    """单单元格执行 sub。None 保留 None。"""
    if 文本 is None:
        return None
    if 模式 is None:
        return 文本
    try:
        return 模式.sub(替换文本, 文本)
    except Exception:
        return 文本


# ---- 3.5) 复制数据集 ----
def _复制数据集(原数据集, 新数据框: pd.DataFrame) -> "数据集":
    """基于原数据集来源信息构造新的 数据集 实例。"""
    return 数据集(数据框=新数据框,
               来源类型=原数据集.来源类型,
               来源路径=原数据集.来源路径,
               工作簿名=原数据集.工作簿名,
               工作表名=原数据集.工作表名,
               表头行号=原数据集.表头行号)
