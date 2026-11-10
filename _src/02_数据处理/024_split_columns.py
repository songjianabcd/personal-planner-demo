# -*- coding: utf-8 -*-
# ============================================================
# 024_split_columns.py
# 功能：
#   1) 按分隔符分列
#      1.1) 按分隔符分列(数据集对象, 源列, 分隔符=",", 最大列数=0,
#                        新列名前缀="", 保留原列=True, 去除空白=True)
#           - 最大列数=0 → 取所有行的最大切分列数
#           - 新列名 = "{前缀}{序号}"，序号从 1 开始；无前缀=用 "{源列}_part1"
#   2) 按正则分列（多组捕获）
#      2.1) 按正则分列(数据集对象, 源列, 正则, 新列名列表=None,
#                      保留原列=True)
#           - 必须含至少一个捕获组；新列名列表 长度 < 捕获组数 时自动补 part_N
#           - 用 pandas.Series.str.extract 一次性向量化
#   3) 通用：列引用解析、原数据集元信息保留
# ============================================================

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Optional, Sequence

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


# ---- 1.1) 按分隔符分列 ----
def 按分隔符分列(数据集对象,
           源列: str,
           分隔符: str = ",",
           最大列数: int = 0,
           新列名前缀: str = "",
           保留原列: bool = True,
           去除空白: bool = True) -> "数据集":
    """
    把单列按分隔符切分成多列。

    :param 数据集对象: 数据集
    :param 源列: 源列引用
    :param 分隔符: 切分用分隔符（普通字符串，非正则）
    :param 最大列数: >0 时限制输出列数；0 = 自动按数据最长行数
    :param 新列名前缀: 新列名前缀；空 = 用 "{源列}_part"
    :param 保留原列: True 保留源列在原位
    :param 去除空白: 切分后两端是否 strip
    :return: 新数据集
    """
    if 分隔符 == "":
        记录错误("按分隔符分列：分隔符为空，原样返回")
        return _复制数据集(数据集对象, 数据集对象.数据框.copy())

    df = 数据集对象.数据框.copy()
    源列真名 = 解析列引用(源列, df)

    源序列 = _安全取列(df, 源列真名).astype("object").map(_转文本)
    切分序列 = 源序列.map(lambda 文本, 分隔符=分隔符:
                   _切分单格(文本, 分隔符, 去除空白))

    最大 = max((len(项) for 项 in 切分序列 if isinstance(项, list)), default=0)
    if 最大列数 > 0:
        最大 = min(最大, 最大列数)
    if 最大 == 0:
        debug_print("按分隔符分列：未发现可切分内容，原样返回")
        return _复制数据集(数据集对象, df)

    前缀 = 新列名前缀 or f"{源列真名}_part"
    新列名列表 = [_生成不冲突列名(df, f"{前缀}{i+1}") for i in range(最大)]

    切分df = pd.DataFrame(
        [_对齐到最大(项, 最大) for 项 in 切分序列],
        columns=新列名列表,
        index=df.index)

    df = _按位置插入(df, 源列真名, 切分df)
    if not 保留原列:
        df = df.drop(columns=[源列真名])

    debug_print(f"按分隔符分列：源={源列真名} → {新列名列表} | 分隔符={分隔符!r}")
    return _复制数据集(数据集对象, df)


# ---- 2.1) 按正则分列（多组捕获） ----
def 按正则分列(数据集对象,
          源列: str,
          正则: str,
          新列名列表: Optional[Sequence[str]] = None,
          保留原列: bool = True) -> "数据集":
    """
    用正则把单列按捕获组切分到多列。

    :param 数据集对象: 数据集
    :param 源列: 源列引用
    :param 正则: 含至少一个捕获组的正则
    :param 新列名列表: 与捕获组数对应；不足自动补 "part_N"
    :param 保留原列: True 保留源列
    :return: 新数据集
    """
    try:
        编译后 = re.compile(正则)
    except re.error as 异常:
        记录错误(f"按正则分列：正则非法 {正则} | {异常}")
        raise ValueError(f"正则非法：{异常}") from 异常

    if 编译后.groups < 1:
        记录错误("按正则分列：正则必须含至少一个捕获组")
        raise ValueError("正则必须含至少一个捕获组（如 (\\d+)）")

    df = 数据集对象.数据框.copy()
    源列真名 = 解析列引用(源列, df)

    # ---- 生成新列名 ----
    候选 = list(新列名列表) if 新列名列表 else []
    候选 = [str(s).strip() for s in 候选 if str(s).strip()]
    while len(候选) < 编译后.groups:
        候选.append(f"{源列真名}_part{len(候选)+1}")
    候选 = 候选[:编译后.groups]
    最终新列名 = [_生成不冲突列名(df, n) for n in 候选]

    源序列 = _安全取列(df, 源列真名).astype("object").map(_转文本)
    提取df = 源序列.str.extract(正则, expand=True)
    提取df.columns = 最终新列名

    df = _按位置插入(df, 源列真名, 提取df)
    if not 保留原列:
        df = df.drop(columns=[源列真名])

    debug_print(f"按正则分列：源={源列真名} → {最终新列名} | 正则={正则!r}")
    return _复制数据集(数据集对象, df)


# ---- 3.1) 单格转文本（None 保留为空串） ----
def _转文本(值) -> str:
    """None/NaN → ''；其它 str()"""
    if 值 is None:
        return ""
    try:
        if pd.isna(值):
            return ""
    except Exception:
        pass
    return str(值)


# ---- 3.2) 单格切分 ----
def _切分单格(文本: str, 分隔符: str, 去除空白: bool) -> List:
    """按 分隔符 切分单格文本；空文本返回空列表。"""
    if not 文本:
        return []
    片段 = 文本.split(分隔符)
    if 去除空白:
        片段 = [s.strip() for s in 片段]
    return 片段


# ---- 3.3) 切分结果对齐到最大列数（不足补 None） ----
def _对齐到最大(项, 最大: int) -> List:
    """切分结果不足 最大 长度时尾部补 None。"""
    if not isinstance(项, list):
        项 = []
    if len(项) >= 最大:
        return 项[:最大]
    return 项 + [None] * (最大 - len(项))


# ---- 3.4) 不冲突列名 ----
def _生成不冲突列名(数据框: pd.DataFrame, 期望列名: str) -> str:
    """期望列名已存在则追加 _2/_3。"""
    if 期望列名 not in 数据框.columns:
        return 期望列名
    序号 = 2
    while True:
        新名 = f"{期望列名}_{序号}"
        if 新名 not in 数据框.columns:
            return 新名
        序号 += 1


# ---- 3.5) 把多列在源列右侧插入（保持视觉相邻） ----
def _按位置插入(数据框: pd.DataFrame, 锚点列: str,
           新列df: pd.DataFrame) -> pd.DataFrame:
    """把 新列df 插入到 数据框 中 锚点列 的右侧。"""
    位置 = 数据框.columns.get_loc(锚点列) + 1
    左 = 数据框.iloc[:, :位置]
    右 = 数据框.iloc[:, 位置:]
    return pd.concat([左, 新列df, 右], axis=1)


# ---- 3.6) 复制数据集 ----
def _复制数据集(原数据集, 新数据框: pd.DataFrame) -> "数据集":
    """基于原数据集来源信息构造新的 数据集 实例。"""
    return 数据集(数据框=新数据框,
               来源类型=原数据集.来源类型,
               来源路径=原数据集.来源路径,
               工作簿名=原数据集.工作簿名,
               工作表名=原数据集.工作表名,
               表头行号=原数据集.表头行号)
