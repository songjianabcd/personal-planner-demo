# -*- coding: utf-8 -*-
# ============================================================
# 022_regex_extract.py
# 功能：
#   1) 正则提取（在指定列旁追加新列）
#      1.1) 正则提取(数据集对象, 源列, 正则, 目标列名=None, 捕获组=1, 全部命中=False, 合并分隔符=",")
#           - 全部命中=False：取第 捕获组 个分组的第一次命中
#           - 全部命中=True ：findall 全部，按 合并分隔符 拼接
#           - 目标列名=None  → 自动生成 "{源列}_提取"，与已有列冲突自动加 _2/_3
#      1.2) 内部用 pandas.Series.str.extract / str.findall 走向量化
#   2) 容错
#      2.1) 正则编译失败 → 记录错误 并抛 ValueError
#      2.2) 列引用解析失败 → 记录错误 并抛 ValueError
#      2.3) 单元格非字符串 → 自动 str() 后再匹配
# ============================================================

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Optional

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
解析目标列名 = _reader.解析目标列名
_安全取列 = _reader._安全取列


# ---- 1.1) 正则提取主函数 ----
def 正则提取(数据集对象,
         源列: str,
         正则: str,
         目标列名: Optional[str] = None,
         捕获组: int = 1,
         全部命中: bool = False,
         合并分隔符: str = ",",
         覆盖同名列: bool = False) -> "数据集":
    """
    在数据集中为 源列 提取一个结果列。

    放置规则：
        - 指定了目标列名 → 直接覆盖/追加到该列
        - 未指定目标列名 → 在最大列+3 处插入（前面留 2 个空列），列名自动生成

    :param 数据集对象: 数据集
    :param 源列: 源列引用（列名/列字母）
    :param 正则: Python 正则表达式（建议含至少一个捕获组）
    :param 目标列名: 新列名；None/空 = 自动生成并按规则放置
    :param 捕获组: 取第几个捕获组（1-based），无捕获组时填 0=整个匹配
    :param 全部命中: True=findall 全部并合并；False=只取第一次命中
    :param 合并分隔符: 全部命中模式下的分隔符
    :return: 新的数据集
    :raises ValueError: 正则非法 / 源列不存在
    """
    try:
        编译后 = re.compile(正则)
    except re.error as 异常:
        记录错误(f"正则编译失败：{正则} | {异常}")
        raise ValueError(f"正则非法：{异常}") from 异常

    df = 数据集对象.数据框.copy()
    源列真名 = 解析列引用(源列, df)

    解析后目标, 是列字母 = 解析目标列名(目标列名, df)
    期望列名 = 解析后目标 or f"{源列真名}_提取"
    # 用户指定了有效目标列，且（显式要求覆盖 或 输入的是 Excel 列字母）→ 直接覆盖
    if 解析后目标 and (覆盖同名列 or 是列字母):
        新列名 = 期望列名
    else:
        新列名 = _生成不冲突列名(df, 期望列名)

    源序列 = _安全取列(df, 源列真名).astype("object").map(_单元格转文本)

    if 全部命中:
        结果序列 = 源序列.map(lambda 文本: _全部命中拼接(编译后, 文本, 捕获组, 合并分隔符))
    else:
        结果序列 = 源序列.map(lambda 文本: _首次命中(编译后, 文本, 捕获组))

    if 解析后目标:
        # 指定了目标列：直接覆盖/追加
        df[新列名] = 结果序列
    else:
        # 未指定目标列：最大列+3 处插入，前面留 2 个空列
        df.insert(len(df.columns), " ", "", allow_duplicates=True)
        df.insert(len(df.columns), "  ", "", allow_duplicates=True)
        df.insert(len(df.columns), 新列名, 结果序列)

    debug_print(f"正则提取完成：列={源列真名} → {新列名} | 全部命中={全部命中} | 捕获组={捕获组}")
    return _复制数据集(数据集对象, df)


# ---- 1.2) 单元格转文本：保留原值；空值返回空串 ----
def _单元格转文本(值) -> str:
    """把任意单元格值转为字符串，空值视为空串。"""
    if 值 is None:
        return ""
    try:
        if pd.isna(值):
            return ""
    except Exception:
        pass
    return str(值)


# ---- 1.3) 首次命中：返回第 捕获组 个分组（0=整个匹配） ----
def _首次命中(编译后, 文本: str, 捕获组: int) -> str:
    """返回第一次命中的指定捕获组文本，没命中返回空串。"""
    if not 文本:
        return ""
    匹配 = 编译后.search(文本)
    if 匹配 is None:
        return ""
    try:
        return 匹配.group(捕获组) or ""
    except IndexError:
        return 匹配.group(0) or ""


# ---- 1.4) 全部命中：findall + 拼接 ----
def _全部命中拼接(编译后, 文本: str, 捕获组: int, 分隔符: str) -> str:
    """findall 全部命中后按 分隔符 拼接；多分组时取指定下标。"""
    if not 文本:
        return ""
    命中列表 = 编译后.findall(文本)
    if not 命中列表:
        return ""
    片段 = []
    for 项 in 命中列表:
        if isinstance(项, tuple):
            下标 = max(0, 捕获组 - 1) if 捕获组 >= 1 else 0
            if 下标 < len(项):
                片段.append(str(项[下标]))
            elif 项:
                片段.append(str(项[0]))
        else:
            片段.append(str(项))
    return 分隔符.join(s for s in 片段 if s)


# ---- 1.5) 不冲突列名 ----
def _生成不冲突列名(数据框: pd.DataFrame, 期望列名: str) -> str:
    """期望列名已存在则追加 _2/_3 直到不冲突。"""
    if 期望列名 not in 数据框.columns:
        return 期望列名
    序号 = 2
    while True:
        新名 = f"{期望列名}_{序号}"
        if 新名 not in 数据框.columns:
            return 新名
        序号 += 1


# ---- 1.6) 复制数据集（保留元信息） ----
def _复制数据集(原数据集, 新数据框: pd.DataFrame) -> "数据集":
    """基于原数据集来源信息构造新的 数据集 实例。"""
    return 数据集(数据框=新数据框,
               来源类型=原数据集.来源类型,
               来源路径=原数据集.来源路径,
               工作簿名=原数据集.工作簿名,
               工作表名=原数据集.工作表名,
               表头行号=原数据集.表头行号)
