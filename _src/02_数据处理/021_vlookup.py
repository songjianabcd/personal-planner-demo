# -*- coding: utf-8 -*-
# ============================================================
# 021_vlookup.py
# 功能：
#   1) 类 VLOOKUP 匹配（pandas 向量化实现，支持一对多）
#      1.1) vlookup_合并(左数据集, 右数据集, 左键列, 右键列, 取值列列表,
#                        模式="全部合并", 合并分隔符=",", 去重=True, 新列名前缀="")
#           - 模式："全部合并"  → 把命中右行的值用 分隔符 拼到一格
#                  "第一条"   → 取第一条命中
#                  "最后一条" → 取最后一条命中
#           - 内部走 groupby + agg 一次性向量化，避免逐行循环
#      1.2) _合并值(序列, 分隔符, 去重) → 字符串：把若干值拼为一格
#   2) 边界处理
#      2.1) 左/右键列经 解析列引用 转真实列名
#      2.2) 右表中存在的取值列才会被处理；不存在的列写入 记录错误 并跳过
#      2.3) 合并完成后回写到 数据集.数据框 的副本，原数据集不破坏
# ============================================================

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Sequence

import pandas as pd

# ---- 拉起兄弟模块（基础设施 + 数据读写） ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent)
                )     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402

_log = 加载("04_基础设施", "041_log")
debug_print = _log.debug_print
记录错误 = _log.记录错误

_reader = 加载("01_数据读写", "011_excel_reader")
数据集 = _reader.数据集
解析列引用 = _reader.解析列引用
_安全取列 = _reader._安全取列


# ---- 1.2) 把若干值合并为单元格字符串 ----
def _合并值(序列: Sequence, 分隔符: str, 去重: bool) -> str:
    """
    把一列值拼成单格字符串。

    :param 序列: 待合并的值序列
    :param 分隔符: 分隔符
    :param 去重: 是否按出现顺序去重
    :return: 合并后的单元格文本
    """
    结果: List[str] = []
    已见 = set()
    for 元素 in 序列:
        if 元素 is None:
            continue
        try:
            if pd.isna(元素):
                continue
        except Exception:
            pass
        文本 = str(元素).strip()
        if not 文本:
            continue
        if 去重:
            if 文本 in 已见:
                continue
            已见.add(文本)
        结果.append(文本)
    return 分隔符.join(结果)


# ---- 1.1) 类 VLOOKUP 多值合并 ----
def vlookup_合并(左数据集,
               右数据集,
               左键列: str,
               右键列: str,
               取值列列表: List[str],
               模式: str = "全部合并",
               合并分隔符: str = ",",
               去重: bool = True,
               新列名前缀: str = "") -> "数据集":
    """
    在左表上做类 VLOOKUP 匹配，按右键列把右表多列值合并/取首末写入左表。

    :param 左数据集: 数据集对象（左侧主表）
    :param 右数据集: 数据集对象（右侧匹配表）
    :param 左键列:   左表用于匹配的列引用（列名 / Excel 列字母）
    :param 右键列:   右表用于匹配的列引用
    :param 取值列列表: 右表中要回填的列引用列表
    :param 模式: "全部合并" / "第一条" / "最后一条"
    :param 合并分隔符: 全部合并模式下使用的分隔符
    :param 去重: 全部合并模式下是否去重
    :param 新列名前缀: 写回左表时的新列名前缀（防止冲突）
    :return: 新的数据集（不修改入参）
    :raises ValueError: 列引用解析失败时
    """
    try:
        左df = 左数据集.数据框.copy()
        右df = 右数据集.数据框

        左键真名 = 解析列引用(左键列, 左df)
        右键真名 = 解析列引用(右键列, 右df)

        有效取值列: List[str] = []
        for 列引用 in 取值列列表:
            try:
                真名 = 解析列引用(列引用, 右df)
                有效取值列.append(真名)
            except Exception as 异常:
                记录错误(f"vlookup 跳过取值列 {列引用}：{异常}")

        if not 有效取值列:
            记录错误("vlookup 没有可用的取值列，原样返回左表")
            return _复制数据集(左数据集, 左df)

        # ---- 准备 join 用的右表（key 列改名以匹配左键） ----
        右键列名 = 右键真名
        右子集 = 右df[[右键列名] + 有效取值列].copy()

        # ---- 按模式聚合 ----
        if 模式 == "第一条":
            聚合 = 右子集.groupby(右键列名, as_index=False, sort=False).first()
        elif 模式 == "最后一条":
            聚合 = 右子集.groupby(右键列名, as_index=False, sort=False).last()
        else:                                                          # 全部合并
            模式 = "全部合并"
            分组 = 右子集.groupby(右键列名, as_index=False, sort=False)
            字段聚合 = {列: (lambda 序列, c=列: _合并值(序列, 合并分隔符, 去重))
                    for 列 in 有效取值列}
            聚合 = 分组.agg(字段聚合)

        # ---- 重命名右键列以与左键对齐 ----
        聚合 = 聚合.rename(columns={右键列名: 左键真名})

        # ---- 给取值列加前缀（避免左右同名冲突） ----
        前缀 = 新列名前缀 or ""
        重命名 = {}
        for 列 in 有效取值列:
            目标列名 = f"{前缀}{列}" if 前缀 else 列
            目标列名 = _生成不冲突列名(左df, 目标列名, 排除={左键真名})
            重命名[列] = 目标列名
        聚合 = 聚合.rename(columns=重命名)

        # ---- left join 回左表 ----
        合并键临时 = "__vlookup_左键__"
        左df[合并键临时] = _安全取列(左df, 左键真名).astype("object")
        聚合[合并键临时] = _安全取列(聚合, 左键真名).astype("object")
        聚合 = 聚合.drop(columns=[左键真名])
        合并后 = 左df.merge(聚合, how="left", on=合并键临时)
        合并后 = 合并后.drop(columns=[合并键临时])

        debug_print(f"vlookup 完成：模式={模式} 新列={list(重命名.values())} "
                    f"左行={len(左df)} 右键唯一={聚合.shape[0]}")
        return _复制数据集(左数据集, 合并后)

    except Exception as 异常:
        记录错误(f"vlookup_合并 失败：{异常}")
        raise


# ---- 1.3) 不冲突列名生成 ----
def _生成不冲突列名(数据框: pd.DataFrame, 期望列名: str, 排除: set) -> str:
    """若期望列名在数据框中已存在（且不在排除集合内），追加 _2/_3 直到不冲突。"""
    if 期望列名 not in 数据框.columns or 期望列名 in 排除:
        return 期望列名
    序号 = 2
    while True:
        新名 = f"{期望列名}_{序号}"
        if 新名 not in 数据框.columns:
            return 新名
        序号 += 1


# ---- 1.4) 复制数据集（保留来源元信息） ----
def _复制数据集(原数据集, 新数据框: pd.DataFrame) -> "数据集":
    """基于原数据集来源信息构造新的 数据集 实例。"""
    return 数据集(数据框=新数据框,
               来源类型=原数据集.来源类型,
               来源路径=原数据集.来源路径,
               工作簿名=原数据集.工作簿名,
               工作表名=原数据集.工作表名,
               表头行号=原数据集.表头行号)
