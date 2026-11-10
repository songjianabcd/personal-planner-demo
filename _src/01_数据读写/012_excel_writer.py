# -*- coding: utf-8 -*-
# ============================================================
# 012_excel_writer.py
# 功能：
#   1) 写出入口
#      1.1) 写出_新建文件(数据集, 路径, 保留表头)：.xls 走 xlwings；.xlsx 走 pandas
#      1.2) 写出_当前激活簿(数据集, 表名, 保留表头, 起始格, 先清空)：xlwings 整片写回，保留 Excel 格式
#   2) 私有写出实现
#      2.1) _写出_xls_via_xlwings：xlwings 新建簿 + SaveAs(FileFormat=56)
#      2.2) _df转二维：DataFrame → 二维 list（含/不含表头）
#   3) 大量写入约束（excel_rules.md）
#      - 关闭 screen_updating / calculation 提速
#      - 整片 range("A1").value = 二维列表，禁止逐格
# ============================================================

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

import pandas as pd

# ---- 拉起基础设施 + 兄弟模块（拿到 数据集 类型） ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402
_log = 加载("04_基础设施", "041_log")
debug_print = _log.debug_print

_reader = 加载("01_数据读写", "011_excel_reader")
数据集 = _reader.数据集                                               # 复用读模块定义的类型


# ---- 1.1) 写出到新建文件 ----
def 写出_新建文件(数据集对象: 数据集,
                  目标路径: str,
                  保留表头: bool = True) -> str:
    """
    把数据集另存到指定路径。.xls 用 xlwings，其它默认 .xlsx + pandas。

    :param 数据集对象: 待写出的 数据集
    :param 目标路径:   绝对/相对路径，扩展名决定写法
    :param 保留表头:   是否在第一行写出列名
    :return: 实际写出的路径字符串
    """
    路径 = Path(目标路径)
    路径.parent.mkdir(parents=True, exist_ok=True)
    后缀 = 路径.suffix.lower()
    if 后缀 == ".xls":
        return _写出_xls_via_xlwings(数据集对象, 路径, 保留表头)
    if 后缀 not in (".xlsx", ".xlsm"):
        路径 = 路径.with_suffix(".xlsx")
    数据集对象.数据框.to_excel(路径, index=False, header=保留表头,
                                sheet_name=(数据集对象.工作表名 or "Sheet1")[:31])
    debug_print(f"另存为成功：{路径}")
    return str(路径)


# ---- 2.1) .xls 写出（xlwings 唯一可写 .xls 的方案） ----
def _写出_xls_via_xlwings(数据集对象: 数据集, 路径: Path, 保留表头: bool) -> str:
    """以隐藏 Excel 实例新建工作簿，整片写入并 SaveAs(FileFormat=56) 为 .xls。"""
    import xlwings as xw
    应用 = None
    工作簿 = None
    try:
        应用 = xw.App(visible=False, add_book=False)
        应用.screen_updating = False                                  # 关屏幕刷新提速
        应用.calculation = "manual"                                   # 关自动计算提速
        工作簿 = 应用.books.add()
        工作表 = 工作簿.sheets[0]
        if 数据集对象.工作表名:
            工作表.name = (数据集对象.工作表名 or "Sheet1")[:31]
        二维 = _df转二维(数据集对象.数据框, 保留表头)
        if 二维:
            工作表.range("A1").value = 二维                            # 整片写入，禁止逐格
        工作簿.api.SaveAs(str(路径), FileFormat=56)                    # 56 = xlExcel8 (.xls)
        debug_print(f"xlwings 另存 .xls 成功：{路径}")
        return str(路径)
    finally:
        try:
            if 工作簿 is not None:
                工作簿.close()
        except Exception:
            pass
        try:
            if 应用 is not None:
                应用.calculation = "automatic"
                应用.screen_updating = True
                应用.quit()
        except Exception:
            pass


# ---- 1.2) 写出到当前激活簿 ----
def 写出_当前激活簿(数据集对象: 数据集,
                    目标表名: Optional[str] = None,
                    保留表头: bool = True,
                    起始单元格: str = "A1",
                    先清空: bool = True) -> str:
    """
    写回当前激活 Excel 的指定/激活 sheet（保留 Excel 进程中的格式）。

    :param 数据集对象: 待写出的 数据集
    :param 目标表名:   sheet 名，不存在则自动新建；None = 激活 sheet
    :param 保留表头:   是否在第一行写出列名
    :param 起始单元格: 写入起点，默认 A1
    :param 先清空:     写入前是否先 clear_contents
    :return: 形如 "工作簿.xlsx / Sheet1" 的描述串
    :raises RuntimeError: 没有激活 Excel / 工作簿
    """
    import xlwings as xw
    应用 = xw.apps.active
    if 应用 is None:
        raise RuntimeError("没有检测到正在运行的 Excel，请先打开 Excel。")
    工作簿 = 应用.books.active
    if 工作簿 is None:
        raise RuntimeError("当前 Excel 没有激活的工作簿。")

    # ---- 选择/创建目标 sheet ----
    if 目标表名:
        现有 = [s.name for s in 工作簿.sheets]
        if 目标表名 in 现有:
            工作表 = 工作簿.sheets[目标表名]
        else:
            工作表 = 工作簿.sheets.add(name=目标表名[:31])
    else:
        工作表 = 工作簿.sheets.active

    应用.screen_updating = False
    应用.calculation = "manual"
    try:
        if 先清空:
            工作表.clear_contents()
        二维 = _df转二维(数据集对象.数据框, 保留表头)
        if 二维:
            工作表.range(起始单元格).value = 二维                       # 整片写入
        工作簿.save()
    finally:
        应用.screen_updating = True
        应用.calculation = "automatic"
    debug_print(f"写回激活簿成功：{工作簿.name} / {工作表.name}")
    return f"{工作簿.name} / {工作表.name}"


# ---- 2.2) DataFrame → 二维 list ----
def _df转二维(数据框: pd.DataFrame, 保留表头: bool) -> list:
    """
    把 DataFrame 转为二维 list，便于 xlwings 整片写入。

    :param 数据框: pandas DataFrame
    :param 保留表头: True = 首行写列名
    :return: list[list]，可能为空
    """
    if 数据框 is None or 数据框.empty:
        return [list(数据框.columns)] if (保留表头 and 数据框 is not None) else []
    数据 = 数据框.where(pd.notna(数据框), None).values.tolist()
    if 保留表头:
        return [list(数据框.columns)] + 数据
    return 数据
