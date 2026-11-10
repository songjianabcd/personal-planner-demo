# -*- coding: utf-8 -*-
# ============================================================
# 011_excel_reader.py
# 功能：
#   1) 类型与基础工具
#      1.1) 数据集 类：统一的内存数据载体（DataFrame + 来源元信息）
#      1.2) 解析列引用(引用, 数据框)：支持表头名 或 Excel 列字母（A/B/AA），返回真实列名
#   2) 读取入口分流（按扩展名 / 来源选择库）
#      2.1) 读取_文件(路径, 表名, 表头行号)：.xls 走 xlwings；.xlsx/.xlsm 走 pandas
#      2.2) 读取_激活簿(表名, 表头行号)：xlwings 取 apps.active.books.active
#   3) 私有读取实现
#      3.1) _读取_xlsx_via_pandas：pandas.read_excel + dtype=object（保留原始类型）
#      3.2) _读取_xls_via_xlwings：xlwings 一次性 expand("table") 整片读取
#      3.3) _二维列表转df：把 xlwings 二维 list 按 表头行号 切分构造 DataFrame
#   4) Sheet 列表
#      4.1) 列出文件表名(路径)：.xls 走 xlwings，其它走 pandas.ExcelFile
#      4.2) 列出激活簿表名()：xlwings 直接读
# ============================================================

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

# ---- 拉起基础设施（debug_print / 记录错误） ----
sys.path.insert(0, str(Path(__file__).resolve().parent.parent)
                )     # 把 src 加入 sys.path
from _bootstrap import 加载                                          # noqa: E402
_log = 加载("04_基础设施", "041_log")
debug_print = _log.debug_print
记录错误 = _log.记录错误


# ---- 0) 表头行识别配置 ----
默认表头关键字: set = {
    "集团", "物料", "品目", "部件名称", "物料号", "物料号长参数表",
    "工厂", "公司代码", "采购组", "MRP控制者", "库存地点",
    "计量单位", "物料描述", "物料类型", "行业领域",
}
识别扫描行数: int = 30


def _识别表头行号(二维: list, 默认值: int = 1,
            关键字集合: Optional[set] = None,
            扫描行数: int = 识别扫描行数) -> int:
    """
    扫描二维列表前 N 行，找出含关键字最多（≥ 2）的行作为表头行号（1-based）。
    命中不足 2 个时返回 默认值，允许用户手动覆盖。
    """
    关键字 = 关键字集合 if 关键字集合 is not None else 默认表头关键字
    if not 二维 or not isinstance(二维, list):
        return 默认值
    上限 = min(len(二维), max(1, 扫描行数))
    最佳行号 = 默认值
    最佳命中 = 0
    for 行索引 in range(上限):
        行 = 二维[行索引]
        if not isinstance(行, (list, tuple)):
            行 = [行]
        命中 = 0
        for 单元格 in 行:
            if 单元格 is None:
                continue
            文本 = str(单元格).strip()
            if 文本 in 关键字:
                命中 += 1
        if 命中 >= 2 and 命中 > 最佳命中:
            最佳命中 = 命中
            最佳行号 = 行索引 + 1     # 1-based
    debug_print(f"表头识别：最佳行={最佳行号} | 命中={最佳命中} | 默认={默认值} | 扫描={上限}行")
    return 最佳行号


# ---- 1.1) 数据集类：统一的内存数据载体 ----
class 数据集:
    """
    统一的内存数据载体（pandas DataFrame）+ 来源元信息。

    属性：
        数据框    : pandas.DataFrame
        来源类型  : "file" / "active"
        来源路径  : 原始文件全路径；激活簿时为 fullname
        工作簿名  : 文件名（含扩展名）
        工作表名  : sheet 名
        表头行号  : 1-based，第几行作为表头
    """

    def __init__(self,
                 数据框: pd.DataFrame,
                 来源类型: str = "file",
                 来源路径: Optional[str] = None,
                 工作簿名: Optional[str] = None,
                 工作表名: Optional[str] = None,
                 表头行号: int = 1) -> None:
        """初始化数据集；不做任何 IO 操作，由读取函数构造。"""
        self.数据框 = 数据框
        self.来源类型 = 来源类型
        self.来源路径 = 来源路径
        self.工作簿名 = 工作簿名
        self.工作表名 = 工作表名
        self.表头行号 = 表头行号

    @property
    def 列名列表(self) -> List[str]:
        """返回列名字符串列表。"""
        return [str(c) for c in self.数据框.columns]

    @property
    def 行数(self) -> int:
        """数据框行数（不含表头）。"""
        return len(self.数据框)

    @property
    def 列数(self) -> int:
        """数据框列数。"""
        return self.数据框.shape[1]

    def 摘要(self) -> str:
        """返回一行人类可读摘要，便于日志展示。"""
        来源 = self.来源路径 or self.工作簿名 or "未知"
        return (f"来源类型={self.来源类型} | 簿={来源} | "
                f"表={self.工作表名} | 表头行={self.表头行号} | "
                f"行={self.行数} 列={self.列数}")


# ---- 1.2) 解析列引用：支持表头名 / Excel 列字母 / 中文关键字 ----
_列引用预处理正则 = re.compile(r"^[第\s]*(\d+)[\s列]*$")


def _预处理列引用(文本: str) -> str:
    """
    预处理用户输入的列引用，统一常见写法：
      - 全角数字、全角空格 → 半角
      - "第3列" / "第 3 列" / "3列" / "３" → "3"
    """
    半角 = "".join(
        chr(ord(ch) - 0xFEE0) if 0xFF10 <= ord(ch) <= 0xFF19 else
        " " if ch == "\u3000" else ch
        for ch in 文本
    ).strip()
    匹配 = _列引用预处理正则.match(半角)
    if 匹配:
        return 匹配.group(1)
    return 半角


def _标准化列名文本(文本: str) -> str:
    """去掉首尾空白，并把内部连续空白压缩为单个空格，便于模糊匹配。"""
    return " ".join(str(文本).split())


def _按内容查找列(标准文本: str, 数据框: pd.DataFrame, 扫描行数: int = 5) -> Optional[str]:
    """
    在数据框前 N 行中查找包含 标准文本 的单元格，返回唯一命中的列名。
    用于表头未被识别为列名时，按用户输入的表头关键字定位列。
    """
    if 数据框.empty:
        return None
    上限 = min(扫描行数, len(数据框))
    命中列: set = set()
    for 行索引 in range(上限):
        行 = 数据框.iloc[行索引]
        for 列名, 值 in 行.items():
            if 值 is None:
                continue
            try:
                if pd.isna(值):
                    continue
            except Exception:
                pass
            单元格文本 = _标准化列名文本(str(值))
            if 标准文本 == 单元格文本 or 标准文本 in 单元格文本:
                命中列.add(列名)
                break
    if len(命中列) == 1:
        return next(iter(命中列))
    return None


def 解析列引用(列引用: str, 数据框: pd.DataFrame) -> str:
    """
    解析"列引用"字符串到真实列名。

    优先级：
        1) 与列名完全一致 → 直接返回
        2) 标准化（去空白、压缩空格）后匹配列名
        3) 忽略大小写匹配列名
        4) 纯数字 → 当作 1-based 列号
        5) 全为 A-Z 字母 → 当作 Excel 列字母（A=0, B=1, AA=26）
        6) 子串/包含匹配列名（仅唯一命中时返回）
        7) 在前 N 行单元格内容中匹配（仅唯一命中时返回）

    :param 列引用: 列名、列号、Excel 列字母 或 列名关键字
    :param 数据框: 目标 DataFrame
    :return: 真实列名
    :raises ValueError: 找不到列或存在歧义时抛出
    """
    if 列引用 is None:
        raise ValueError("列引用不能为空")
    s = _预处理列引用(str(列引用).strip())
    if not s:
        raise ValueError("列引用不能为空")

    列名列表 = list(数据框.columns)
    if not 列名列表:
        raise ValueError(f"找不到列：{列引用}（数据框没有任何列）")

    # 1) 原始精确匹配
    if s in 列名列表:
        return s

    # 2) 标准化后精确匹配
    标准s = _标准化列名文本(s)
    标准列名映射 = {str(c): _标准化列名文本(c) for c in 列名列表}
    for 原名, 标准列名 in 标准列名映射.items():
        if 标准列名 == 标准s:
            return 原名

    # 3) 忽略大小写匹配
    s_lower = s.lower()
    for 原名 in 列名列表:
        if str(原名).lower() == s_lower:
            return 原名

    # 4) 纯数字 → 1-based 列号
    if s.isdigit():
        索引 = int(s) - 1
        if 0 <= 索引 < len(列名列表):
            return 列名列表[索引]

    # 5) Excel 列字母
    if _是列字母(s):
        索引 = 0
        for ch in s.upper():
            索引 = 索引 * 26 + (ord(ch) - ord('A') + 1)
        索引 -= 1
        if 0 <= 索引 < len(列名列表):
            return 列名列表[索引]

    # 6) 子串匹配列名
    子串匹配 = [
        原名 for 原名, 标准列名 in 标准列名映射.items()
        if 标准s in 标准列名 or 标准列名 in 标准s
    ]
    if len(子串匹配) == 1:
        return 子串匹配[0]
    if len(子串匹配) > 1:
        raise ValueError(
            f"列引用存在歧义：{列引用} 匹配到多个列（{', '.join(str(c) for c in 子串匹配)}）")

    # 7) 按前 N 行内容查找
    内容匹配 = _按内容查找列(标准s, 数据框)
    if 内容匹配 is not None:
        return 内容匹配

    raise ValueError(
        f"找不到列：{列引用}（可用列：{', '.join(str(c) for c in 列名列表)}）")


def _是列字母(s: str) -> bool:
    """判断字符串是否全为 A-Z 字母（Excel 列字母）。"""
    return bool(s) and all('A' <= ch.upper() <= 'Z' for ch in s)


def 解析目标列名(目标列名: Optional[str], 数据框: pd.DataFrame) -> Tuple[Optional[str], bool]:
    """
    把用户输入的"目标列名"解析为 DataFrame 中的真实列名。

    返回值：(
        真实列名（已存在的列）或用户输入的原字符串（新列名）或 None,
        原输入是否为 Excel 列字母（A/AA 等）
    )

    解析规则（列字母优先，且绕过 解析列引用 以避免其“列名优先”干扰）：
        - 留空/None → (None, False)
        - 是 Excel 列字母且在范围内 → (该位置真实列名, True)  ★优先
        - 是 Excel 列字母但超出范围 → (原字符串, True)
        - 与现有列名完全一致 → (该列名, False)
        - 其它全新列名 → (原字符串, False)
    """
    if 目标列名 is None:
        return None, False
    s = str(目标列名).strip()
    if not s:
        return None, False
    # ★ 列字母直接算索引，不走 解析列引用（后者会优先匹配列名“Z”而非位置）
    if _是列字母(s):
        索引 = 0
        for ch in s.upper():
            索引 = 索引 * 26 + (ord(ch) - ord('A') + 1)
        索引 -= 1
        if 0 <= 索引 < len(数据框.columns):
            return 数据框.columns[索引], True
        return s, True
    if s in 数据框.columns:
        return s, False
    return s, False


def _安全取列(数据框: pd.DataFrame, 列名: str) -> pd.Series:
    """
    安全取单列。SAP 导出表常见重名列，此时 df[列名] 会返回 DataFrame，
    本函数统一取第一列并返回 Series，避免后续向量化操作报错。
    """
    取出 = 数据框[列名]
    if isinstance(取出, pd.DataFrame):
        return 取出.iloc[:, 0]
    return 取出


# ---- 2.1) 文件读取入口：按扩展名分流 ----
def 读取_文件(文件路径: str,
          工作表名: Optional[str] = None,
          表头行号: int = 1) -> 数据集:
    """
    从外部 Excel 文件读取数据。

    :param 文件路径: .xls / .xlsx / .xlsm 路径
    :param 工作表名: sheet 名，None = 第一张
    :param 表头行号: 1-based，第几行作为表头
    :return: 数据集
    :raises FileNotFoundError: 文件不存在
    :raises ValueError: 不支持的扩展名
    """
    路径 = Path(文件路径)
    if not 路径.exists():
        raise FileNotFoundError(f"文件不存在：{文件路径}")
    后缀 = 路径.suffix.lower()
    if 后缀 == ".xls":
        return _读取_xls_via_xlwings(路径, 工作表名, 表头行号)         # .xls 必须 xlwings
    if 后缀 in (".xlsx", ".xlsm"):
        return _读取_xlsx_via_pandas(路径, 工作表名, 表头行号)
    raise ValueError(f"不支持的扩展名：{后缀}（仅支持 .xls/.xlsx/.xlsm）")


# ---- 3.1) .xlsx/.xlsm 走 pandas（无 Excel 进程依赖、速度快） ----
def _读取_xlsx_via_pandas(路径: Path, 工作表名: Optional[str], 表头行号: int) -> 数据集:
    """用 pandas.read_excel 整片读取 .xlsx；先不带 header 读全，再识别表头行。"""
    完整df = pd.read_excel(路径, sheet_name=(工作表名 or 0),
                         header=None, dtype=object)
    完整df = 完整df.where(pd.notna(完整df), None)
    二维 = 完整df.values.tolist()
    最终表头行 = _识别表头行号(二维, 默认值=表头行号)
    数据框 = _二维列表转df(二维, 最终表头行)
    实际表名 = 工作表名 or _首个表名_pandas(路径)
    debug_print(f"pandas 读取成功：{路径.name} / {实际表名} / {数据框.shape} / 表头行={最终表头行}")
    return 数据集(数据框=数据框,
               来源类型="file",
               来源路径=str(路径),
               工作簿名=路径.name,
               工作表名=实际表名,
               表头行号=最终表头行)


def _首个表名_pandas(路径: Path) -> str:
    """读取 .xlsx 第一个 sheet 名；失败回退默认值。"""
    try:
        return pd.ExcelFile(路径).sheet_names[0]
    except Exception:
        return "Sheet1"


# ---- 3.2) .xls 走 xlwings（excel_rules.md 强制要求） ----
def _读取_xls_via_xlwings(路径: Path, 工作表名: Optional[str], 表头行号: int) -> 数据集:
    """以隐藏 Excel 实例打开 .xls，整片读取后立即关闭。"""
    import xlwings as xw
    应用 = None
    工作簿 = None
    try:
        应用 = xw.App(visible=False, add_book=False)
        应用.screen_updating = False
        工作簿 = 应用.books.open(str(路径), read_only=True, update_links=False)
        工作表 = 工作簿.sheets[工作表名] if 工作表名 else 工作簿.sheets[0]
        实际表名 = 工作表.name
        二维 = _整片读取从_A1(工作表)                            # 整片读取，禁止逐格
        最终表头行 = _识别表头行号(二维, 默认值=表头行号)
        数据框 = _二维列表转df(二维, 最终表头行)
        debug_print(
            f"xlwings 读取 .xls：{路径.name} / {实际表名} / {数据框.shape} / 表头行={最终表头行}")
        return 数据集(数据框=数据框,
                   来源类型="file",
                   来源路径=str(路径),
                   工作簿名=路径.name,
                   工作表名=实际表名,
                   表头行号=最终表头行)
    finally:
        try:
            if 工作簿 is not None:
                工作簿.close()
        except Exception:
            pass
        try:
            if 应用 is not None:
                应用.screen_updating = True
                应用.quit()
        except Exception:
            pass


# ---- 3.2.x) 整片读取辅助：used_range 计算边界 → 从A1一次拿全 ----
def _整片读取从_A1(工作表) -> list:
    """
    先读 used_range 拿到最后一个单元格，再从 A1 到 该点 一次性读。
    避免 range("A1").expand("table") 在 SAP 导出表上遇到空桌/合并单元格 提前截断。
    """
    最后行, 最后列 = 0, 0
    try:
        ur = 工作表.used_range
        最后行 = int(ur.last_cell.row)
        最后列 = int(ur.last_cell.column)
    except Exception:
        pass
    if 最后行 < 1 or 最后列 < 1:
        try:
            备选 = 工作表.range("A1").expand("table").value
        except Exception:
            备选 = None
        return 备选 if 备选 is not None else []
    二维 = 工作表.range((1, 1), (最后行, 最后列)).value
    return 二维 if 二维 is not None else []


# ---- 2.2) 激活簿读取入口 ----
def 读取_激活簿(工作表名: Optional[str] = None,
           表头行号: int = 1) -> 数据集:
    """
    读取当前已打开 Excel 中激活的工作簿/表。
    :raises RuntimeError: 没有激活的 Excel 或工作簿
    """
    import xlwings as xw
    应用 = xw.apps.active
    if 应用 is None:
        raise RuntimeError("没有检测到正在运行的 Excel，请先打开一个 Excel 文件。")
    工作簿 = 应用.books.active
    if 工作簿 is None:
        raise RuntimeError("当前 Excel 没有激活的工作簿。")
    工作表 = 工作簿.sheets[工作表名] if 工作表名 else 工作簿.sheets.active
    二维 = _整片读取从_A1(工作表)
    最终表头行 = _识别表头行号(二维, 默认值=表头行号)
    数据框 = _二维列表转df(二维, 最终表头行)
    debug_print(
        f"xlwings 读取激活簿：{工作簿.name} / {工作表.name} / {数据框.shape} / 表头行={最终表头行}")
    return 数据集(数据框=数据框,
               来源类型="active",
               来源路径=工作簿.fullname,
               工作簿名=工作簿.name,
               工作表名=工作表.name,
               表头行号=最终表头行)


# ---- 4.1) 列出文件 sheet 列表 ----
def 列出文件表名(文件路径: str) -> List[str]:
    """返回 Excel 文件全部 sheet 名；失败返回空列表。"""
    路径 = Path(文件路径)
    if not 路径.exists():
        return []
    if 路径.suffix.lower() == ".xls":
        import xlwings as xw
        应用 = None
        工作簿 = None
        try:
            应用 = xw.App(visible=False, add_book=False)
            工作簿 = 应用.books.open(str(路径), read_only=True, update_links=False)
            return [s.name for s in 工作簿.sheets]
        finally:
            try:
                if 工作簿 is not None:
                    工作簿.close()
            except Exception:
                pass
            try:
                if 应用 is not None:
                    应用.quit()
            except Exception:
                pass
    try:
        return pd.ExcelFile(路径).sheet_names
    except Exception as 异常:
        记录错误(f"列出工作表失败：{异常}")
        return []


# ---- 4.2) 列出激活簿 sheet 列表 ----
def 列出激活簿表名() -> List[str]:
    """返回激活 Excel 中激活簿的全部 sheet 名；无 Excel 时返回 []。"""
    import xlwings as xw
    应用 = xw.apps.active
    if 应用 is None or 应用.books.active is None:
        return []
    return [s.name for s in 应用.books.active.sheets]


# ---- 4.3) 获取当前激活簿标识（轮询检测切换/内容变化）----
def 获取激活簿标识() -> Optional[str]:
    """
    返回当前激活 Excel 激活簿的"指纹"，用于主窗口轮询检测变化。
    指纹包含：全路径 | 激活 Sheet 名 | used_range 行列数 | 表头行与末行内容 hash。
    任何下面变化都会导致指纹不同，触发自动重读：
        - 工作簿 / Sheet 切换
        - 行、列增减（used_range 变化）
        - 表头修改、最后一行以下区域内容修改（按采样检测）
    无 Excel / 无激活簿返回 None。
    """
    try:
        import xlwings as xw
        import hashlib

        应用 = xw.apps.active
        if 应用 is None:
            return None
        工作簿 = 应用.books.active
        if 工作簿 is None:
            return None
        工作表 = 工作簿.sheets.active
        表名 = 工作表.name if 工作表 is not None else ""

        最后行, 最后列 = 0, 0
        if 工作表 is not None:
            try:
                ur = 工作表.used_range
                最后行 = int(ur.last_cell.row)
                最后列 = int(ur.last_cell.column)
            except Exception:
                pass

        哈希器 = hashlib.md5()
        哈希器.update(f"{最后行}x{最后列}".encode("utf-8"))

        # 采样几个关键区域参与 hash，避免全表读取带来开销：
        #   - 表头行（首行）
        #   - 首中间行、末行
        # 对于 SAP 导出表，这些采样点足以反映绝大多数修改。
        if 工作表 is not None and 最后行 >= 1 and 最后列 >= 1:
            采样行号集 = {1, 最后行}
            if 最后行 >= 3:
                采样行号集.add(最后行 // 2)
            if 最后行 >= 2:
                采样行号集.add(最后行 - 1)
            for 行号 in sorted(采样行号集):
                try:
                    行值 = 工作表.range((行号, 1), (行号, 最后列)).value
                    哈希器.update(repr(行值).encode("utf-8", errors="replace"))
                except Exception:
                    哈希器.update(f"采样失败:{行号}".encode("utf-8"))

        指纹 = 哈希器.hexdigest()
        return f"{工作簿.fullname}|{表名}|{最后行}x{最后列}|{指纹}"
    except Exception:
        return None


# ---- 3.3) 二维列表 → DataFrame ----
def _二维列表转df(二维, 表头行号: int) -> pd.DataFrame:
    """
    把 xlwings 返回的二维列表按 表头行号 切分为 DataFrame。
    单单元格/单列等"非二维"输入会被自动归一化。
    """
    if 二维 is None:
        return pd.DataFrame()
    if not isinstance(二维, list):
        二维 = [[二维]]
    if 二维 and not isinstance(二维[0], (list, tuple)):
        二维 = [[行] for 行 in 二维]
    skip = max(0, 表头行号 - 1)
    if len(二维) <= skip:
        return pd.DataFrame()
    表头 = [str(c) if c is not None else f"列{i+1}"
          for i, c in enumerate(二维[skip])]
    数据 = 二维[skip + 1:]
    if not 数据:
        return pd.DataFrame(columns=表头)
    最大列 = max(len(表头), max(len(行) for 行 in 数据))
    while len(表头) < 最大列:
        表头.append(f"列{len(表头)+1}")
    数据 = [list(行) + [None] * (最大列 - len(行)) for 行 in 数据]
    return pd.DataFrame(数据, columns=表头)
