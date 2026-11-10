# -*- coding: utf-8 -*-
# ============================================================
# main.py —— SAP 表格数据处理工具 · 网页版后端（FastAPI）
#
# 设计原则：算法层零转写。
#   01_数据读写 / 02_数据处理 / 04_基础设施 三个原版模块通过原版
#   _bootstrap.加载() 按物理路径直接加载（与桌面版同一份代码），
#   本文件只做：HTTP 适配、数据集缓存、日志桥接、UI 消息编排。
#
# 与原版桌面程序的行为差异（平台所致，均已在下述注释标注）：
#   - 「读取激活簿 / 写回激活簿」依赖本机 Excel（xlwings COM），
#     网页版不支持 → 数据来源改为「上传文件」，输出改为「下载 .xlsx」。
#   - 原版 023_replace 用 dtype == object 枚举文本列，新版 pandas
#     字符串列为 StringDtype 会漏列 → 适配层用 is_string_dtype 兼容。
#
# 启动：python main.py （监听 $PORT，默认 8000，绑定 0.0.0.0）
# ============================================================

from __future__ import annotations

import os

# ============================================================
# 2026-11-04：钉钉 CLI（dws）凭据改用「文件后端」，不再写注册表。
# 现场：dws 登录授权成功，但保存 token 时 `registry create/open failed: Access is denied`
# （HKCU\Software\DwsCli\keychain\dws-cli 的 ACL 明明正常、PowerShell 也能写，dws 自己写就被拒），
# 结果就是「授权成功 → 存不下来 → 一上线就掉」，反复发作。
# dws 自带 DWS_DISABLE_KEYCHAIN=1 开关（README 有），改用它走 file-DEK；
# 站点所有 dws 子进程都继承本环境变量，保证与手工登录用的是同一个后端。
os.environ.setdefault("DWS_DISABLE_KEYCHAIN", "1")

import re as _re
import sys
import json
import base64
import hashlib
import hmac
import secrets
import uuid
import shutil
import subprocess
import threading
import time
import urllib.parse
from collections import deque
from io import BytesIO
from datetime import datetime
import datetime as dt
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.middleware.gzip import GZipMiddleware
from pydantic import BaseModel

# ---- 0) 定位原版源码目录（./_src，已随 app 一起打包，保证云端部署自包含）----
APP_DIR = Path(__file__).resolve().parent
SRC_DIR = APP_DIR / "_src"
if not SRC_DIR.exists():
    raise RuntimeError(f"找不到原版源码目录：{SRC_DIR}（应与 main.py 同级的 _src 目录）")
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from _bootstrap import 加载                                          # noqa: E402

_io = 加载("01_数据读写", "010_excel_io_main")          # 原版：读取/写出/列解析
_proc = 加载("02_数据处理", "020_processors_main")      # 原版：五大处理器
_log = 加载("04_基础设施", "041_log")                   # 原版：日志 + UI日志桥

UI信息 = _log.UI信息
UI错误 = _log.UI错误
UI日志桥 = _log.UI日志桥

import ai_regex                                                    # noqa: E402
import zpp_cpgl_parse                                              # noqa: E402  ZPP_CPGL 报表解析（纯函数，可单测）

app = FastAPI(title="SAP 表格数据处理工具（网页版）", version="1.0.0")
app.add_middleware(GZipMiddleware, minimum_size=800)   # JSON 响应压致 800B 以上才压缩，省 CPU

# 启动自检（2026-10-27）：pywin32 是 VBA/Excel 全链路的前提，缺了会整批接口 500 且报错被
# uvicorn 吞成纯文本，页面上只剩一句「VBA 运行失败」。启动时先大声提醒用的是哪个解释器。
try:
    import pythoncom as _自检_pythoncom  # noqa: F401
    import win32com.client as _自检_win32  # noqa: F401
except Exception as _自检异常:
    print(f"[启动自检] ⚠ 当前解释器缺少 pywin32（{type(_自检异常).__name__}: {_自检异常}）")
    print(f"[启动自检] ⚠ 解释器={sys.executable}")
    print("[启动自检] ⚠ VBA/Excel 相关接口都会失败，请改用 C:\\Python311\\python.exe 启动 main.py")

UPLOAD_DIR = APP_DIR / "uploads"
TMP_DIR = APP_DIR / "tmp"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
TMP_DIR.mkdir(parents=True, exist_ok=True)

支持的后缀 = (".xlsx", ".xlsm")

# ============ 日志桥接（复刻原版 GUI 日志面板）============

_日志锁 = threading.Lock()
_日志列表: List[dict] = []


def _收集日志(消息: str, 级别: str = "INFO") -> None:
    with _日志锁:
        _日志列表.append({
            "t": datetime.now().strftime("%H:%M:%S"),
            "msg": 消息,
            "level": 级别,
        })
        if len(_日志列表) > 500:
            del _日志列表[:-500]


UI日志桥.注册(_收集日志)          # 与原版 _注册日志桥 同一机制


# ============ 数据集缓存（内存，上限 60 份）============

_缓存锁 = threading.Lock()
_数据集缓存: "dict[str, dict]" = {}
_数据集顺序: List[str] = []
_缓存上限 = 60
_预览行数 = 100                    # 网页预览行数（原版表格模型上限 5000，网页取 100 兼顾流畅）


def _存入(ds, 文件路径: Optional[str] = None, 原始文件名: Optional[str] = None) -> str:
    ds_id = uuid.uuid4().hex[:12]
    with _缓存锁:
        _数据集缓存[ds_id] = {
            "ds": ds, "path": 文件路径, "name": 原始文件名 or ds.工作簿名,
        }
        _数据集顺序.append(ds_id)
        while len(_数据集顺序) > _缓存上限:
            旧 = _数据集顺序.pop(0)
            _数据集缓存.pop(旧, None)
    return ds_id


def _取(ds_id: str) -> dict:
    with _缓存锁:
        项 = _数据集缓存.get(ds_id)
    if 项 is None:
        raise HTTPException(status_code=404, detail="数据已过期，请重新上传文件")
    return 项


# ============ 展示辅助（复刻原版 036_table_model._格式化值）============

def _格式化值(值) -> str:
    if 值 is None:
        return ""
    try:
        if pd.isna(值):
            return ""
    except Exception:
        pass
    if isinstance(值, float) and 值.is_integer():
        return str(int(值))
    return str(值)


def _摘要(ds_id: str, 项: dict) -> dict:
    ds = 项["ds"]
    df = ds.数据框
    预览 = [[_格式化值(v) for v in row]
            for row in df.head(_预览行数).values.tolist()]
    return {
        "id": ds_id,
        "name": ds.工作簿名,
        "sheet": ds.工作表名,
        "header_row": ds.表头行号,
        "rows": int(len(df)),
        "cols": int(df.shape[1]),
        "columns": ds.列名列表,
        "preview": 预览,
        "preview_rows": min(len(df), _预览行数),
        "summary": ds.摘要(),
    }


def _执行结果(ds_id: str, 项: dict, 原列数: int, 消息: str) -> dict:
    结果 = _摘要(ds_id, 项)
    结果["new_columns"] = 项["ds"].列名列表[原列数:]
    结果["message"] = 消息
    return 结果


# ============ 文本列解析（适配新版 pandas StringDtype）============

def _是文本列(序列) -> bool:
    try:
        from pandas.api.types import is_object_dtype, is_string_dtype
        return bool(is_object_dtype(序列) or is_string_dtype(序列))
    except Exception:
        return 序列.dtype == object


def _解析作用列(数据框: pd.DataFrame, 列列表) -> List[str]:
    """原版 _应用到列列表 的适配版：空列表 = 全部文本列（兼容 StringDtype）。"""
    if not 列列表:
        return [c for c in 数据框.columns if _是文本列(数据框[c])]
    结果: List[str] = []
    for 引用 in 列列表:
        try:
            结果.append(_io.解析列引用(引用, 数据框))
        except Exception as 异常:
            UI错误(f"替换跳过列 {引用}：{异常}")
    return 结果


# ============ 基础接口 ============

@app.get("/api/health")
def health():
    """健康检查。2026-10-25 起带上「解释器 + playwright 可用性」，缺依赖一眼可见（不再靠点按钮才发现）。
    2026-10-27 再加 pywin32：服务若被别的解释器（无 pywin32）拉起，所有 VBA/Excel 接口会全挂，
    页面上只显示「VBA 运行失败」，这里直接暴露出来。"""
    信息 = {"status": "ok", "app": "03#箱工作站", "version": "1.0.0",
            "interpreter": sys.executable}
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        信息["pywin32"] = {"ok": True, "interpreter": sys.executable}
    except Exception as _e:
        信息["pywin32"] = {"ok": False, "interpreter": sys.executable,
                          "error": f"{type(_e).__name__}: {_e}",
                          "hint": "请用 C:\\Python311\\python.exe 启动 main.py（该解释器才装了 pywin32）"}
    try:
        import modify_todo_auto as _ma
        if hasattr(_ma, "playwright_可用性"):
            信息["playwright"] = _ma.playwright_可用性()
        else:
            信息["playwright"] = {"ok": bool(getattr(_ma, "_PLAYWRIGHT_OK", False)),
                                  "error": getattr(_ma, "_PLAYWRIGHT_ERR", "")}
    except Exception as _e:
        信息["playwright"] = {"ok": False, "error": f"模块未挂载：{type(_e).__name__}: {_e}"}
    return 信息


@app.post("/api/mail/check-bounce")
def mail_check_bounce():
    """手动触发退信感知 + 自动重发（2026-10-28 用户要求「发送失败自动重发」的异步层，纯 DWS 方案）。

    下游 MX 异步弹回的退信（如 502）系统原本无感知；本接口对 sent_mail_queue.json 中窗口内、
    尚未确认投递的已发送邮件调用 dws mail message verify 查 sendStatus，failed 即自动补发。
    完全基于 DWS（复用已有登录态），无需任何额外环境变量。
    """
    from pathlib import Path as _P
    通知目录 = (_P(__file__).resolve().parent.parent / "web_auto" / "公共_钉钉通知").resolve()
    try:
        if str(通知目录) not in sys.path:
            sys.path.insert(0, str(通知目录))
        import mail_bounce_watch
        return mail_bounce_watch.检查退信并重发()
    except Exception as exc:
        return {"enabled": False, "ok": False, "error": f"{type(exc).__name__}: {exc}"}


# ============ 字典 / 工厂日历 / 预投汇总（原「03#箱工作站」接口，原样保留）============

WORKSHOPS = ["一号车间", "二号车间", "三号车间", "装配车间"]
PLAN_TYPES = ["日计划", "周计划", "月计划", "临时计划"]
PRIORITIES = ["高", "中", "低"]
ORDER_STATUSES = ["待开始", "进行中", "已完成", "已取消"]
WEEKDAY_NAMES = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


# ============ 工厂日历 · 持久化存储（对齐 APS 生产计划平台行为） ============
# 数据文件 calendar_store.json：{"overrides": {"2026-10-01": {"type":"N","remark":"非工作日","by":"演示用户"}}}
# 无覆盖的日期按周一~五=工作日(W)、周六日=非工作日(N) 兜底。
CALENDAR_STORE_PATH = APP_DIR / "calendar_store.json"


def _calendar_store_load() -> dict:
    try:
        with CALENDAR_STORE_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("overrides"), dict):
            return data
    except Exception:
        pass
    return {"overrides": {}}


def _calendar_store_save(store: dict):
    CALENDAR_STORE_PATH.write_text(
        json.dumps(store, ensure_ascii=False, indent=2), encoding="utf-8")


def _calendar_default_type(cur: "date") -> str:
    return "W" if cur.weekday() < 5 else "N"


@app.get("/api/factory-calendar")
def factory_calendar(
    year: Optional[int] = Query(None, ge=2000, le=2100),
    month: Optional[int] = Query(None, ge=1, le=12),
):
    """按自然月返回工厂日历（含人工覆盖）：周一至周五默认工作日，周六日默认休息日。"""
    import calendar as pycal
    from datetime import date
    today = date.today()
    y = year or today.year
    m = month or today.month

    store = _calendar_store_load()
    overrides = store["overrides"]
    days_in_month = pycal.monthrange(y, m)[1]
    days: List[dict] = []
    for d in range(1, days_in_month + 1):
        cur = date(y, m, d)
        iso = cur.isoformat()
        ov = overrides.get(iso) or {}
        wtype = ov.get("type") or _calendar_default_type(cur)
        days.append({
            "date": iso, "day": d, "weekday": cur.weekday(),
            "weekday_name": WEEKDAY_NAMES[cur.weekday()],
            "is_workday": wtype == "W", "note": ov.get("remark"), "by": ov.get("by"),
        })
    return {
        "year": y, "month": m, "total_days": len(days),
        "workdays": sum(1 for x in days if x["is_workday"]),
        "holidays": sum(1 for x in days if not x["is_workday"]),
        "days": days,
    }


class CalendarUpdate(BaseModel):
    dates: List[str]
    type: str  # "W" | "N"
    remark: str = ""   # 备注可选
    by: Optional[str] = None


@app.post("/api/factory-calendar/update")
def factory_calendar_update(payload: CalendarUpdate):
    """批量把若干日期设为工作日/非工作日（备注可选）。"""
    from datetime import date as _date
    if payload.type not in ("W", "N"):
        raise HTTPException(status_code=400, detail="type 仅支持 W（工作日）/ N（非工作日）")
    parsed = []
    for s in payload.dates:
        try:
            parsed.append(_date.fromisoformat(s))
        except Exception:
            raise HTTPException(status_code=400, detail=f"日期格式不正确：{s}")
    if not parsed:
        raise HTTPException(status_code=400, detail="未选择日期")
    store = _calendar_store_load()
    for cur in parsed:
        store["overrides"][cur.isoformat()] = {
            "type": payload.type, "remark": payload.remark.strip(),
            "by": (payload.by or "").strip() or None,
        }
    _calendar_store_save(store)
    UI信息(f"工厂日历更新：{payload.dates} → {'工作日' if payload.type == 'W' else '非工作日'}（{payload.remark.strip()}）")
    return {"ok": True, "updated": len(parsed)}


# ============ 看板布局偏好（通用 ns 命名空间；2026-10-17 用户要求：手工拖卡片长宽 + 文字样式可改，且保存） ============
# board_layout_store.json：{"apd22": {"cards": {"a22Daily15": {"w": "1.3fr", "h": 460}}, "texts": {"a22Kpi1.title": {"color": "#f00", "size": 14, "text": "当日下达箱数"}}}}
# cards：键=卡稳定键（卡内图表 id），w=整排 grid-template-columns，h=卡高 px；texts：键=稳定键.title/.sub。
BOARD_LAYOUT_PATH = APP_DIR / "board_layout_store.json"


def _board_layout_load() -> dict:
    try:
        with BOARD_LAYOUT_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _board_layout_save(store: dict):
    BOARD_LAYOUT_PATH.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")


class BoardLayoutPayload(BaseModel):
    ns: str
    cards: Optional[dict] = None
    texts: Optional[dict] = None
    charts: Optional[dict] = None


@app.get("/api/board/layout")
def board_layout_get(ns: str):
    data = _board_layout_load()
    nsd = data.get(ns) or {}
    return {"cards": nsd.get("cards") or {}, "texts": nsd.get("texts") or {}, "charts": nsd.get("charts") or {}}


@app.post("/api/board/layout")
def board_layout_post(payload: BoardLayoutPayload):
    data = _board_layout_load()
    nsd = data.get(payload.ns) or {}
    if payload.cards is not None:
        nsd["cards"] = payload.cards
    if payload.texts is not None:
        nsd["texts"] = payload.texts
    if payload.charts is not None:
        nsd["charts"] = payload.charts
    data[payload.ns] = nsd
    _board_layout_save(data)
    UI信息(f"看板布局偏好已保存（{payload.ns}）")
    return {"ok": True}


@app.get("/api/meta/options")
def meta_options():
    """筛选项字典，前端下拉框从这里取。"""
    return {
        "车间": ["全部车间"] + WORKSHOPS,
        "计划类型": ["全部类型"] + PLAN_TYPES,
        "优先级": ["全部"] + PRIORITIES,
        "订单状态": ["全部"] + ORDER_STATUSES,
    }


class ForecastRows(BaseModel):
    rows: List[dict]


# ============ 预投基础数据表 · 工作表导入/导出/种子 ============
# 数据真身在云数据库 forecast_sheet_meta / forecast_sheet_rows（Python 无法直连数据面，
# 增删改查由前端 SDK 完成；后端只做「解析上传文件、生成导出文件、提供初始种子」）。

FORECAST_SEED_PATH = APP_DIR / "forecast_seed" / "seed.json"
FORECAST_STORE_PATH = APP_DIR / "forecast_store" / "local_store.json"   # 本地预览模式的行存储


def _store_load() -> dict:
    if FORECAST_STORE_PATH.exists():
        try:
            return json.loads(FORECAST_STORE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"next_id": 1, "sheets": {}}


def _store_save(store: dict) -> None:
    FORECAST_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    FORECAST_STORE_PATH.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")


def _store_seed_if_empty() -> dict:
    """首次使用：把种子（Excel 全量）灌入本地存储，行 id 用自增计数。"""
    store = _store_load()
    if store.get("sheets"):
        return store
    if FORECAST_SEED_PATH.exists():
        seed = json.loads(FORECAST_SEED_PATH.read_text(encoding="utf-8"))
        nid = 1
        for name, info in seed.get("sheets", {}).items():
            rows = []
            for r in info.get("rows", []):
                rows.append({"id": nid, "row_no": r["row_no"], "data": r["data"]})
                nid += 1
            store["sheets"][name] = rows
        store["next_id"] = nid
        _store_save(store)
    return store


class StoreUpsert(BaseModel):
    sheet: str
    id: int = 0              # 0 = 新增；>0 = 按 id 更新
    row_no: int
    data: dict


class StoreReplace(BaseModel):
    sheet: str
    rows: List[dict]          # [{row_no, data}]


@app.get("/api/forecast/store")
def store_get():
    """本地存储全量（sheet → [{id,row_no,data}]），供本地预览模式读写。"""
    return _store_seed_if_empty()


@app.post("/api/forecast/store/upsert")
def store_upsert(payload: StoreUpsert):
    """id=0 → 追加新行；id>0 → 按 id 更新 data。"""
    store = _store_seed_if_empty()
    rows = store["sheets"].setdefault(payload.sheet, [])
    if payload.id:
        for r in rows:
            if r["id"] == payload.id:
                r["data"] = payload.data
                _store_save(store)
                UI信息(f"本地存储修改 {payload.sheet} 第 {r['row_no']} 行")
                return {"ok": True, "action": "update"}
        raise HTTPException(status_code=404, detail=f"未找到行 id={payload.id}")
    rid = store.get("next_id", 1)
    store["next_id"] = rid + 1
    rows.append({"id": rid, "row_no": payload.row_no, "data": payload.data})
    _store_save(store)
    UI信息(f"本地存储新增 1 行（{payload.sheet}）")
    return {"ok": True, "action": "insert"}


@app.post("/api/forecast/store/delete")
def store_delete(payload: dict):
    """删除行：id=单行；ids=批量多行（勾选删除用）。"""
    sheet = payload.get("sheet")
    ids = payload.get("ids")
    if not isinstance(ids, list):
        row_id = payload.get("id")
        ids = [row_id] if row_id is not None else []
    id_set = set(ids)
    store = _store_seed_if_empty()
    rows = store["sheets"].get(sheet, [])
    store["sheets"][sheet] = [r for r in rows if r["id"] not in id_set]
    _store_save(store)
    return {"ok": True, "deleted": len(rows) - len(store["sheets"][sheet])}


@app.post("/api/forecast/store/replace")
def store_replace(payload: StoreReplace):
    """整表替换（导入-覆盖用）：清空该 sheet 后写入。

    2026-10-25 追加：rows 里带了合法 `id` 就沿用该 id（撤销/重做走这里回写，
    沿用 id 才能让前端选中状态、外部按 id 的引用保持一致）；没带 id（导入覆盖）
    或 id 冲突时才重新自增。
    """
    store = _store_seed_if_empty()
    rows = []
    已用 = set()
    for item in payload.rows:
        rid = item.get("id")
        if not isinstance(rid, int) or rid <= 0 or rid in 已用:
            rid = store.get("next_id", 1)
            store["next_id"] = rid + 1
        else:
            已用.add(rid)
            if rid >= store.get("next_id", 1):
                store["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": item.get("row_no"), "data": item.get("data") or {}})
    store["sheets"][payload.sheet] = rows
    _store_save(store)
    UI信息(f"本地存储覆盖导入 {payload.sheet}（{len(rows)} 行）")
    return {"ok": True, "count": len(rows)}


def _fc_单元格值(v) -> str:
    """归一化 Excel 单元格为字符串（与种子提取脚本同一规则，保证模板可比对）。"""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int,)):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    s = str(v)
    return "" if s.lower() == "nan" else s


def _fc_列字母(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _fc_读取网格(ws):
    """openpyxl 工作表 → 裁剪尾部空行空列后的字符串网格（原样返回，不预设表头）。"""
    grid = [[_fc_单元格值(c.value) for c in row] for row in ws.iter_rows()]
    used_cols = 0
    for row in grid:
        for j, v in enumerate(row):
            if v != "" and j + 1 > used_cols:
                used_cols = j + 1
    grid = [row[:used_cols] for row in grid]
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()
    return grid


@app.get("/api/forecast/seed")
def forecast_seed():
    """返回初始种子（yutouyouhuaa-ribbon.xlsm 全部 sheet 的内容，与原表一致）。"""
    if not FORECAST_SEED_PATH.exists():
        raise HTTPException(status_code=404, detail="种子文件不存在")
    return FileResponse(str(FORECAST_SEED_PATH),
                        media_type="application/json")


@app.post("/api/forecast/parse")
async def forecast_parse(file: UploadFile = File(...), sheet: Optional[str] = Form(None),
                         meta: Optional[str] = Form(None)):
    """解析上传的工作簿（.xlsx/.xlsm），返回原始字符串网格（不预设表头）。
    表头/数据行的划分与模板一致性判断由前端按 forecast_sheet_meta 完成。"""
    import openpyxl
    文件名 = file.filename or "上传.xlsx"
    后缀 = Path(文件名).suffix.lower()
    if 后缀 not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=400,
                            detail=f"仅支持 .xlsx / .xlsm 文件（{后缀 or '无扩展名'} 请先另存为 .xlsx）")
    内容 = await file.read()
    保存路径 = TMP_DIR / f"fcparse_{uuid.uuid4().hex[:8]}{后缀}"
    with 保存路径.open("wb") as f:
        f.write(内容)
    try:
        wb = openpyxl.load_workbook(保存路径, data_only=True)
    except Exception as 异常:
        raise HTTPException(status_code=400, detail=f"文件无法读取：{异常}")
    names = wb.sheetnames
    target = sheet if (sheet and sheet in names) else names[0]
    if meta:
        # 只要「每个工作表的表头 + 数据行数」：导入时给用户选工作表用（不传整张网格，省流量）
        清单 = []
        for 名 in names:
            网格 = _fc_读取网格(wb[名])
            清单.append({"name": 名,
                         "header": (网格[0] if 网格 else []),
                         "row_count": max(0, len(网格) - 1)})
        UI信息(f"预投导入解析（工作表清单）：{文件名} → {len(names)} 个工作表")
        return {"filename": 文件名, "sheets": names, "sheet": target, "meta": 清单}
    grid = _fc_读取网格(wb[target])
    UI信息(f"预投导入解析：{文件名} / {target} → {len(grid)} 行 × {len(grid[0]) if grid else 0} 列")
    return {"filename": 文件名, "sheets": names, "sheet": target, "grid": grid}


@app.post("/api/zppmd04n/parse")
async def zppmd04n_parse(file: UploadFile = File(...)):
    """解析 SAP Zpp_Md04n 导出的结余文件（.xls / .xlsx / .xlsm），返回字符串网格。

    供自动化任务 zpp_md04n_balance 使用：SAP 老导出为 .xls（BIFF），
    openpyxl 不支持，改用 pandas.read_excel（本地引擎 xlrd/openpyxl 自动选择）。
    网格经 _fc_读取网格 同款裁剪（去尾部空行空列、单元格归一化为字符串）。
    """
    文件名 = file.filename or "zpp_md04n.xls"
    后缀 = Path(文件名).suffix.lower()
    if 后缀 not in (".xls", ".xlsx", ".xlsm"):
        raise HTTPException(status_code=400,
                            detail=f"仅支持 .xls / .xlsx / .xlsm 文件（收到 {后缀 or '无扩展名'}）")
    内容 = await file.read()
    保存路径 = TMP_DIR / f"md04n_{uuid.uuid4().hex[:8]}{后缀}"
    with 保存路径.open("wb") as f:
        f.write(内容)
    try:
        df = pd.read_excel(保存路径, sheet_name=0, header=None, dtype=object)
    except Exception as 异常:
        raise HTTPException(status_code=400, detail=f"文件无法读取：{异常}")
    grid = [[_fc_单元格值(v) for v in row] for row in df.itertuples(index=False, name=None)]
    # 与 _fc_读取网格 相同的裁剪：去尾部空行、裁掉整列皆空的尾列
    used_cols = 0
    for row in grid:
        for j, v in enumerate(row):
            if v != "" and j + 1 > used_cols:
                used_cols = j + 1
    grid = [row[:used_cols] for row in grid]
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()
    UI信息(f"Zpp_Md04n 结余文件解析：{文件名} → {len(grid)} 行 × {used_cols} 列")
    return {"filename": 文件名, "rows": len(grid), "grid": grid}


class ForecastExport(BaseModel):
    filename: str = "导出.xlsx"
    sheet: Optional[str] = None
    headers: List[str]
    rows: List[List[str]]


@app.post("/api/forecast/export")
def forecast_export(payload: ForecastExport):
    """把当前工作表（表头+数据）写成 xlsx 供下载。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = (payload.sheet or "Sheet1")[:31]
    ws.append(payload.headers)
    for row in payload.rows:
        padded = list(row) + [""] * (len(payload.headers) - len(row))
        ws.append([_还原数值(v) for v in padded[:len(payload.headers)]])
    临时路径 = TMP_DIR / f"fcexport_{uuid.uuid4().hex[:8]}.xlsx"
    wb.save(临时路径)
    UI信息(f"预投数据已导出：{payload.filename}（{len(payload.rows)} 行）")
    return FileResponse(str(临时路径),
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition":
                                 f"attachment; filename*=UTF-8''{urllib.parse.quote(payload.filename)}"})


# ============ 通用「导出当前页」（右键菜单 → 导出）：存到用户指定目录，默认桌面 ============
# 用户 2026-10-25 明令：「右击增加一个导出功能，弹窗显示我是否导出当前页面并打开、保存路径，
# 默认保存在桌面，但可以选别的位置；文件名用当前功能模块标题 + 时间。」
# 与 /api/forecast/export 的区别：那个走浏览器下载（落到浏览器下载目录，用户改不了）；
# 这个由后端直接写盘到指定目录，写成功才回显路径，需要时再用 Excel 静默打开（绝不弹格式告警）。

def _桌面目录() -> Path:
    """本机「真」桌面：优先注册表 User Shell Folders（OneDrive 重定向后旧路径会失效），再回落常见路径。"""
    候选: List[Path] = []
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as k:
            v, _ = winreg.QueryValueEx(k, "Desktop")
            if v:
                候选.append(Path(os.path.expandvars(str(v))))
    except Exception:
        pass
    候选.append(Path(os.path.expanduser("~/Desktop")))
    候选.append(Path(os.path.expanduser("~/OneDrive/Desktop")))
    if os.environ.get("USERPROFILE"):
        候选.append(Path(os.environ["USERPROFILE"]) / "Desktop")
    for p in 候选:
        try:
            if p and p.is_dir():
                return p
        except Exception:
            continue
    return Path(os.path.expanduser("~"))


@app.get("/api/export/defaults")
def export_defaults():
    """导出弹窗的默认值：桌面路径 + 常用位置（下载 / 文档），供快捷按钮用。"""
    桌面 = _桌面目录()
    主页 = Path(os.path.expanduser("~"))
    常用 = {}
    for 名, p in (("下载", 主页 / "Downloads"), ("文档", 主页 / "Documents")):
        try:
            if p.is_dir():
                常用[名] = str(p)
        except Exception:
            pass
    return {"ok": True, "dir": str(桌面), "desktop": str(桌面), "common": 常用}


class PageExport(BaseModel):
    """当前页导出：前端把可见表格的表头/数据整块送过来，后端落盘（可选打开）。"""
    filename: str = "导出.xlsx"
    sheet: Optional[str] = None
    headers: List[str] = []
    rows: List[List[str]] = []
    dir: Optional[str] = ""
    open_after: bool = True


_非法文件名 = _re.compile(r'[\\/:*?"<>|\r\n\t]')


@app.post("/api/export/page")
def export_page(payload: PageExport):
    """把「当前页面」的表格存成 xlsx 到指定目录（默认桌面），并可按需用 Excel 打开。

    - 只写一个文件、写完才回显路径（不做「可能成功」的假回显）；
    - 同名文件已存在时自动加 (2)(3)…，绝不覆盖用户已有文件；
    - 打开走 `_用Excel打开不弹框`（DisplayAlerts=False + DispatchEx 独立实例），不弹格式告警。
    """
    import openpyxl
    if not payload.headers:
        raise HTTPException(status_code=400, detail="当前页面没有可导出的列")
    原始目录 = (payload.dir or "").strip()
    if 原始目录:
        目录 = Path(os.path.expandvars(os.path.expanduser(原始目录)))
    else:
        目录 = _桌面目录()
    try:
        目录.mkdir(parents=True, exist_ok=True)
        if not 目录.is_dir():
            raise NotADirectoryError(str(目录))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"保存路径不可用：{exc}")
    名 = _非法文件名.sub("_", (payload.filename or "导出.xlsx")).strip() or "导出.xlsx"
    if not 名.lower().endswith(".xlsx"):
        名 += ".xlsx"
    路径 = 目录 / 名
    序号 = 2
    while 路径.exists():
        路径 = 目录 / f"{Path(名).stem}({序号}){Path(名).suffix}"
        序号 += 1
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = _非法文件名.sub("_", (payload.sheet or "Sheet1"))[:31] or "Sheet1"
        ws.append([str(h) for h in payload.headers])
        for row in payload.rows:
            padded = list(row) + [""] * (len(payload.headers) - len(row))
            ws.append([_还原数值(v) for v in padded[:len(payload.headers)]])
        wb.save(str(路径))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"写文件失败：{exc}")
    结果 = {"ok": True, "path": str(路径), "dir": str(目录), "name": 路径.name,
            "rows": len(payload.rows), "cols": len(payload.headers), "opened": False, "openTip": ""}
    if payload.open_after:
        try:
            r = _用Excel打开不弹框(路径)
            结果["opened"] = bool(r.get("ok"))
            结果["openTip"] = str(r.get("说明") or r.get("error") or "")
        except Exception as exc:
            结果["openTip"] = f"打开失败：{exc}"
    UI信息(f"导出当前页：{路径}（{len(payload.rows)} 行 × {len(payload.headers)} 列）")
    return 结果


# ============ BOM 分层（常用功能 → Bom分层） ============
# 数据源用 xlwings 直接读 Excel：.xls / .xlsx 都吃（原 exe 走 pandas，.xls 依赖 xlrd 不可靠）。
# 分层算法在 bom_layering.py（由 zpp_cs12.exe 逆向还原的规则移植，存档见项目根 zpp_cs12_逆向恢复/）。
# 四个动作口径对齐 SAPGUI 页：直接打开 / 转化为xlsx / 导出为Excel（复用 /api/export/page）。


class BomPathIn(BaseModel):
    path: str = ""
    sheet: str = ""
    表头行: Optional[int] = None      # None＝自动识别（SAP 导出常有标题行）


class BomRunIn(BaseModel):
    path: str = ""
    sheet: str = ""
    headers: List[str] = []
    rows: List[list] = []
    表头行: int = 1
    中文描述列: Optional[int] = 10    # -1 / None＝自动找「描述」类列；越界也不报错


def _选Excel文件(标题: str = "选择 Excel 文件") -> str:
    """弹系统文件对话框取真实路径（ctypes/comdlg32）。

    为什么不用 <input type=file>：④「直接打开」和⑤「转化为xlsx」要操作**用户原本那个文件**，
    浏览器上传只能拿到一份副本，路径就丢了。也不用 tkinter（便携版 Python 常常没有 tcl/tk）。
    hwndOwner 传当前前台窗口，避免对话框被浏览器盖在下面。
    """
    if os.name != "nt":
        return ""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR), ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR), ("lCustData", wintypes.LPARAM),
            ("lpfnHook", wintypes.LPVOID), ("lpTemplateName", wintypes.LPCWSTR),
        ]

    缓冲 = ctypes.create_unicode_buffer(2048)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    try:
        ofn.hwndOwner = ctypes.windll.user32.GetForegroundWindow()
    except Exception:
        ofn.hwndOwner = 0
    ofn.lpstrFilter = "Excel 文件\0*.xlsx;*.xls;*.xlsm;*.csv\0所有文件\0*.*\0\0"
    ofn.lpstrFile = ctypes.cast(缓冲, wintypes.LPWSTR)
    ofn.nMaxFile = len(缓冲)
    ofn.lpstrTitle = 标题
    ofn.lpstrDefExt = "xlsx"
    # OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR | OFN_PATHMUSTEXIST
    ofn.Flags = 0x00080000 | 0x00001000 | 0x00000008 | 0x00000800
    结果 = ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn))
    return 缓冲.value if 结果 else ""


@app.post("/api/bom/pick")
def bom_pick():
    """① zpp_cs12数据源导入：弹本机文件对话框，返回选中文件的绝对路径。"""
    try:
        路径 = _选Excel文件("选择 BOM 数据源（.xlsx / .xls）")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"打不开文件对话框：{exc}")
    if not 路径:
        return {"ok": False, "cancelled": True}
    try:
        import bom_layering as BL
        信息 = BL.读取簿(路径)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    UI信息(f"Bom分层：导入 {信息['file']}（{信息['totalRows']} 行 × {len(信息['headers'])} 列）")
    return {"ok": True, **信息}


@app.post("/api/bom/load")
def bom_load(payload: BomPathIn):
    """按路径重新读簿（切工作表 / 改表头行 / 手动填路径 / 转化xlsx之后回读都走这里）。"""
    try:
        import bom_layering as BL
        return {"ok": True, **BL.读取簿(payload.path, payload.sheet or None, payload.表头行 or None)}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/bom/run")
def bom_run(payload: BomRunIn):
    """② Bom分层：把缩进式多层 BOM 拍平。带了 rows 就直接算（认页面上的现状），否则按路径现读。"""
    try:
        import bom_layering as BL
        if payload.rows:
            表头 = payload.headers or []
            行们 = payload.rows
            源行数 = len(行们)
        else:
            簿 = BL.读取簿(payload.path, payload.sheet or None)
            表头, 行们, 源行数 = 簿["headers"], 簿["rows"], len(簿["rows"])
        描述列 = payload.中文描述列
        if 描述列 is not None and 描述列 < 0:
            描述列 = None            # -1 ＝自动找「描述」类列
        结果 = BL.分层(表头, 行们, 中文描述列=描述列, 表头行=max(1, payload.表头行))
        统计 = dict(结果["stats"])
        统计["源行数"] = 源行数
        UI信息(f"Bom分层：{统计['合同数']} 个合同 / {统计['明细行数']} 条明细 / 最大 {统计['最大层级']} 层")
        return {"ok": True, "headers": 结果["headers"], "rows": 结果["rows"],
                "stats": 统计, "notes": 结果.get("notes", [])}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/bom/open")
def bom_open(payload: BomPathIn):
    """④ 直接打开：用本机 Excel 打开数据源（不弹格式告警、不抢用户正在编辑的实例）。"""
    路径 = Path(os.path.expandvars(os.path.expanduser(payload.path or "")))
    if not 路径.is_file():
        raise HTTPException(status_code=400, detail=f"文件不存在：{路径}")
    r = _用Excel打开不弹框(路径)
    UI信息(f"Bom分层：直接打开 {路径.name}")
    return {"ok": bool(r.get("ok")), "path": str(路径), "说明": str(r.get("说明") or r.get("error") or "")}


@app.post("/api/bom/convert")
def bom_convert(payload: BomPathIn):
    """⑤ 转化为xlsx：.xls → 同名 .xlsx（xlwings 另存，原 .xls 保留并顺手关掉旧窗口）。"""
    路径 = Path(os.path.expandvars(os.path.expanduser(payload.path or "")))
    if not 路径.is_file():
        raise HTTPException(status_code=400, detail=f"文件不存在：{路径}")
    r = _xls转xlsx(路径)
    if not r.get("ok"):
        raise HTTPException(status_code=400, detail=str(r.get("说明") or "转换失败"))
    UI信息(f"Bom分层：转化为 xlsx → {r.get('文件')}")
    return {"ok": True, "path": r.get("路径"), "file": r.get("文件"), "说明": r.get("说明")}


# ============ 安全钳基础数据 · 本地存储接口（页面结构参考预投工作表页） ============
# 数据真身本地 JSON（safety_store/local_store.json）；线上发布后可按预投双通道同款扩展云表。
# /api/safety/seed 返回安全钳总图号分类种子（图片提取，已剔除空行，第1/3列）。

SAFETY_SEED_PATH = APP_DIR / "safety_seed" / "safety_seed.json"
SAFETY_STORE_PATH = APP_DIR / "safety_store" / "local_store.json"


def _safety_store_load() -> dict:
    if SAFETY_STORE_PATH.exists():
        try:
            return json.loads(SAFETY_STORE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"next_id": 1, "sheets": {}}


def _safety_store_save(store: dict) -> None:
    SAFETY_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    SAFETY_STORE_PATH.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")


# 2026-11-06「供应商基础数据 > 停产供方」：供应商代号必须 6 位；写入时自动去掉全部前导 0
# （用户口径：「可能 3 个可能 N 个 0 开头」都删）。只在「停产供方」这一张 sheet 生效，
# 其它安全钳/限速器基础数据表、预投 store 一律不受影响。
停产供方_SHEET = "停产供方"
停产供方_CODE = "供应商代号"


def _停产供方_清洗(sheet: str, data: dict) -> dict:
    if sheet != 停产供方_SHEET or not isinstance(data, dict) or 停产供方_CODE not in data:
        return data
    v = str(data.get(停产供方_CODE) or "").strip()
    out = dict(data)
    out[停产供方_CODE] = v.lstrip("0") if v else ""
    return out


def _safety_store_seed_if_empty() -> dict:
    """首次使用：把安全钳/限速器分类种子灌入本地存储。"""
    store = _safety_store_load()
    if store.get("sheets"):
        # 框架配置分类 sheet 缺失时补灌（2026-10-16 拆分：两视图独立数据）
        if "安全钳框架配置分类" not in store["sheets"]:
            _safety_seed_frame(store)
        # 限速器参数 sheet 缺失时补灌（来源=桌面「限速器参数.xls」解析生成的 limiter_seed.json）
        if "限速器参数" not in store["sheets"]:
            _safety_seed_limiter(store)
        # 限速器物料号 sheet 缺失时补灌（2026-11-03 新增：ZLSTMA 清单，limiter_material_seed.json）
        if "限速器物料号" not in store["sheets"]:
            _safety_seed_limiter_material(store)
        # 停产供方 sheet 缺失时补灌（2026-11-06 新增：311.xls C/I 列，80 行）
        if 停产供方_SHEET not in store["sheets"]:
            _safety_seed_stopped_supplier(store)
        return store
    if SAFETY_SEED_PATH.exists():
        seed = json.loads(SAFETY_SEED_PATH.read_text(encoding="utf-8"))
        nid = 1
        name = seed.get("sheet", "安全钳总图号分类")
        rows = []
        for r in seed.get("rows", []):
            rows.append({"id": nid, "row_no": r["row_no"], "data": r["data"]})
            nid += 1
        store["sheets"][name] = rows
        store["next_id"] = nid
        _safety_seed_frame(store)
        _safety_seed_limiter(store)
        _safety_seed_limiter_material(store)
        _safety_seed_stopped_supplier(store)
        _safety_store_save(store)
    return store


def _safety_seed_frame(store: dict) -> None:
    """灌入「安全钳框架配置分类」种子（图号+分类 两列，来源=用户附件 工作簿12.xlsx 52 行）。"""
    seed_p = APP_DIR / "safety_seed" / "frame_seed.json"
    if not seed_p.exists():
        return
    try:
        seed = json.loads(seed_p.read_text(encoding="utf-8"))
    except Exception:
        return
    rows = []
    nid = store.get("next_id", 1)
    for r in seed.get("rows", []):
        rows.append({"id": nid, "row_no": r["row_no"], "data": r["data"]})
        nid += 1
    store["sheets"]["安全钳框架配置分类"] = rows
    store["next_id"] = nid
    _safety_store_save(store)
    UI信息(f"安全钳框架配置分类种子灌入 {len(rows)} 行")


def _safety_seed_limiter(store: dict) -> None:
    """灌入「限速器参数」种子（集团/图号变量/型号/额定速度/超速开关/上行动作速度/下行动作速度/
    钢丝绳直径/创建日期/创建者，来源=limiter_seed.json，612 行）。"""
    seed_p = APP_DIR / "safety_seed" / "limiter_seed.json"
    if not seed_p.exists() or store["sheets"].get("限速器参数"):
        return
    try:
        seed = json.loads(seed_p.read_text(encoding="utf-8"))
    except Exception:
        return
    rows = []
    nid = store.get("next_id", 1)
    for r in seed.get("rows", []):
        rows.append({"id": nid, "row_no": r["row_no"], "data": r["data"]})
        nid += 1
    store["sheets"]["限速器参数"] = rows
    store["next_id"] = nid
    _safety_store_save(store)
    UI信息(f"限速器参数种子灌入 {len(rows)} 行")


def _safety_seed_limiter_material(store: dict) -> None:
    """灌入「限速器物料号」种子（2026-11-03 新增：SAP ZLSTMA 清单导出，31 列，4266 行）。
    来源=limiter_material_seed.json（用户附件 2.xls：第4行=字段名、第5行起=数据，换行碎片行已剔除；
    原「物料/标记」重复列名去重为 物料(类型)/标记2）。"""
    seed_p = APP_DIR / "safety_seed" / "limiter_material_seed.json"
    if not seed_p.exists() or store["sheets"].get("限速器物料号"):
        return
    try:
        seed = json.loads(seed_p.read_text(encoding="utf-8"))
    except Exception:
        return
    rows = []
    nid = store.get("next_id", 1)
    for r in seed.get("rows", []):
        rows.append({"id": nid, "row_no": r["row_no"], "data": r["data"]})
        nid += 1
    store["sheets"]["限速器物料号"] = rows
    store["next_id"] = nid
    _safety_store_save(store)
    UI信息(f"限速器物料号种子灌入 {len(rows)} 行")

    # 用户 2026-11-03 口径：按「创建日期」倒序（最新的在最前）
    def _lm日期键(r):
        for k in ("创建日期", "输入日期"):
            v = str((r.get("data") or {}).get(k) or "").strip()
            try:
                return dt.datetime.strptime(v, "%Y.%m.%d")
            except Exception:
                continue
        return dt.datetime(1970, 1, 1)
    rows.sort(key=_lm日期键, reverse=True)
    for i, r in enumerate(rows, 1):
        r["row_no"] = i
    _safety_store_save(store)


def _safety_seed_stopped_supplier(store: dict) -> None:
    """灌入「停产供方」种子（2026-11-06 新增，两列 供应商代号/供应商名称，80 行）。
    来源=用户附件 311.xls（SAP LFA1 ALV 导出，实为 GBK+Tab 文本）：C 列「供应商」、
    I 列「名称」；首行字段名、空行已剔除，代号写库前统一去掉全部前导 0。"""
    seed_p = APP_DIR / "safety_seed" / "stopped_supplier_seed.json"
    if not seed_p.exists() or store["sheets"].get(停产供方_SHEET):
        return
    try:
        seed = json.loads(seed_p.read_text(encoding="utf-8"))
    except Exception:
        return
    rows = []
    nid = store.get("next_id", 1)
    for r in seed.get("rows", []):
        rows.append({"id": nid, "row_no": r["row_no"],
                     "data": _停产供方_清洗(停产供方_SHEET, r.get("data") or {})})
        nid += 1
    store["sheets"][停产供方_SHEET] = rows
    store["next_id"] = nid
    _safety_store_save(store)
    UI信息(f"停产供方种子灌入 {len(rows)} 行")


class SafetyUpsert(BaseModel):
    sheet: str
    id: int = 0              # 0 = 新增；>0 = 按 id 更新
    row_no: int
    data: dict


class SafetyReplace(BaseModel):
    sheet: str
    rows: List[dict]          # [{row_no, data}]


def _limiter_material_sort_desc(store: dict) -> dict:
    """「限速器物料号」按 创建日期 倒序整理（2026-11-03 用户口径：最新的在最前）。

    每次读接口都核对一次顺序：已经倒序就原样返回（不重写盘）；发现乱序（历史数据/
    手工新增/导入插入的行）才落盘重排。row_no 全表重编，id 保持不变（撤销/引用不破）。"""
    rows = (store.get("sheets") or {}).get("限速器物料号") or []
    if len(rows) < 2:
        return store
    def _键(r):
        for k in ("创建日期", "输入日期"):
            v = str((r.get("data") or {}).get(k) or "").strip()
            try:
                return dt.datetime.strptime(v, "%Y.%m.%d")
            except Exception:
                continue
        return dt.datetime(1970, 1, 1)
    有序 = sorted(rows, key=_键, reverse=True)
    if [r.get("id") for r in 有序] == [r.get("id") for r in rows]:
        return store                          # 顺序已正确，不折腾
    for i, r in enumerate(有序, 1):
        r["row_no"] = i
    store["sheets"]["限速器物料号"] = 有序
    _safety_store_save(store)
    UI信息("限速器物料号已按 创建日期 倒序重排")
    return store


@app.get("/api/safety/store")
def safety_store_get():
    """安全钳本地存储全量（sheet → [{id,row_no,data}]）。「限速器物料号」按创建日期倒序。"""
    return _limiter_material_sort_desc(_safety_store_seed_if_empty())


@app.post("/api/safety/store/upsert")
def safety_store_upsert(payload: SafetyUpsert):
    """id=0 → 追加新行；id>0 → 按 id 更新 data。"""
    store = _safety_store_seed_if_empty()
    rows = store["sheets"].setdefault(payload.sheet, [])
    data = _停产供方_清洗(payload.sheet, payload.data)
    if payload.id:
        for r in rows:
            if r["id"] == payload.id:
                r["data"] = data
                _safety_store_save(store)
                UI信息(f"安全钳本地存储修改 {payload.sheet} 第 {r['row_no']} 行")
                return {"ok": True, "action": "update"}
        raise HTTPException(status_code=404, detail=f"未找到行 id={payload.id}")
    rid = store.get("next_id", 1)
    store["next_id"] = rid + 1
    rows.append({"id": rid, "row_no": payload.row_no, "data": data})
    if payload.sheet == "限速器物料号":
        _limiter_material_sort_desc(store)
    else:
        _safety_store_save(store)
    UI信息(f"安全钳本地存储新增 1 行（{payload.sheet}）")
    return {"ok": True, "action": "insert"}


@app.post("/api/safety/store/delete")
def safety_store_delete(payload: dict):
    """删除行：id=单行；ids=批量多行（勾选删除用）。"""
    sheet = payload.get("sheet")
    ids = payload.get("ids")
    if not isinstance(ids, list):
        row_id = payload.get("id")
        ids = [row_id] if row_id is not None else []
    id_set = set(ids)
    store = _safety_store_seed_if_empty()
    rows = store["sheets"].get(sheet, [])
    store["sheets"][sheet] = [r for r in rows if r["id"] not in id_set]
    _safety_store_save(store)
    return {"ok": True, "deleted": len(rows) - len(store["sheets"][sheet])}


@app.post("/api/safety/store/replace")
def safety_store_replace(payload: SafetyReplace):
    """整表替换（导入-覆盖用）：清空该 sheet 后写入（带合法 id 则沿用，撤销回写用）。"""
    store = _safety_store_seed_if_empty()
    rows = []
    已用 = set()
    for item in payload.rows:
        rid = item.get("id")
        if not isinstance(rid, int) or rid <= 0 or rid in 已用:
            rid = store.get("next_id", 1)
            store["next_id"] = rid + 1
        else:
            已用.add(rid)
            if rid >= store.get("next_id", 1):
                store["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": item.get("row_no"),
                     "data": _停产供方_清洗(payload.sheet, item.get("data") or {})})
    store["sheets"][payload.sheet] = rows
    if payload.sheet == "限速器物料号":
        _limiter_material_sort_desc(store)
    else:
        _safety_store_save(store)
    UI信息(f"安全钳本地存储覆盖导入 {payload.sheet}（{len(rows)} 行）")
    return {"ok": True, "count": len(rows)}

@app.get("/api/safety/seed")
def safety_seed():
    """安全钳总图号分类初始种子（图片提取：第1/3列，空行已删）。"""
    if not SAFETY_SEED_PATH.exists():
        raise HTTPException(status_code=404, detail="安全钳种子文件不存在")
    return FileResponse(str(SAFETY_SEED_PATH), media_type="application/json")


# ============ 安全钳物料号 · 大表服务端分页接口（2026-11-03 新增） ============
# 数据源：用户附件 6.xls（SAP ZLSTMA 安全钳物料清单，GBK 制表符文本）：
#   第4行=字段名（40 列，含 9 个空白列，用户要求一个不删）、第5行空、第6行起=数据，共 224,510 行。
# 体量远超 /api/safety/store「整表 JSON 一次拉取」的浏览器承载 → 落 SQLite（safety_material.db），
# 页面按页取数。列序完全保留原 Excel：c01..c40 ↔ meta.json 的 display 显示名（空白列→(空白列N)、
# 重复名→物料(2)/标记(2)），编辑/导出均按显示名映射回 c 列。
# 生成脚本：工作区 build_safety_material_db.py（xlwings 已交叉验证行列数与表头）。

SAFETY_MATERIAL_DB = APP_DIR / "safety_material_store" / "safety_material.db"
SAFETY_MATERIAL_META = APP_DIR / "safety_material_store" / "meta.json"
_safety_material_cols_cache: Optional[List[dict]] = None


def _safety_material_cols() -> List[dict]:
    global _safety_material_cols_cache
    if _safety_material_cols_cache is None:
        try:
            meta = json.loads(SAFETY_MATERIAL_META.read_text(encoding="utf-8"))
            _safety_material_cols_cache = meta.get("cols") or []
        except Exception:
            _safety_material_cols_cache = []
    return _safety_material_cols_cache


def _safety_material_conn():
    import sqlite3
    if not SAFETY_MATERIAL_DB.exists():
        raise HTTPException(status_code=404, detail="安全钳物料号数据库不存在（safety_material.db）")
    conn = sqlite3.connect(str(SAFETY_MATERIAL_DB))
    conn.row_factory = sqlite3.Row
    return conn


def _safety_material_where(q: str, keys: List[str]):
    """关键字查询：空格分词 AND，每个词命中任意一列即留下（与前端 kit.quick 同语义）。
    返回（裸条件串或空串, 参数）——WHERE 关键字由 _safety_material_join 统一拼，避免双 WHERE。"""
    tokens = [t for t in (q or "").split() if t]
    if not tokens:
        return "", []
    conds, params = [], []
    for t in tokens:
        like = f"%{t}%"
        conds.append("(" + " OR ".join([f'"{k}" LIKE ?' for k in keys]) + ")")
        params.extend([like] * len(keys))
    return " AND ".join(conds), params


def _safety_material_join(*cond_parts) -> str:
    """把若干「裸条件串」用 AND 拼成 WHERE 子句；全空则返回空串。"""
    parts = [c.strip() for c in cond_parts if c and c.strip()]
    return (" WHERE " + " AND ".join(parts)) if parts else ""


@app.get("/api/safety_material/meta")
def safety_material_meta():
    """列元数据（key/显示名，40 列原样含空白列）+ 总行数。"""
    cols = _safety_material_cols()
    if not cols:
        raise HTTPException(status_code=500, detail="安全钳物料号列元数据缺失（meta.json）")
    conn = _safety_material_conn()
    try:
        n = conn.execute('SELECT COUNT(*) FROM "material"').fetchone()[0]
    finally:
        conn.close()
    return {"cols": cols, "rows": n, "sheet": "安全钳物料号"}


@app.get("/api/safety_material/values")
def safety_material_values(col: str, limit: int = 500):
    """列筛选面板取值：某列去重值+计数（按计数降序，默认前 500 个）。col=显示名或 c 键。"""
    cols = _safety_material_cols()
    by_disp = {str(c["display"]).strip(): c["key"] for c in cols}
    k = by_disp.get(str(col or "").strip()) or (str(col).strip() if str(col).strip() in by_disp.values() else "")
    if not k:
        raise HTTPException(status_code=404, detail=f"未知列：{col}")
    limit = min(max(10, int(limit)), 2000)
    conn = _safety_material_conn()
    try:
        rs = conn.execute(
            f'SELECT "{k}" AS v, COUNT(*) AS n FROM "material" GROUP BY "{k}" ORDER BY n DESC, v LIMIT ?',
            [limit]).fetchall()
        return {"col": col, "values": [{"v": "" if r["v"] is None else str(r["v"]), "n": r["n"]} for r in rs],
                "truncated": len(rs) >= limit}
    finally:
        conn.close()


@app.get("/api/safety_material/page")
def safety_material_page(page: int = 1, size: int = 500, q: str = ""):
    """按页取数（服务端分页）：q=关键字全表匹配（空格分词 AND）。rows:[{row_no,id,data{显示名:值}}]。"""
    cols = _safety_material_cols()
    if not cols:
        raise HTTPException(status_code=500, detail="安全钳物料号列元数据缺失（meta.json）")
    keys = [c["key"] for c in cols]
    disps = [c["display"] for c in cols]
    page = max(1, int(page))
    size = min(max(50, int(size)), 10000)
    wq, params = _safety_material_where(q, keys)
    where = _safety_material_join(wq)
    sel = ", ".join([f'"{k}"' for k in keys])
    conn = _safety_material_conn()
    try:
        total = conn.execute(f'SELECT COUNT(*) FROM "material"{where}', params).fetchone()[0]
        pages = max(1, (total + size - 1) // size)
        page = min(page, pages)
        off = (page - 1) * size
        rs = conn.execute(
            f'SELECT "row_no", {sel} FROM "material"{where} ORDER BY "row_no" LIMIT ? OFFSET ?',
            params + [size, off]).fetchall()
        rows = [{"row_no": r["row_no"], "id": r["row_no"],
                 "data": {disps[i]: ("" if r[keys[i]] is None else str(r[keys[i]])) for i in range(len(keys))}}
                for r in rs]
        return {"rows": rows, "total": total, "page": page, "size": size, "pages": pages}
    finally:
        conn.close()


class SafetyMaterialPageReq(BaseModel):
    """前端大表取数请求（POST：筛选值列表可能很长，GET URL 放不下）。"""
    page: int = 1
    size: int = 500
    q: str = ""                      # 关键字（空格分词 AND，命中任意列）
    sort_col: str = ""               # 列显示名或 c 键；空=按行号
    sort_dir: str = "asc"            # asc / desc
    filter_key: str = ""             # 列筛选/文本筛选的列（二选一生效）
    filter_vals: List[str] = []      # 勾选值精确匹配（IN）
    filter_text: str = ""            # 文本模糊匹配（LIKE）
    cond_key: str = ""               # 查询抽屉第一条件
    cond_op: str = "包含"            # 包含/等于/不包含/不等于
    cond_val: str = ""


def _safety_material_filters_sql(req: SafetyMaterialPageReq, cols: List[dict]):
    """把前端列筛选/文本筛选/抽屉条件翻译成 SQL（与前端 kit 语义一致，全部 AND 组合）。
    返回（裸条件串或空串, 参数）——与关键字条件一起经 _safety_material_join 统一拼 WHERE。"""
    by_disp = {str(c["display"]).strip(): c["key"] for c in cols}
    by_key = {c["key"]: c["key"] for c in cols}

    def key_of(name: str) -> str:
        n = str(name or "").strip()
        return by_disp.get(n) or by_key.get(n) or ""

    conds, params = [], []
    if req.filter_key:
        k = key_of(req.filter_key)
        if k:
            vals = [str(v) for v in (req.filter_vals or [])]
            if vals:
                ph = ",".join(["?"] * len(vals))
                conds.append(f'"{k}" IN ({ph})')
                params.extend(vals)
            elif req.filter_text:
                conds.append(f'"{k}" LIKE ?')
                params.append(f"%{req.filter_text}%")
    if req.cond_key and str(req.cond_val).strip():
        k = key_of(req.cond_key)
        if k:
            v = str(req.cond_val).strip()
            op = req.cond_op or "包含"
            if op == "等于":
                conds.append(f'"{k}" = ?'); params.append(v)
            elif op == "不等于":
                conds.append(f'"{k}" != ?'); params.append(v)
            elif op == "不包含":
                conds.append(f'"{k}" NOT LIKE ?'); params.append(f"%{v}%")
            else:  # 包含
                conds.append(f'"{k}" LIKE ?'); params.append(f"%{v}%")
    sql = (" AND ".join(conds)) if conds else ""
    return sql, params


def _safety_material_order_sql(req: SafetyMaterialPageReq, cols: List[dict]):
    """排序下推：纯数字串走数值序、其余按文本序（与前端 isNumericCell 口径一致，避开日期串被截断的坑）。"""
    by_disp = {str(c["display"]).strip(): c["key"] for c in cols}
    n = str(req.sort_col or "").strip()
    if not n or n == "行号":
        # 2026-11-03 用户口径：安全钳物料号默认按「创建日期」倒序（最新在最前；解析不了的排最后）
        return ("ORDER BY CASE WHEN \"c18\" <> '' AND \"c18\" GLOB '[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]' "
                "THEN \"c18\" END DESC, \"row_no\" ASC")
    if n == "创建日期":
        # 点「创建日期」表头排序：asc=正序、desc=倒序（含 GLOB 校验，空/畸形日期排最后）
        d = "DESC" if (req.sort_dir or "asc").lower() != "asc" else "ASC"
        # 有效日期按方向排；空/畸形日期一律垫底（升序排有效值后面，倒序排最后）—— 两侧都用「无效=1」首键
        return (f"ORDER BY (CASE WHEN \"c18\" <> '' AND \"c18\" GLOB '[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]' THEN 0 ELSE 1 END) ASC, "
                f"(CASE WHEN \"c18\" <> '' AND \"c18\" GLOB '[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]' THEN \"c18\" END) {d}, \"row_no\" ASC")
    k = by_disp.get(n) or (n if n in [c["key"] for c in cols] else "")
    if not k:
        return 'ORDER BY "row_no"'
    d = "ASC" if (req.sort_dir or "asc").lower() != "desc" else "DESC"
    # GLOB '*[^0-9]*' = 含非数字字符；NOT GLOB = 整串纯数字 → 数值排序，其余文本排序（空值排最后）
    return (f'ORDER BY CASE WHEN "{k}" IS NOT NULL AND "{k}" <> \'\' AND "{k}" NOT GLOB \'*[^0-9]*\' '
            f'THEN CAST("{k}" AS REAL) END {d} NULLS LAST, "{k}" {d} NULLS LAST, "row_no" ASC')


@app.post("/api/safety_material/page")
def safety_material_page_post(req: SafetyMaterialPageReq):
    """按页取数（服务端分页，POST 版）：关键字/列筛选/文本筛选/抽屉条件/排序全部下推 SQLite。"""
    cols = _safety_material_cols()
    if not cols:
        raise HTTPException(status_code=500, detail="安全钳物料号列元数据缺失（meta.json）")
    keys = [c["key"] for c in cols]
    disps = [c["display"] for c in cols]
    page = max(1, int(req.page))
    size = min(max(50, int(req.size)), 10000)
    wq, pq = _safety_material_where(req.q, keys)
    wf, pf = _safety_material_filters_sql(req, cols)
    where = _safety_material_join(wq, wf)
    params = pq + pf
    sel = ", ".join([f'"{k}"' for k in keys])
    conn = _safety_material_conn()
    try:
        total = conn.execute(f'SELECT COUNT(*) FROM "material"{where}', params).fetchone()[0]
        pages = max(1, (total + size - 1) // size)
        page = min(page, pages)
        off = (page - 1) * size
        rs = conn.execute(
            f'SELECT "row_no", {sel} FROM "material"{where} {_safety_material_order_sql(req, cols)} LIMIT ? OFFSET ?',
            params + [size, off]).fetchall()
        rows = [{"row_no": r["row_no"], "id": r["row_no"],
                 "data": {disps[i]: ("" if r[keys[i]] is None else str(r[keys[i]])) for i in range(len(keys))}}
                for r in rs]
        return {"rows": rows, "total": total, "page": page, "size": size, "pages": pages}
    finally:
        conn.close()


class SafetyMaterialUpsert(BaseModel):
    row_no: int              # 0=追加到末尾；>0=按行号整行更新
    data: dict               # {显示名: 值}（与页面列名一致，服务端映射回 c 列）


@app.post("/api/safety_material/upsert")
def safety_material_upsert(payload: SafetyMaterialUpsert):
    """编辑/新增一行：data 键=列显示名；row_no=0 时追加到末尾（行号=MAX+1）。"""
    cols = _safety_material_cols()
    keys = [c["key"] for c in cols]
    by_disp = {str(c["display"]).strip(): c["key"] for c in cols}
    vals = [""] * len(cols)
    for disp, v in (payload.data or {}).items():
        k = by_disp.get(str(disp).strip())
        if k:
            vals[keys.index(k)] = "" if v is None else str(v)
    conn = _safety_material_conn()
    act = "update"
    try:
        cur = conn.cursor()
        if payload.row_no and payload.row_no > 0:
            ex = cur.execute('SELECT 1 FROM "material" WHERE "row_no"=?', (payload.row_no,)).fetchone()
            if not ex:
                raise HTTPException(status_code=404, detail=f"未找到行 row_no={payload.row_no}")
            sets = ", ".join([f'"{k}"=?' for k in keys])
            cur.execute(f'UPDATE "material" SET {sets} WHERE "row_no"=?', vals + [payload.row_no])
        else:
            nxt = cur.execute('SELECT MAX("row_no") FROM "material"').fetchone()[0] or 0
            payload.row_no = int(nxt) + 1
            ph = ",".join(["?"] * (len(keys) + 1))
            sel = ", ".join([f'"{k}"' for k in keys])
            cur.execute(f'INSERT INTO "material" ("row_no",{sel}) VALUES ({ph})',
                        [payload.row_no] + vals)
            act = "insert"
        conn.commit()
    finally:
        conn.close()
    UI信息(f"安全钳物料号{'修改' if act == 'update' else '新增'} 第 {payload.row_no} 行")
    return {"ok": True, "action": act, "row_no": payload.row_no}


class SafetyMaterialBulk(BaseModel):
    """大表批量写（2026-11-04）：拖动填充 / 区域粘贴 / Delete 清空 / 撤销回写。
    data 里只放要改的列 → **只更新这些列**（其它列原样不动），行不存在则跳过。"""
    updates: List[dict]      # [{row_no, data:{显示名:值}}]


@app.post("/api/safety_material/bulk")
def safety_material_bulk(payload: SafetyMaterialBulk):
    """一条事务提交多行改动：一次往返，替代「每格一次 upsert + 每次整页重取」。"""
    cols = _safety_material_cols()
    if not cols:
        raise HTTPException(status_code=500, detail="安全钳物料号列元数据缺失（meta.json）")
    by_disp = {str(c["display"]).strip(): c["key"] for c in cols}
    conn = _safety_material_conn()
    n = 0
    try:
        cur = conn.cursor()
        for u in (payload.updates or [])[:5000]:
            u = u or {}
            try:
                row_no = int(u.get("row_no") or 0)
            except Exception:
                continue
            data = u.get("data") or {}
            if row_no <= 0 or not isinstance(data, dict):
                continue
            pairs = [(by_disp[str(k).strip()], ("" if v is None else str(v)))
                     for k, v in data.items() if str(k).strip() in by_disp]
            if not pairs:
                continue
            sets = ", ".join([f'"{k}"=?' for k, _ in pairs])
            cur.execute(f'UPDATE "material" SET {sets} WHERE "row_no"=?',
                        [v for _, v in pairs] + [row_no])
            if cur.rowcount and cur.rowcount > 0:
                n += 1
        conn.commit()
    finally:
        conn.close()
    if n:
        UI信息(f"安全钳物料号批量修改 {n} 行")
    return {"ok": True, "count": n}


class SafetyMaterialDelete(BaseModel):
    row_nos: List[int]


@app.post("/api/safety_material/delete")
def safety_material_delete(payload: SafetyMaterialDelete):
    """按行号批量删除（行内删除/勾选删除共用）。"""
    ids = [int(x) for x in (payload.row_nos or []) if int(x) > 0]
    if not ids:
        return {"ok": True, "deleted": 0}
    conn = _safety_material_conn()
    try:
        cur = conn.cursor()
        ph = ",".join(["?"] * len(ids))
        before = cur.execute('SELECT COUNT(*) FROM "material"').fetchone()[0]
        cur.execute(f'DELETE FROM "material" WHERE "row_no" IN ({ph})', ids)
        conn.commit()
        deleted = before - cur.execute('SELECT COUNT(*) FROM "material"').fetchone()[0]
    finally:
        conn.close()
    UI信息(f"安全钳物料号删除 {deleted} 行")
    return {"ok": True, "deleted": deleted}


class SafetyMaterialExport(BaseModel):
    q: str = ""                      # 可选：按关键字导出匹配行（空=全表）
    filename: str = "安全钳物料号.xlsx"


@app.post("/api/safety_material/export")
def safety_material_export(payload: SafetyMaterialExport):
    """整表导出 xlsx（22万行走 openpyxl write_only 流式写，不进内存全量）。"""
    import openpyxl
    cols = _safety_material_cols()
    if not cols:
        raise HTTPException(status_code=500, detail="安全钳物料号列元数据缺失（meta.json）")
    keys = [c["key"] for c in cols]
    disps = [c["display"] for c in cols]
    wq, params = _safety_material_where(payload.q, keys)
    where = _safety_material_join(wq)
    sel = ", ".join([f'"{k}"' for k in keys])
    conn = _safety_material_conn()
    out = TMP_DIR / f"sfm_export_{uuid.uuid4().hex[:8]}.xlsx"
    try:
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet("安全钳物料号")
        ws.append(disps)
        cur = conn.cursor()
        B = 5000
        off = 0
        n_export = 0
        while True:
            rs = cur.execute(
                f'SELECT "row_no", {sel} FROM "material"{where} ORDER BY "row_no" LIMIT ? OFFSET ?',
                params + [B, off]).fetchall()
            if not rs:
                break
            for r in rs:
                ws.append([r["row_no"]] + [("" if r[keys[i]] is None else str(r[keys[i]])) for i in range(len(keys))])
                n_export += 1
            off += B
        wb.save(str(out))
    finally:
        conn.close()
    UI信息(f"安全钳物料号已导出 xlsx（{n_export} 行）")
    return FileResponse(str(out), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition":
                                 f"attachment; filename*=UTF-8''{urllib.parse.quote(payload.filename)}"})


# ============ 修改模块自动化 · 本地存储接口（2026-10-24，结构复用安全钳 store 模式） ============
# 数据真身本地 JSON（modify_store/local_store.json）；网页「修改模块自动化」5 个竖向页签
# 各对应一个 sheet，仅「修改代办」有种子（modify_seed/modify_seed.json：12 列 × 6 行，
# 与用户截图 2 逐格一致，含第 4 行重复物料号——属真实业务数据，种子不做去重）。
# 首次 GET /api/modify/store 时若 sheets 为空 → 读种子灌入「修改代办」（5 个 sheet 键都建，
# 其余 4 个空数组），并把每行 data["时间戳"] 重写为灌入时刻（种子里的固定值仅作占位）。

MODIFY_STORE_PATH = APP_DIR / "modify_store" / "local_store.json"
MODIFY_SEED_PATH = APP_DIR / "modify_seed" / "modify_seed.json"
MOD_SHEETS = ['修改代办', '修改合同状态', '修改计划', '修改计划汇总', '其他代办']   # sheet 白名单
MOD_SHEET_RENAMES = {'其他督办': '其他代办'}      # 2026-10-25 二改：页签改名（老库老键自动搬到新键）
MOD_BASE_SHEET = '修改代办'                     # 唯一有种子的 sheet
MOD_MAX_ROWS = 100000                           # 单次 replace/baseline 行数上限


def _modify_store_load() -> dict:
    if MODIFY_STORE_PATH.exists():
        try:
            return json.loads(MODIFY_STORE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"next_id": 1, "sheets": {}}


def _modify_store_save(store: dict) -> None:
    MODIFY_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    MODIFY_STORE_PATH.write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")


# ---- 本地存储「破坏性写入前自动快照」+ 一键恢复（2026-10-24：用户全选后批量删/导入覆盖，一秒钟能把整表清空）----
# 任何 delete / replace / baseline 之前先把当前 JSON 复制一份 `.snap-YYYYmmdd-HHMMSS`，保留最近 20 份；
# 前端「恢复」按钮拿最近一份覆盖回来（恢复前同样先快照当前状态，可来回切）。
SNAP_KEEP = 20


def _快照存储(路径: Path, 原因: str = "") -> Optional[Path]:
    """把本地 JSON 复制成 <名>.snap-<时间戳>；返回快照路径（失败返回 None，不影响主流程）。"""
    try:
        if not 路径.exists():
            return None
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        dst = 路径.with_name(f"{路径.name}.snap-{ts}")
        n = 1
        while dst.exists():   # 同一秒内多次快照（批量删连点）不互相覆盖
            dst = 路径.with_name(f"{路径.name}.snap-{ts}-{n}")
            n += 1
        shutil.copy2(路径, dst)
        snaps = sorted(路径.parent.glob(路径.name + ".snap-*"))
        for old in snaps[:-SNAP_KEEP]:
            try:
                old.unlink()
            except Exception:
                pass
        if 原因:
            print(f"[快照] {路径.name} → {dst.name}（{原因}）", flush=True)
        return dst
    except Exception as e:
        print(f"[快照失败] {路径}：{type(e).__name__}: {e}", flush=True)
        return None


def _快照清单(路径: Path) -> list:
    """最近在前：[{file, mtime, rows, sheets}]（rows = 各 sheet 行数合计，给前端显示「这份快照有多少行」）。
    时间取文件名里的时间戳（快照生成时刻），不取文件 mtime（copy2 会把源文件的旧 mtime 带过来，更具误导性）。"""
    out = []
    try:
        files = list(路径.parent.glob(路径.name + ".snap-*"))
        def 时刻(f):
            s = f.name.split(".snap-")[-1].replace("-", "")[:14]   # 20261024-222759 → 20261024222759
            try:
                return datetime.strptime(s, "%Y%m%d%H%M%S")
            except Exception:
                return datetime.fromtimestamp(f.stat().st_mtime)
        files.sort(key=时刻, reverse=True)
        for f in files:
            rows, sheets = 0, {}
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
                for k, v in (d.get("sheets") or {}).items():
                    sheets[k] = len(v or [])
                    rows += len(v or [])
            except Exception:
                pass
            out.append({"file": f.name, "mtime": 时刻(f).strftime("%Y-%m-%d %H:%M:%S"),
                        "size": f.stat().st_size, "rows": rows, "sheets": sheets})
    except Exception:
        pass
    return out


def _还原快照(路径: Path, 文件名: str) -> dict:
    """用指定快照覆盖当前存储（覆盖前先给当前状态也存一份快照）。文件名必须在本目录快照白名单内。"""
    if not 文件名 or "/" in 文件名 or "\\" in 文件名 or not 文件名.startswith(路径.name + ".snap-"):
        raise HTTPException(status_code=400, detail="快照文件名不合法")
    src = 路径.parent / 文件名
    if not src.exists():
        raise HTTPException(status_code=404, detail="快照不存在")
    try:
        json.loads(src.read_text(encoding="utf-8"))   # 先校验快照本身没坏
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"快照文件损坏，无法恢复：{type(e).__name__}")
    _快照存储(路径, "恢复前备份当前状态")
    shutil.copy2(src, 路径)
    d = json.loads(路径.read_text(encoding="utf-8"))
    return {"ok": True, "restored": 文件名, "rows": sum(len(v or []) for v in (d.get("sheets") or {}).values()),
            "sheets": {k: len(v or []) for k, v in (d.get("sheets") or {}).items()}}


def _modify_store_seed_if_empty() -> dict:
    """首次使用灌种子；老数据缺 sheet 键时只补空数组（不复活用户已删的数据）。"""
    store = _modify_store_load()
    # 页签改名迁移（2026-10-25 二改：「其他督办」→「其他代办」）——老键的行搬到新键后删掉老键
    改名过 = False
    for 老键, 新键 in MOD_SHEET_RENAMES.items():
        if 老键 in store.get("sheets", {}):
            老行 = store["sheets"].pop(老键) or []
            store["sheets"][新键] = (store["sheets"].get(新键) or []) + 老行
            改名过 = True
    if store.get("sheets"):
        缺键 = [s for s in MOD_SHEETS if s not in store["sheets"]]
        for s in 缺键:
            store["sheets"][s] = []
        if 缺键 or 改名过:
            _modify_store_save(store)
        return store
    store = {"next_id": 1, "sheets": {s: [] for s in MOD_SHEETS}}
    if MODIFY_SEED_PATH.exists():
        try:
            seed = json.loads(MODIFY_SEED_PATH.read_text(encoding="utf-8"))
        except Exception as e:
            # 种子文件损坏时不静默灌 0 行（会被误读成「平台无待办」）——明确告警
            print(f"[MODIFY][警告] 种子文件解析失败，本次灌入 0 行：{MODIFY_SEED_PATH}（{type(e).__name__}: {e}）", flush=True)
            UI信息(f"修改模块种子文件损坏，已灌入 0 行（{MODIFY_SEED_PATH.name}）")
            seed = {}
        nid = 1
        灌入时刻 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")   # 服务端本地时间
        rows = []
        for r in seed.get("rows", []):
            data = dict(r.get("data") or {})
            data["时间戳"] = 灌入时刻             # 覆盖种子文件里的占位时间戳
            rows.append({"id": nid, "row_no": r.get("row_no"), "data": data})
            nid += 1
        store["sheets"][MOD_BASE_SHEET] = rows
        store["next_id"] = nid
        _modify_store_save(store)
        UI信息(f"修改模块种子灌入「{MOD_BASE_SHEET}」{len(rows)} 行（时间戳={灌入时刻}）")
    else:
        _modify_store_save(store)
    return store


def _mod_sheet_or_400(sheet) -> str:
    """sheet 白名单校验（upsert/delete/replace/baseline 共用）。"""
    if not isinstance(sheet, str) or sheet not in MOD_SHEETS:
        raise HTTPException(status_code=400,
                            detail=f"sheet 必须为白名单之一：{'、'.join(MOD_SHEETS)}（收到 {sheet!r}）")
    return sheet


def _mod_id_or_400(v, 字段: str) -> int:
    if isinstance(v, bool) or not isinstance(v, int) or v < 0 or v > 10_000_000:
        raise HTTPException(status_code=400, detail=f"{字段} 必须为 0~10000000 的整数（收到 {v!r}）")
    return v


@app.get("/api/modify/store")
def modify_store_get():
    """修改模块本地存储全量（sheet → [{id,row_no,data}]）；首次调用触发种子灌入。"""
    return _modify_store_seed_if_empty()


@app.post("/api/modify/store/upsert")
def modify_store_upsert(payload: dict):
    """id=0 → 追加新行；id>0 → 按 id 更新 data（与 /api/safety/store 同参同结构）。"""
    sheet = _mod_sheet_or_400(payload.get("sheet"))
    data = payload.get("data")
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="data 必须为对象（列名→值）")
    row_id = _mod_id_or_400(payload.get("id") or 0, "id")
    if row_id:
        store = _modify_store_seed_if_empty()
        for r in store["sheets"].get(sheet, []):
            if r["id"] == row_id:
                r["data"] = data
                _modify_store_save(store)
                UI信息(f"修改模块本地存储修改 {sheet} 第 {r['row_no']} 行")
                return {"ok": True, "action": "update"}
        raise HTTPException(status_code=404, detail=f"未找到行 id={row_id}")
    row_no = payload.get("row_no")
    if isinstance(row_no, bool) or not isinstance(row_no, int) or row_no < 0 or row_no > 10_000_000:
        raise HTTPException(status_code=400, detail=f"row_no 必须为 0~10000000 的整数（收到 {row_no!r}）")
    store = _modify_store_seed_if_empty()
    rows = store["sheets"].setdefault(sheet, [])
    rid = store.get("next_id", 1)
    store["next_id"] = rid + 1
    rows.append({"id": rid, "row_no": row_no, "data": data})
    _modify_store_save(store)
    UI信息(f"修改模块本地存储新增 1 行（{sheet}）")
    return {"ok": True, "action": "insert"}


@app.get("/api/modify/store/snapshots")
def modify_store_snapshots():
    """最近 20 份「批量删除 / 导入覆盖」前的自动快照（最近在前），给前端「恢复」按钮列表用。"""
    当前 = _modify_store_load()
    return {
        "current": {"rows": sum(len(v or []) for v in (当前.get("sheets") or {}).values()),
                    "sheets": {k: len(v or []) for k, v in (当前.get("sheets") or {}).items()}},
        "snapshots": _快照清单(MODIFY_STORE_PATH),
    }


@app.post("/api/modify/store/restore")
def modify_store_restore(payload: dict):
    """用快照恢复（payload: {file} 省略则用最近一份）。恢复前会先把当前状态也存一份快照。"""
    file = (payload or {}).get("file")
    if not file:
        snaps = _快照清单(MODIFY_STORE_PATH)
        if not snaps:
            raise HTTPException(status_code=404, detail="还没有任何快照（本次服务启动后未做过批量删除/导入覆盖）")
        file = snaps[0]["file"]
    return _还原快照(MODIFY_STORE_PATH, file)


@app.post("/api/modify/store/delete")
def modify_store_delete(payload: dict):
    """删除行：id=单行；ids=批量多行（勾选删除用）。"""
    sheet = _mod_sheet_or_400(payload.get("sheet"))
    ids = payload.get("ids")
    if not isinstance(ids, list):
        row_id = payload.get("id")
        ids = [row_id] if row_id is not None else []
    if not ids:
        raise HTTPException(status_code=400, detail="缺少删除目标：请传 id（单行）或 ids（批量行 id 数组）")
    for v in ids:
        if isinstance(v, bool) or not isinstance(v, int) or v <= 0:
            raise HTTPException(status_code=400, detail=f"ids 必须为正整数数组（收到 {v!r}）")
    id_set = set(ids)
    store = _modify_store_seed_if_empty()
    rows = store["sheets"].get(sheet, [])
    if len(rows) and len(id_set) >= len(rows):
        # 整表一次性删光（用户全选后点「删除所选」）最危险 → 删前必须留快照，可一键恢复
        _快照存储(MODIFY_STORE_PATH, f"批量删除 {sheet} {len(id_set)} 行（整表）")
    elif len(id_set) > 50:
        _快照存储(MODIFY_STORE_PATH, f"批量删除 {sheet} {len(id_set)} 行")
    store["sheets"][sheet] = [r for r in rows if r["id"] not in id_set]
    _modify_store_save(store)
    return {"ok": True, "deleted": len(rows) - len(store["sheets"][sheet])}


@app.post("/api/modify/store/replace")
def modify_store_replace(payload: dict):
    """整表替换（导入-覆盖用）：清空该 sheet 后写入，id 重新自增。"""
    sheet = _mod_sheet_or_400(payload.get("sheet"))
    rows_in = payload.get("rows")
    if not isinstance(rows_in, list):
        raise HTTPException(status_code=400, detail="rows 必须为数组（[{row_no, data}]）")
    if len(rows_in) > MOD_MAX_ROWS:
        raise HTTPException(status_code=400, detail=f"rows 超过单次上限 {MOD_MAX_ROWS} 行")
    for i, item in enumerate(rows_in):
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail=f"第 {i + 1} 行必须为对象 {{row_no, data}}")
        rn = item.get("row_no")
        if rn is not None and (isinstance(rn, bool) or not isinstance(rn, int) or rn < 0):
            raise HTTPException(status_code=400, detail=f"第 {i + 1} 行 row_no 必须为非负整数")
        data = item.get("data")
        if data is not None and not isinstance(data, dict):
            raise HTTPException(status_code=400, detail=f"第 {i + 1} 行 data 必须为对象")
    store = _modify_store_seed_if_empty()
    # 导入-覆盖会把该 sheet 整表换掉（甚至换成 0 行）→ 覆盖前留快照
    旧行数 = len(store["sheets"].get(sheet, []) or [])
    if 旧行数:
        _快照存储(MODIFY_STORE_PATH, f"导入覆盖 {sheet}（原 {旧行数} 行 → {len(rows_in)} 行）")
    rows = []
    for item in rows_in:
        rid = store.get("next_id", 1)
        store["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": item.get("row_no"), "data": item.get("data") or {}})
    store["sheets"][sheet] = rows
    _modify_store_save(store)
    UI信息(f"修改模块本地存储覆盖导入 {sheet}（{len(rows)} 行）")
    return {"ok": True, "count": len(rows)}


@app.post("/api/modify/store/baseline")
def modify_store_baseline(payload: dict):
    """「通知并反馈」抓取结果增量落库（前端「保存为基准」按钮调用）。

    去重语义（与参考脚本 dedup_rows 一致，外加与库内已有物料号比对）：
      1) 批次内按「物料号」去重，保留首条；
      2) 再与该 sheet 已有行的「物料号」集合比对：已存在 → skipped，不存在 → 追加；
      3) 追加行 row_no = 当前 sheet 最大 row_no 续接；id 走全局 next_id；
      4) 「时间戳」一律由服务端生成（%Y-%m-%d %H:%M:%S 本地时间）并覆盖调用方传入值，
         防客户端时钟不一致；
      5) total = 追加后该 sheet 总行数。
    """
    sheet = _mod_sheet_or_400(payload.get("sheet"))
    rows_in = payload.get("rows")
    if not isinstance(rows_in, list):
        raise HTTPException(status_code=400, detail="rows 必须为数组（12 列对象数组）")
    if len(rows_in) > MOD_MAX_ROWS:
        raise HTTPException(status_code=400, detail=f"rows 超过单次上限 {MOD_MAX_ROWS} 行")
    for i, item in enumerate(rows_in):
        if not isinstance(item, dict):
            raise HTTPException(status_code=400, detail=f"第 {i + 1} 行必须为对象")
        if not str(item.get("物料号") or "").strip():
            raise HTTPException(status_code=400, detail=f"第 {i + 1} 行缺少必填列「物料号」")
    服务端时间 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    store = _modify_store_seed_if_empty()
    rows = store["sheets"].setdefault(sheet, [])

    def _行号值(r) -> int:
        v = r.get("row_no")
        return v if isinstance(v, int) and not isinstance(v, bool) else 0

    # 1) 批次内去重（保留首条）
    已见, 批次 = set(), []
    for item in rows_in:
        mat = str(item.get("物料号") or "").strip()
        if mat in 已见:
            continue
        已见.add(mat)
        批次.append(item)
    # 2) 与 sheet 已有物料号比对 → 存在跳过 / 不存在追加
    已有物料号 = {str((r.get("data") or {}).get("物料号") or "").strip() for r in rows}
    已有物料号.discard("")
    added = skipped = 0
    next_row_no = max([_行号值(r) for r in rows], default=0) + 1
    for item in 批次:
        mat = str(item.get("物料号") or "").strip()
        if mat in 已有物料号:
            skipped += 1
            continue
        data = dict(item)
        data["时间戳"] = 服务端时间       # 服务端生成并覆盖传入值
        rid = store.get("next_id", 1)
        store["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": next_row_no, "data": data})
        next_row_no += 1
        已有物料号.add(mat)              # 本轮已追加的物料号也计入，防同批次重复穿透
        added += 1
    _modify_store_save(store)
    UI信息(f"修改模块 baseline：{sheet} 新增 {added} 行、跳过 {skipped} 行（现有 {len(rows)} 行）")
    return {"ok": True, "added": added, "skipped": skipped,
            "total": len(rows), "timestamp": 服务端时间}


# ============ 权限管控 + 登录管理日志 · 本地存储接口（结构复用安全钳 store 模式） ============
# 权限表：账号/密码/拥有权限/可查看表（内容由用户后续配置；admin 行仅前端种子展示）
# 登录日志：登录账号/登录密码/登录时间/登录IP/主机类型（公司电脑→主机名；安卓/iPhone→设备类型）

PERM_STORE_PATH = APP_DIR / "perm_store" / "local_store.json"
LOGINLOG_STORE_PATH = APP_DIR / "loginlog_store" / "local_store.json"


def _generic_store_load(path) -> dict:
    if Path(path).exists():
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"next_id": 1, "sheets": {}}


def _generic_store_save(path, store: dict) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(store, ensure_ascii=False), encoding="utf-8")


def _generic_store_reorder(path, sheet: str, order: list) -> dict:
    """按给定 id 顺序重排一张表，并重写 row_no = 1..N（**保留原 id**，不动 data）。

    为什么要有它：store/upsert 命中 id 时只改 data、**不更新 row_no**，
    store/replace 又会**重分配全表 id**（把用户手填的其它列冲掉、把外部引用打乱）。
    「按某个字段重新排一遍顺序、但一行都不许丢也不许改」这种活，只能单独给一个接口。

    · order 里没出现的行，保持原相对次序接在后面（不会凭空丢行）。
    · 返回 {"ok":True,"count":总行数,"moved":顺序是否真的变了}。
    """
    store = _generic_store_load(path) or {}
    rows = list(store.get("sheets", {}).get(sheet, []))
    by_id = {}
    for r in rows:
        try:
            by_id[int(r.get("id") or 0)] = r
        except Exception:
            continue
    out, used = [], set()
    for rid in (order or []):
        try:
            key = int(rid)
        except Exception:
            continue
        r = by_id.get(key)
        if r is not None and key not in used:
            out.append(r)
            used.add(key)
    for r in rows:
        try:
            key = int(r.get("id") or 0)
        except Exception:
            key = None
        if key not in used:
            out.append(r)
            used.add(key)
    for i, r in enumerate(out, start=1):
        r["row_no"] = i
    store.setdefault("sheets", {})[sheet] = out
    _generic_store_save(path, store)
    old_order = [int(r.get("id") or 0) for r in rows]
    new_order = [int(r.get("id") or 0) for r in out]
    return {"ok": True, "count": len(out), "moved": old_order != new_order}


def _perm_seed_if_empty() -> dict:
    """首次使用：权限表灌入 admin 种子行（密码为空＝还没设，请用「重置密码」生成）。"""
    store = _generic_store_load(PERM_STORE_PATH)
    if store.get("sheets"):
        return store
    rows = [{"id": 1, "row_no": 1, "data": {"账号": "admin", "密码": "", "拥有权限": "全部", "可操作范围": "全部"}}]
    store = {"next_id": 2, "sheets": {"权限管控": rows}}
    _generic_store_save(PERM_STORE_PATH, store)
    return store


# ---------- 权限管控 · 密码单向加密（scrypt，不可逆） ----------
# 用户口径（2026-10-25，明令）：密码「只许写、不许看」——明文只在录入的那一瞬间存在于输入框里，
# 落盘立刻变成 scrypt(随机盐) 哈希，接口一律只回 ********，任何界面（包括 admin 自己）、导出文件、
# JSON 文件、日志里都拿不到原文，也无法反推。这样即使网站数据被人拷走，密码也解不出来。
#
# 忘记密码的兜底（都不需要记任何密钥）：
#   ① 页面「修改」弹窗里的「重置密码」→ 系统生成一次性随机密码，只回显一次（复制走），之后只有哈希；
#   ② 本机命令行：cd app && python reset_admin_password.py admin → 直接改文件里的哈希，打印新密码一次。
#   ③ 终极兜底：直接删 app/perm_store/local_store.json 里的账号行，或删整个文件（会重新灌 admin 种子行）。
_PW_MASK = "********"          # 接口对外统一占位符（前端再打码显示成 ••••••）
_PW_ALGO = "scrypt"
_PW_N, _PW_R, _PW_P, _PW_DKLEN = 2 ** 14, 8, 1, 32
_PW_MIN_LEN = 4                # 用户自己的本机工具，门槛不设太高（够挡一眼看穿就行）
_PW_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789"   # 去掉 0O1lI 易混字符
_SCOPE_KEY_OLD = "可查看表"     # 老列名（2026-10-25 改名）
_SCOPE_KEY_NEW = "可操作范围"


def _pw_hash(明文: str) -> str:
    """scrypt + 每条独立随机盐：同一条密码两次哈希结果也不同，无彩虹表可用。"""
    盐 = secrets.token_bytes(16)
    dk = hashlib.scrypt(明文.encode("utf-8"), salt=盐, n=_PW_N, r=_PW_R, p=_PW_P, dklen=_PW_DKLEN)
    return f"{_PW_ALGO}${_PW_N}${_PW_R}${_PW_P}${base64.b64encode(盐).decode()}${base64.b64encode(dk).decode()}"


def _pw_is_hash(v) -> bool:
    return isinstance(v, str) and v.startswith(_PW_ALGO + "$") and v.count("$") == 5


def _pw_verify(明文: str, 存储值: str) -> bool:
    """常数时间比对（hmac.compare_digest），防时序侧信道。"""
    if not _pw_is_hash(存储值):
        return False
    try:
        _, n, r, p, 盐b64, dkb64 = 存储值.split("$")
        盐 = base64.b64decode(盐b64)
        dk0 = base64.b64decode(dkb64)
        dk = hashlib.scrypt(str(明文).encode("utf-8"), salt=盐, n=int(n), r=int(r), p=int(p), dklen=len(dk0))
        return hmac.compare_digest(dk, dk0)
    except Exception:
        return False


def _perm_load_migrated() -> dict:
    """读权限表并把老数据就地升级（幂等，每次读都跑，代价只有几十行 JSON）：
    ① 列名「可查看表」→「可操作范围」；
    ② 任何明文密码 → 立即哈希（用户 2026-10-25 之前录的明文口令，第一次打开就永久变成哈希）。
    """
    store = _generic_store_load(PERM_STORE_PATH)
    if not store.get("sheets"):
        store = _perm_seed_if_empty()
    改了 = False
    for _sheet, rows in (store.get("sheets") or {}).items():
        for r in rows:
            d = r.setdefault("data", {})
            if _SCOPE_KEY_OLD in d:
                d[_SCOPE_KEY_NEW] = d.pop(_SCOPE_KEY_OLD)
                改了 = True
            pw = d.get("密码")
            if isinstance(pw, str) and pw.strip() and pw != _PW_MASK and not _pw_is_hash(pw):
                d["密码"] = _pw_hash(pw.strip())
                改了 = True
    if 改了:
        _generic_store_save(PERM_STORE_PATH, store)
        UI信息("权限管控：已就地升级老数据（列名→可操作范围；明文密码→scrypt 哈希，不可逆）")
    return store


def _perm_public(store: dict) -> dict:
    """对外视图：密码列一律只回掩码，绝不外泄哈希或原文。"""
    out = json.loads(json.dumps(store))
    for rows in (out.get("sheets") or {}).values():
        for r in rows:
            d = r.get("data") or {}
            if _SCOPE_KEY_OLD in d:
                d[_SCOPE_KEY_NEW] = d.pop(_SCOPE_KEY_OLD)
            if d.get("密码"):
                d["密码"] = _PW_MASK
    return out


def _perm_find(store: dict, 账号: str = "", row_id: int = 0):
    for r in (store.get("sheets") or {}).get("权限管控", []):
        d = r.get("data") or {}
        if (账号 and str(d.get("账号", "")).strip() == 账号.strip()) or (row_id and r.get("id") == row_id):
            return r
    return None


class PwResetReq(BaseModel):
    """重置密码：只给账号名，系统生成新密码（明文只在响应里出现一次）。"""
    account: str
    length: int = 12


class PwSetReq(BaseModel):
    account: str
    password: str


class PwVerifyReq(BaseModel):
    account: str
    password: str


class PermUpsert(BaseModel):
    sheet: str
    id: int = 0
    row_no: int
    data: dict


class PermReplace(BaseModel):
    sheet: str
    rows: List[dict]


def _store_routes(prefix: str, store_path, seed_fn, label: str):
    """为 perm/loginlog 生成同构的 store 路由（结构=安全钳 store）。"""
    app.state.__dict__  # noqa: 保引用

    def load():
        s = _generic_store_load(store_path)
        if not s.get("sheets"):
            s = seed_fn()
        return s

    def upsert(payload: PermUpsert):
        s = load()
        rows = s["sheets"].setdefault(payload.sheet, [])
        if payload.id:
            for r in rows:
                if r["id"] == payload.id:
                    r["data"] = payload.data
                    _generic_store_save(store_path, s)
                    UI信息(f"{label}修改 {payload.sheet} 第 {r['row_no']} 行")
                    return {"ok": True, "action": "update"}
            raise HTTPException(status_code=404, detail=f"未找到行 id={payload.id}")
        rid = s.get("next_id", 1)
        s["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": payload.row_no, "data": payload.data})
        _generic_store_save(store_path, s)
        UI信息(f"{label}新增 1 行（{payload.sheet}）")
        return {"ok": True, "action": "insert"}

    def delete(payload: dict):
        sheet = payload.get("sheet")
        ids = payload.get("ids")
        if not isinstance(ids, list):
            row_id = payload.get("id")
            ids = [row_id] if row_id is not None else []
        id_set = set(ids)
        s = load()
        rows = s["sheets"].get(sheet, [])
        s["sheets"][sheet] = [r for r in rows if r["id"] not in id_set]
        _generic_store_save(store_path, s)
        return {"ok": True, "deleted": len(rows) - len(s["sheets"][sheet])}

    def replace(payload: PermReplace):
        s = load()
        rows = []
        for item in payload.rows:
            rid = s.get("next_id", 1)
            s["next_id"] = rid + 1
            rows.append({"id": rid, "row_no": item.get("row_no"), "data": item.get("data") or {}})
        s["sheets"][payload.sheet] = rows
        _generic_store_save(store_path, s)
        UI信息(f"{label}覆盖导入 {payload.sheet}（{len(rows)} 行）")
        return {"ok": True, "count": len(rows)}

    return load, upsert, delete, replace


_perm_load, _perm_upsert, _perm_delete, _perm_replace = _store_routes(
    "perm", PERM_STORE_PATH, _perm_seed_if_empty, "权限管控")


@app.get("/api/perm/store")
def perm_store_get():
    """权限管控全量（密码列永远只回掩码；哈希不进前端）。"""
    return _perm_public(_perm_load_migrated())


@app.post("/api/perm/store/upsert")
def perm_store_upsert(payload: PermUpsert):
    """写入一行。密码规则（关键）：
    空 / 掩码 → 视为「不修改」，保留库里原哈希（前端表格里只有掩码，改权限时必然带回掩码）；
    明文 → 立即 scrypt 哈希后写库；库里已是哈希的（比如导入导出的文件）原样保留，不重复哈希。
    """
    store = _perm_load_migrated()
    data = dict(payload.data or {})
    if _SCOPE_KEY_OLD in data:
        data[_SCOPE_KEY_NEW] = data.pop(_SCOPE_KEY_OLD)
    pw = str(data.get("密码", "") or "").strip()
    if pw in ("", _PW_MASK, "••••••", "••••••••"):
        old = _perm_find(store, str(data.get("账号", "") or ""), payload.id)
        原哈希 = ((old or {}).get("data") or {}).get("密码", "")
        if 原哈希:
            data["密码"] = 原哈希
        else:
            data.pop("密码", None)
    elif not _pw_is_hash(pw):
        if len(pw) < _PW_MIN_LEN:
            raise HTTPException(status_code=400, detail=f"密码至少 {_PW_MIN_LEN} 位")
        data["密码"] = _pw_hash(pw)
    payload.data = data
    _perm_touch()
    return _perm_upsert(payload)


@app.post("/api/perm/store/delete")
def perm_store_delete(payload: dict):
    """删除行；admin 是内置超级账号，禁止删除（否则权限表会被清空到没人能管）。"""
    store = _perm_load_migrated()
    ids = payload.get("ids")
    if not isinstance(ids, list):
        ids = [payload.get("id")]
    id_set = {i for i in ids if i is not None}
    for r in (store.get("sheets") or {}).get("权限管控", []):
        if r.get("id") in id_set and str((r.get("data") or {}).get("账号", "")) == "admin":
            raise HTTPException(status_code=400, detail="admin 是内置超级账号，不允许删除")
    _perm_touch()
    return _perm_delete(payload)


@app.post("/api/perm/store/replace")
def perm_store_replace(payload: PermReplace):
    rows = []
    for item in (payload.rows or []):
        d = dict(item.get("data") or {})
        if _SCOPE_KEY_OLD in d:
            d[_SCOPE_KEY_NEW] = d.pop(_SCOPE_KEY_OLD)
        pw = str(d.get("密码", "") or "").strip()
        if pw and pw not in (_PW_MASK, "••••••") and not _pw_is_hash(pw):
            d["密码"] = _pw_hash(pw)
        elif pw in (_PW_MASK, "••••••"):
            d.pop("密码", None)
        rows.append({"row_no": item.get("row_no"), "data": d})
    payload.rows = rows
    _perm_touch()
    return _perm_replace(payload)


@app.post("/api/perm/password/reset")
def perm_password_reset(req: PwResetReq):
    """系统生成一次性随机密码：明文只在本响应里出现一次（前端复制走），库里只留哈希。"""
    store = _perm_load_migrated()
    row = _perm_find(store, req.account.strip())
    if not row:
        raise HTTPException(status_code=404, detail=f"未找到账号 {req.account}")
    n = max(6, min(32, int(req.length or 12)))
    明文 = "".join(secrets.choice(_PW_ALPHABET) for _ in range(n))
    row.setdefault("data", {})["密码"] = _pw_hash(明文)
    _generic_store_save(PERM_STORE_PATH, store)
    UI信息(f"权限管控：已重置账号 {req.account} 的密码（明文只回显这一次，不落盘）")
    return {"ok": True, "account": req.account, "password": 明文}


@app.post("/api/perm/password/set")
def perm_password_set(req: PwSetReq):
    """按账号直接设新密码（明文进、哈希出，明文不落盘）。"""
    if len(str(req.password or "")) < _PW_MIN_LEN:
        raise HTTPException(status_code=400, detail=f"密码至少 {_PW_MIN_LEN} 位")
    store = _perm_load_migrated()
    row = _perm_find(store, req.account.strip())
    if not row:
        raise HTTPException(status_code=404, detail=f"未找到账号 {req.account}")
    row.setdefault("data", {})["密码"] = _pw_hash(str(req.password))
    _generic_store_save(PERM_STORE_PATH, store)
    UI信息(f"权限管控：账号 {req.account} 已设置新密码（scrypt 哈希）")
    return {"ok": True}


@app.post("/api/perm/password/verify")
def perm_password_verify(req: PwVerifyReq):
    """校验密码（后续接登录页时直接用这个；现在也能验证重置是否生效）。只回 ok，不泄漏任何信息。"""
    store = _perm_load_migrated()
    row = _perm_find(store, req.account.strip())
    if not row:
        return {"ok": False, "reason": "账号不存在"}
    存 = str((row.get("data") or {}).get("密码", "") or "")
    if not 存:
        return {"ok": False, "reason": "该账号还没设过密码"}
    return {"ok": _pw_verify(str(req.password or ""), 存)}


_loginlog_load, _loginlog_upsert, _loginlog_delete, _loginlog_replace = _store_routes(
    "loginlog", LOGINLOG_STORE_PATH, lambda: {"next_id": 1, "sheets": {}}, "登录日志")


@app.get("/api/loginlog/store")
def loginlog_store_get():
    """登录管理日志本地存储全量。

    2026-10-25 用户要求：按登录时间倒序，最新登录排在最前面。
    实现要点：
      ①先按写盘顺序反转、再按「登录时间」稳定降序——同一分钟内（时间精度只到分钟）后登录的仍在前；
      ②返回副本里把「行号」重编成 1..N（1＝最新一次登录），不写回磁盘、不影响任何行的 id。
    """
    s = _loginlog_load()
    rows = list((s.get("sheets") or {}).get("登录管理日志") or [])
    if rows:
        rows = sorted(list(reversed(rows)),
                      key=lambda r: str((r.get("data") or {}).get("登录时间") or ""),
                      reverse=True)
        rows = [{**r, "row_no": i} for i, r in enumerate(rows, 1)]
        s["sheets"]["登录管理日志"] = rows
    return s


@app.post("/api/loginlog/store/upsert")
def loginlog_store_upsert(payload: PermUpsert):
    return _loginlog_upsert(payload)


@app.post("/api/loginlog/store/delete")
def loginlog_store_delete(payload: dict):
    return _loginlog_delete(payload)


@app.post("/api/loginlog/store/replace")
def loginlog_store_replace(payload: PermReplace):
    return _loginlog_replace(payload)


# ============ 登录门禁（账号密码 + 钉钉扫码，2026-10-25）============
# 用户要求：进站先登录，可以输账号密码，也可以用钉钉扫码。
# 门禁本身在 auth_gate.py：会话 cookie + 中间件（未登录→/login；已登录按「可操作范围」拦接口）。
def _perm_touch() -> None:
    """权限表被改过之后，让登录门禁的账号缓存立即失效（否则新建/删除账号 1 秒内仍按旧表判断）。"""
    try:
        import auth_gate as _ag
        _ag.invalidate_perm_cache()
    except Exception:
        pass


def _perm_upsert_row(row: dict) -> None:
    """把权限表里的一行原样写回盘（登录门禁改密码 / 初始化 admin 密码用）。"""
    store = _perm_load_migrated()
    for r in (store.get("sheets") or {}).get("权限管控", []):
        if r.get("id") == row.get("id"):
            r["data"] = row.get("data") or {}
            break
    _generic_store_save(PERM_STORE_PATH, store)
    try:
        import auth_gate as _ag2
        _ag2.invalidate_perm_cache()
    except Exception:
        pass


try:
    import auth_gate as _auth_gate
    _auth_gate.mount_auth(
        app,
        app_dir=APP_DIR,
        pw_hash=_pw_hash,
        pw_verify=_pw_verify,
        perm_load=_perm_load_migrated,
        perm_upsert_row=_perm_upsert_row,
        loginlog_load=_loginlog_load,
        loginlog_save=lambda s: _generic_store_save(LOGINLOG_STORE_PATH, s),
        ui=UI信息,
    )
except Exception as _auth_e:      # 挂载失败不影响其它功能（宁可先能用）
    print(f"[AUTH] 登录门禁挂载异常，已跳过：{type(_auth_e).__name__}: {_auth_e}")


# ============ 钉钉邮箱群消息配置模块 · 本地存储接口（结构复用 perm/loginlog store 模式） ============
# 三个工作表：钉钉邮箱配置 / 群消息配置 / 常用人员配置
# 数据由菱工按用户近12个月真实收发行为统计生成种子（ddconfig_store/local_store.json），可在网页上增删改。

DDCONFIG_STORE_PATH = APP_DIR / "ddconfig_store" / "local_store.json"


def _ddconfig_seed_if_empty() -> dict:
    """首次使用：若无种子文件则给空结构（种子由菱工统计生成后放入 ddconfig_store）。"""
    s = _generic_store_load(DDCONFIG_STORE_PATH)
    if not s.get("sheets"):
        s = {"next_id": 1, "sheets": {}}
        _generic_store_save(DDCONFIG_STORE_PATH, s)
    return s


_ddconfig_load, _ddconfig_upsert, _ddconfig_delete, _ddconfig_replace = _store_routes(
    "ddconfig", DDCONFIG_STORE_PATH, _ddconfig_seed_if_empty, "钉钉邮箱群消息配置")


@app.get("/api/ddconfig/store")
def ddconfig_store_get():
    """钉钉邮箱群消息配置本地存储全量。"""
    return _ddconfig_load()


@app.post("/api/ddconfig/store/upsert")
def ddconfig_store_upsert(payload: PermUpsert):
    return _ddconfig_upsert(payload)


@app.post("/api/ddconfig/store/delete")
def ddconfig_store_delete(payload: dict):
    return _ddconfig_delete(payload)


@app.post("/api/ddconfig/store/replace")
def ddconfig_store_replace(payload: PermReplace):
    return _ddconfig_replace(payload)


# ============ 触发器模块 · 本地存储接口（2026-10-20，复刻影刀RPA触发器管理） ============
# 侧边栏「触发器(m11)」两个页签：
#   计划任务：行 data = {名称, 类型(定时/文件/热键/邮件), 应用, 启用, config{按类型存完整配置}}
#   运行日志：行 data = {应用, 开始时间, 触发方式, 运行结果, 主题, 日志:[文本行], 截图:[{t,title,url}]}
# 运行错误处理：store["settings"] = {邮件通知, 收件箱[], 发送方式("钉钉邮箱"，固定), 
#   发送内容{截图,最新运行日志,主题,正文内容,主题文本,正文文本}, 钉钉通知}
TRIGGER_STORE_PATH = APP_DIR / "trigger_store" / "local_store.json"
TRIGGER_SHOTS_DIR = APP_DIR / "trigger_store" / "shots"

# 计划任务种子（照用户影刀RPA「计划任务」截图现值抄录，2026-10-20）
_TRIGGER_SEED_TASKS = [
    {"名称": "18.20维护物料", "类型": "定时", "应用": "自动触发维护物料结果", "启用": False,
     "config": {"freq": "每天", "time": "18:20", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "060结余统计", "类型": "定时", "应用": "060结余定时抓取", "启用": False,
     "config": {"freq": "每天", "time": "21:40", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "项目1web网页端自动抓取数据及存储", "类型": "定时", "应用": "1205网页端数据自动化获取与处理（基于影刀和Prep零代码实现）", "启用": False,
     "config": {"freq": "高级", "cronDesc": "1月每周五的21时00分执行", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "维护物料", "类型": "定时", "应用": "自动", "启用": False,
     "config": {"freq": "每天", "time": "20:21", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "短信邮件群发", "类型": "邮件", "应用": "自动触发维护物料结果", "启用": False,
     "config": {"server": "163邮箱", "account": "", "authcode": "", "folder": "INBOX", "queue": False, "timeout": 0,
                "rules": {"from": "", "to": "", "subject": "", "body": ""}}},
    {"名称": "定时维护物料", "类型": "定时", "应用": "自动", "启用": False,
     "config": {"freq": "每天", "time": "15:30", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "按键触发", "类型": "热键", "应用": "安全钳APD一点一点", "启用": False,
     "config": {"ctrl": True, "shift": False, "alt": False, "key": "F1", "queue": False, "timeout": 0}},
    {"名称": "定时修取数据", "类型": "定时", "应用": "间歇东京数据", "启用": False,
     "config": {"freq": "每天", "time": "22:51", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
    {"名称": "223", "类型": "定时", "应用": "060结余定时抓取", "启用": False,
     "config": {"freq": "高级", "cronDesc": "每间隔5分钟执行", "endAfter": False, "calendar": False, "queue": True, "timeout": 40}},
]

# 运行日志种子（照用户影刀RPA「运行日志」截图现值抄录；截图记录为占位条目，真实截图由运行期落盘 shots/）
def _触发器日志行(应用, 开始时间, 触发方式, 运行结果, 主题, 细节, 截图标题):
    return {"应用": 应用, "开始时间": 开始时间, "触发方式": 触发方式, "运行结果": 运行结果,
            "主题": 主题, "错误次数": 1 if 运行结果 == "运行失败" else 0,
            "日志": 细节, "截图": [{"t": "09:26:05", "title": 截图标题[0]}, {"t": "09:26:12", "title": 截图标题[1]}]}


_TRIGGER_SEED_LOGS = [
    _触发器日志行("GEO · 千问", "2025年10月13日 09:26", "[手动运行]", "已取消", "GEO · 千问 · 手动运行",
              ["09:26:01 触发器就绪，开始手动运行", "09:26:03 加载应用配置完成", "09:26:05 用户取消运行，流程中断"],
              ["运行开始 · 应用首页", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月27日 21:40", "[邮件触发器] 03箱APD", "已取消", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 匹配成功：发件人/主题均命中", "21:40:03 用户取消运行，流程中断"],
              ["邮件匹配详情", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月26日 21:40", "[邮件触发器] 03箱APD", "运行中", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 应用启动，正在执行选择屏条件", "21:40:05 正在导出数据，请稍候…"],
              ["应用启动画面", "选择屏条件"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月25日 21:40", "[邮件触发器] 03箱APD", "运行失败", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 应用启动", "21:40:06 ✗ 节点「03箱APD发布」失败：目标元素未找到（超时30s）",
               "21:40:07 已重试 1 次仍失败，流程中断", "21:40:07 已按运行错误处理设置发送失败通知"],
              ["失败瞬间截图", "错误提示弹窗"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月24日 21:40", "[邮件触发器] 03箱APD", "已取消", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 用户取消运行，流程中断"],
              ["邮件匹配详情", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月23日 21:40", "[手动运行]", "已取消", "安全钳APD一点一点 · 手动运行",
              ["21:40:01 触发器就绪，开始手动运行", "21:40:02 用户取消运行，流程中断"],
              ["运行开始 · 应用首页", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月22日 21:40", "[手动运行]", "运行成功", "安全钳APD一点一点 · 手动运行",
              ["21:40:01 触发器就绪，开始手动运行", "21:40:02 应用启动", "21:40:06 执行选择屏条件完成", "21:40:09 数据导出并写回网页存储", "21:40:10 ✓ 运行完成，耗时 8.6s"],
              ["选择屏条件", "运行完成画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月21日 21:40", "[邮件触发器] 03箱APD", "已取消", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 用户取消运行，流程中断"],
              ["邮件匹配详情", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月20日 21:40", "[邮件触发器] 03箱APD", "已取消", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 用户取消运行，流程中断"],
              ["邮件匹配详情", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月19日 21:40", "[手动运行]", "已取消", "安全钳APD一点一点 · 手动运行",
              ["21:40:01 触发器就绪，开始手动运行", "21:40:02 用户取消运行，流程中断"],
              ["运行开始 · 应用首页", "取消前的最后画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月18日 21:40", "[邮件触发器] 03箱APD", "运行成功", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 应用启动", "21:40:07 执行选择屏条件完成", "21:40:11 ✓ 运行完成，耗时 9.8s"],
              ["选择屏条件", "运行完成画面"]),
    _触发器日志行("安全钳APD一点一点", "2025年04月17日 21:40", "[邮件触发器] 03箱APD", "已取消", "03箱APD · 邮件触发",
              ["21:40:01 收到触发邮件，解析匹配规则", "21:40:02 用户取消运行，流程中断"],
              ["邮件匹配详情", "取消前的最后画面"]),
]

_TRIGGER_ERR_DEFAULT = {
    "邮件通知": True, "收件箱": ["demo.user53@example.com"], "发送方式": "钉钉邮箱",
    "发送内容": {"截图": True, "最新运行日志": True, "主题": True, "正文内容": False,
                 "主题文本": "", "正文文本": ""},
    "钉钉通知": True,
}


def _trigger_seed_if_empty() -> dict:
    """首次使用：灌入计划任务 9 条 + 运行日志 12 条 + 默认运行错误处理设置。"""
    s = _generic_store_load(TRIGGER_STORE_PATH)
    if s.get("sheets", {}).get("计划任务") or s.get("sheets", {}).get("运行日志"):
        if "settings" not in s:
            s["settings"] = dict(_TRIGGER_ERR_DEFAULT)
            _generic_store_save(TRIGGER_STORE_PATH, s)
        return s
    s = {"next_id": 1, "sheets": {"计划任务": [], "运行日志": []}, "settings": dict(_TRIGGER_ERR_DEFAULT)}
    for d in _TRIGGER_SEED_TASKS:
        s["sheets"]["计划任务"].append({"id": s["next_id"], "row_no": len(s["sheets"]["计划任务"]) + 1, "data": d})
        s["next_id"] += 1
    for d in _TRIGGER_SEED_LOGS:
        s["sheets"]["运行日志"].append({"id": s["next_id"], "row_no": len(s["sheets"]["运行日志"]) + 1, "data": d})
        s["next_id"] += 1
    _generic_store_save(TRIGGER_STORE_PATH, s)
    UI信息(f"触发器种子灌入：计划任务 {len(_TRIGGER_SEED_TASKS)} 条、运行日志 {len(_TRIGGER_SEED_LOGS)} 条")
    return s


_trigger_load, _trigger_upsert, _trigger_delete, _trigger_replace = _store_routes(
    "trigger", TRIGGER_STORE_PATH, _trigger_seed_if_empty, "触发器")


@app.get("/api/trigger/store")
def trigger_store_get():
    """触发器本地存储全量（sheets: 计划任务/运行日志 + settings: 运行错误处理）。"""
    return _trigger_load()


@app.post("/api/trigger/store/upsert")
def trigger_store_upsert(payload: PermUpsert):
    return _trigger_upsert(payload)


@app.post("/api/trigger/store/delete")
def trigger_store_delete(payload: dict):
    return _trigger_delete(payload)


@app.post("/api/trigger/store/replace")
def trigger_store_replace(payload: PermReplace):
    return _trigger_replace(payload)


@app.get("/api/trigger/error-handling")
def trigger_err_get():
    """运行错误处理设置（邮件通知收件箱/发送内容（截图、最新运行日志、主题/正文自定义）/钉钉通知）。"""
    s = _trigger_load()
    return {"ok": True, "settings": s.get("settings") or dict(_TRIGGER_ERR_DEFAULT)}


class TriggerErrHandling(BaseModel):
    settings: dict


@app.post("/api/trigger/error-handling")
def trigger_err_save(payload: TriggerErrHandling):
    s = _trigger_load()
    s["settings"] = payload.settings or dict(_TRIGGER_ERR_DEFAULT)
    _generic_store_save(TRIGGER_STORE_PATH, s)
    UI信息("触发器运行错误处理设置已保存")
    return {"ok": True}


class TriggerNotify(BaseModel):
    kind: str = "test"          # test=测试邮件（唯一入口，背后统一走 dws 钉钉邮箱链路）
    settings: dict = {}


def _通知_收集截图(最新: dict) -> list:
    """截图来源：① trigger_store/shots 目录最新 PNG；② 最新运行日志行 截图[].url 指向的本地存在的文件。"""
    files: list = []
    try:
        for p in sorted(TRIGGER_SHOTS_DIR.glob("*.png"), key=lambda x: x.stat().st_mtime, reverse=True)[:3]:
            files.append(p)
    except Exception:
        pass
    for shot in (最新.get("截图") or []):
        url = str(shot.get("url") or "").strip()
        if not url:
            continue
        try:
            cand = Path(url)
            if cand.is_file() and cand.suffix.lower() in (".png", ".jpg", ".jpeg") and cand not in files:
                files.append(cand)
        except Exception:
            continue
    return files[:5]


@app.post("/api/trigger/notify")
def trigger_notify(payload: TriggerNotify):
    """运行错误处理 · 测试邮件（唯一发送入口）：
    统一走既有 dws_mail 钉钉邮箱链路（发送方式已固定为「钉钉邮箱发送（默认）」）。
    发送内容按勾选项组装：截图（内联附件）+ 最新运行日志 + 主题（可自定义主题文本）+ 正文内容（可自定义正文文本）。
    """
    st = payload.settings or {}
    s = _trigger_load()
    日志行 = (s.get("sheets", {}).get("运行日志") or [])
    最新 = 日志行[0]["data"] if 日志行 else {}
    内容 = st.get("发送内容") or {}
    自定义主题 = str(内容.get("主题文本") or "").strip()
    if 内容.get("主题", True) and 自定义主题:
        主题 = 自定义主题
    else:
        主题 = (最新.get("主题") or "触发器运行通知") + "（测试）"
    行 = ["【03#箱工作站 · 触发器运行错误处理 · 测试邮件】", ""]
    自定义正文 = str(内容.get("正文文本") or "").strip()
    if 内容.get("正文内容", False) and 自定义正文:
        行 += [自定义正文, ""]
    if 内容.get("主题", True):
        行.append(f"主题：{自定义主题 or (最新.get('主题') or '—')}")
    if 内容.get("最新运行日志", True):
        行 += ["", "最新运行日志：", f"应用：{最新.get('应用') or '—'}", f"开始时间：{最新.get('开始时间') or '—'}",
               f"触发方式：{最新.get('触发方式') or '—'}", f"运行结果：{最新.get('运行结果') or '—'}", ""]
        行 += ["  " + str(x) for x in (最新.get("日志") or [])]
    截图文件 = _通知_收集截图(最新) if 内容.get("截图", True) else []
    if 内容.get("截图", True):
        行 += ["", "截图记录：" + ("、".join(x.get("title", "") for x in (最新.get("截图") or [])) or "（暂无截图记录）")]

    dws_mail = _rpa导dws()
    if dws_mail is None:
        return {"ok": False, "msg": "dws_mail 不可用（web_auto/公共_钉钉通知 拷贝是否完整）"}
    收件 = [x for x in (st.get("收件箱") or []) if str(x).strip()]
    try:
        # dws_mail.发送邮件 返回 (状态, 说明)，状态 ∈ {"已发送","已落盘","彻底失败"}
        kwargs = {"内联图片": 截图文件 or None}
        if 收件:
            kwargs["收件人"] = ", ".join(收件)
        状态, 说明 = dws_mail.发送邮件(主题, "\n".join(行), **kwargs)
        图注 = f"；截图 {len(截图文件)} 张" if 内容.get("截图", True) else "；未勾选截图"
        if 内容.get("截图", True) and not 截图文件:
            图注 += "（shots 目录与日志截图均无可用图片文件）"
        return {"ok": 状态 == "已发送", "msg": f"钉钉邮箱发送：{说明}{图注}"}
    except Exception as exc:
        return {"ok": False, "msg": f"钉钉邮箱发送失败：{exc}"}


# ============ 本机 Windows 计划任务采集 / 快照同步（2026-10-20；触发器「电脑计划任务同步」应用） ============
# 背景：电脑（任务计划程序）里的计划任务迁移进网页「触发器 → 计划任务」。
#   采集：GET  /api/system/win-tasks         → PowerShell Get-ScheduledTask 结构化列表（中文状态/触发类型）
#   落库：POST /api/system/win-tasks/snapshot → 写入触发器存储「电脑计划任务」sheet（全量明细+同步时间）
# 电脑现有任务 9/11 为时间触发，故网页侧计划任务类型取「定时」，应用 = AutoTaskKit 任务 win_task_sync。
WIN_TASKS_PS1 = APP_DIR / "collect_win_tasks.ps1"

_WIN_TASK_STATE_CN = {"Ready": "就绪", "Disabled": "已禁用", "Queued": "已排队", "Running": "运行中"}
_WIN_TRIG_TYPE_CN = {
    "Daily": "每天", "Time": "一次", "Weekly": "每周", "Monthly": "每月",
    "Logon": "登录时", "Boot": "开机时", "Registration": "注册时", "Event": "事件触发",
    "SessionStateChange": "会话状态变化", "Idle": "空闲时",
}


def _重复间隔中文(iso: str) -> str:
    """PT3H30M → 3小时30分钟；PT1H → 1小时；PT30M → 30分钟。"""
    s = str(iso or "").replace("PT", "")
    m = _re.match(r"(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", s)
    if not m:
        return s
    h, mi, sec = m.groups()
    return ((h + "小时") if h else "") + ((mi + "分钟") if mi else "") + ((sec + "秒") if sec else "")


@app.get("/api/system/win-tasks")
def system_win_tasks():
    """采集本机 Windows 计划任务（任务计划程序 root 级 + 全部子路径），返回结构化列表。"""
    if not WIN_TASKS_PS1.exists():
        raise HTTPException(status_code=500, detail=f"采集脚本缺失：{WIN_TASKS_PS1.name}")
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(WIN_TASKS_PS1)],
            capture_output=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"采集失败：{exc}")
    out = (proc.stdout or b"").decode("utf-8", errors="replace").strip()
    if not out:
        err = (proc.stderr or b"").decode("utf-8", errors="replace").strip()
        raise HTTPException(status_code=500, detail=f"采集无输出：{err[:200]}")
    try:
        raw = json.loads(out)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"结果解析失败：{exc}")
    tasks = raw if isinstance(raw, list) else [raw]
    result = []
    for t in tasks:
        触发类型, 触发段 = [], []
        for g in (t.get("triggers") or []):
            类型 = _WIN_TRIG_TYPE_CN.get(g.get("type"), g.get("type") or "其他")
            触发类型.append(类型)
            段 = [类型]
            if g.get("start"):
                段.append(str(g["start"]))
            if g.get("daysOfWeek"):
                段.append("星期" + str(g["daysOfWeek"]))
            if g.get("repeatInterval"):
                段.append("每" + _重复间隔中文(g["repeatInterval"]) + "重复")
            触发段.append(" ".join(段))
        result.append({
            "名称": t.get("taskName") or "",
            "路径": t.get("taskPath") or "",
            "状态": _WIN_TASK_STATE_CN.get(t.get("state"), t.get("state") or ""),
            "触发类型": 触发类型[0] if 触发类型 else "无（仅手动）",
            "触发详情": "；".join(触发段) if 触发段 else "无触发器（手动/其他方式启动）",
            "动作": "；".join(t.get("actions") or []) or "—",
        })
    result.sort(key=lambda x: (x["路径"], x["名称"]))
    return {"ok": True, "count": len(result), "tasks": result}


class WinTasksSnapshot(BaseModel):
    tasks: list = []


@app.post("/api/system/win-tasks/snapshot")
def win_tasks_snapshot(payload: WinTasksSnapshot):
    """把采集到的本机计划任务写入触发器存储「电脑计划任务」sheet（迁移落地，全量覆盖）。"""
    s = _trigger_load()
    rows = []
    for i, t in enumerate(payload.tasks, 1):
        rid = s.get("next_id", 1)
        s["next_id"] = rid + 1
        rows.append({"id": rid, "row_no": i, "data": dict(t)})
    s["sheets"]["电脑计划任务"] = rows
    s["win_tasks_meta"] = {"上次同步时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "任务数": len(rows)}
    _generic_store_save(TRIGGER_STORE_PATH, s)
    UI信息(f"电脑计划任务快照同步：{len(rows)} 行")
    return {"ok": True, "count": len(rows)}


@app.post("/api/trigger/shot")
def trigger_shot_save(payload: dict):
    """运行日志截图归档（2026-10-20）：前端 html2canvas 采集 → base64 PNG
    → trigger_store/shots/，返回可访问 URL，写进日志行 截图[].url（真实截图，非占位）。"""
    image = str((payload or {}).get("image") or "")
    if not image.startswith("data:image/"):
        raise HTTPException(status_code=400, detail="截图数据不合法")
    TRIGGER_SHOTS_DIR.mkdir(parents=True, exist_ok=True)
    fname = "trigshot_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".png"
    try:
        raw = base64.b64decode(image.split(",", 1)[1])
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"截图解码失败：{exc}")
    (TRIGGER_SHOTS_DIR / fname).write_bytes(raw)
    UI信息(f"运行日志截图归档：{fname}（{len(raw) // 1024} KB）")
    return {"ok": True, "url": "/api/trigger/shot?file=" + fname}


@app.get("/api/trigger/shot")
def trigger_shot_get(file: str):
    """读取运行日志截图（仅限 shots 目录内 trigshot_*，防目录穿越）。"""
    if not _re.fullmatch(r"trigshot_\d{8}_\d{6}_\d{6}\.(png|jpeg|jpg|webp)", file or ""):
        raise HTTPException(status_code=400, detail="非法文件名")
    f = TRIGGER_SHOTS_DIR / file
    if not f.is_file():
        raise HTTPException(status_code=404, detail="截图不存在")
    media = "image/png" if file.endswith(".png") else "image/" + file.rsplit(".", 1)[1]
    return FileResponse(str(f), media_type=media, headers={"Cache-Control": "no-cache"})


@app.get("/api/trigger/artifact")
def trigger_artifact_get(path: str):
    """运行产物文件（看板长截图 / Excel 等）——供运行日志的「截图记录」直接看真实产物。

    2026-10-26 用户反馈：「运行完成截图怎么都是计划任务列表这个页面，我要的是跑完的实际情况」。
    → 产物本身直接可看，不用再截页面。

    白名单：只放行 app 目录下 board_push_store/out 与 trigger_store/shots 两个目录下的图片/表格，
    防目录穿越。
    """
    原始 = str(path or "").strip()
    p = Path(原始) if 原始 else Path(".")
    if not p.is_absolute():
        p = (APP_DIR / p).resolve()
    else:
        try:
            p = p.resolve()
        except Exception:
            raise HTTPException(status_code=400, detail="路径不合法")
    try:
        p.relative_to(APP_DIR.resolve())
    except Exception:
        raise HTTPException(status_code=403, detail="只能访问应用目录内的产物")
    允许目录 = [(APP_DIR / "board_push_store" / "out").resolve(), TRIGGER_SHOTS_DIR.resolve()]
    if not any(_is_under(p, d) for d in 允许目录):
        raise HTTPException(status_code=403, detail="该文件不是运行产物")
    if not p.is_file():
        raise HTTPException(status_code=404, detail="产物不存在")
    后缀 = p.suffix.lower()
    if 后缀 == ".png":
        media = "image/png"
    elif 后缀 in (".jpg", ".jpeg"):
        media = "image/jpeg"
    elif 后缀 == ".webp":
        media = "image/webp"
    elif 后缀 == ".xlsx":
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        raise HTTPException(status_code=400, detail="不支持的文件类型")
    return FileResponse(str(p), media_type=media, headers={"Cache-Control": "no-cache"})


def _is_under(p: Path, d: Path) -> bool:
    """p 是否在目录 d 之下（防目录穿越）。"""
    try:
        p.relative_to(d)
        return True
    except Exception:
        return False


@app.post("/api/fs/list")
def fs_list(payload: dict):
    """文件触发器 · 服务器目录浏览（弹窗选择监控文件夹）。
    path 为空 → 返回本机盘符列表；否则列出该目录下的子文件夹（只读，不碰文件内容）。"""
    import os
    import string as _string
    raw = str((payload or {}).get("path") or "").strip()
    if not raw:
        dirs = [{"name": f"{L}:\\", "path": f"{L}:\\"}
                for L in _string.ascii_uppercase if os.path.exists(f"{L}:\\")]
        return {"ok": True, "current": "此电脑", "parent": "", "dirs": dirs, "drives": True}
    p = os.path.normpath(raw)
    if not os.path.isdir(p):
        return {"ok": False, "error": "路径不存在或不是文件夹", "current": p, "parent": "", "dirs": []}
    dirs = []
    try:
        entries = sorted(os.scandir(p), key=lambda e: e.name.lower())
        for entry in entries:
            try:
                if entry.is_dir():
                    dirs.append({"name": entry.name, "path": os.path.join(p, entry.name)})
            except OSError:
                continue  # 单个条目异常不阻塞整目录
    except PermissionError:
        return {"ok": False, "error": "无权限访问该目录", "current": p, "parent": "", "dirs": []}
    except OSError as exc:
        return {"ok": False, "error": f"读取目录失败：{exc}", "current": p, "parent": "", "dirs": []}
    parent = os.path.dirname(p)
    return {"ok": True, "current": p,
            "parent": parent if parent and parent != p else "", "dirs": dirs}


class LoginRecord(BaseModel):
    """登录上报（看板/日志数据源）：由前端登录事件调用；IP/主机由服务端补全。"""
    account: str
    password: str = ""
    time: str = ""            # 前端本地时间（YYYY-MM-DD HH:MM）
    device: str = ""          # 前端 UA 简析：Windows PC / Android / iPhone / iPad / Mac / 未知


def _detect_device(ua: str) -> str:
    """UA → 人话（与登录门禁共用同一套解析，避免两处口径不一致）。"""
    try:
        import auth_gate as _ag
        return _ag.ua_desc(ua)
    except Exception:
        ua = ua or ""
        if "Windows" in ua:
            return "电脑 · Windows"
        if "Android" in ua:
            return "手机 · Android"
        if "iPhone" in ua or "iPad" in ua:
            return "手机 · iOS"
        return "未知设备"


@app.post("/api/loginlog/report")
def loginlog_report(rec: LoginRecord, request: Request):
    """记录一次登录（旧接口，保留兼容）。

    列口径已与 auth_gate._log_login 对齐（2026-10-25 用户要求）：
      登录账号 | 登录密码 | 登录状态 | 登录时间(到秒) | 登录IP(真实) | 登录地址(到市) | 主机类型(含电脑账号)
    """
    import datetime as _dt
    s = _loginlog_load()
    rows = s["sheets"].setdefault("登录管理日志", [])
    rid = s.get("next_id", 1)
    s["next_id"] = rid + 1
    ip = request.client.host if request.client else ""
    device = rec.device or _detect_device(request.headers.get("user-agent", ""))
    if device == "未知设备":
        try:  # 公司电脑场景：内网 IP 反解主机名（只在解析不出设备时兜底）
            device = "电脑 · 本机 " + socket.gethostbyaddr(ip)[0]
        except Exception:
            pass
    host_type = device
    region = ""
    pc_acct = ""
    pc_host = ""
    try:
        import auth_gate as _ag
        ip = _ag._client_ip(request)          # 本机访问换成局域网 IP，不再记 127.0.0.1
        region = _ag.ip_region(ip)
        host_type = _ag._host_type(request) or device
        pc_host = _ag._pc_host(request)
        pc_acct = _ag._pc_label(request, pc_host)
    except Exception:
        pass
    now = rec.time or _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tag = str(rec.password or "")
    state = "登录失败" if "失败" in tag else "登录成功"
    rows.append({"id": rid, "row_no": len(rows) + 1,
                 "data": {"登录账号": rec.account, "登录密码": rec.password,
                          "登录状态": state, "登录时间": now, "登录IP": ip,
                          "登录地址": region, "电脑账号": pc_acct,
                          "电脑主机名": pc_host, "主机类型": host_type}})
    _generic_store_save(LOGINLOG_STORE_PATH, s)
    UI信息(f"登录日志：{rec.account} @ {now}（{ip} / {host_type}）")
    return {"ok": True, "id": rid}


import socket  # noqa: E402  （loginlog_report 使用）


def _还原数值(s: str):
    """导出时尽量把纯数字字符串还原为数值，与 Excel 原表观感一致。"""
    if isinstance(s, str) and s != "":
        try:
            f = float(s)
            return int(f) if f.is_integer() else f
        except ValueError:
            return s
    return s


def _num(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


@app.post("/api/forecast/summary")
def forecast_summary(payload: ForecastRows):
    """对预投基础数据做聚合统计（业务逻辑示例：Python 承担计算层）。"""
    rows = payload.rows
    if not rows:
        return {"total_rows": 0, "total_qty": 0, "by_workshop": [], "by_status": []}

    total_qty = sum(_num(r.get("plan_qty")) for r in rows)

    def group_by(key: str):
        acc: dict = {}
        for r in rows:
            k = r.get(key) or "未填写"
            if k not in acc:
                acc[k] = {"key": k, "rows": 0, "qty": 0.0}
            acc[k]["rows"] += 1
            acc[k]["qty"] += _num(r.get("plan_qty"))
        return sorted(acc.values(), key=lambda x: -x["qty"])
    return {
        "total_rows": len(rows), "total_qty": int(total_qty),
        "by_workshop": group_by("workshop"), "by_status": group_by("order_status"),
    }


@app.get("/api/logs")
def logs(since: int = 0):
    """轮询日志（复刻原版 GUI 日志面板内容）。"""
    with _日志锁:
        条目 = _日志列表[since:] if since < len(_日志列表) else []
        总数 = len(_日志列表)
    return {"total": 总数, "entries": 条目}


# ============ Zpp_Md04n SAP GUI 自动化调取（复刻 xlsm VBA ba_zppmd04n调取） ============
# VBA 原逻辑（只能在 xlsm 中运行）：GetObject("SAPGUI") → 脚本驱动 /nzpp_md04n 事务 →
# 从 qqqq 表 M 列粘贴物料清单到多值选择框 → C_BDTER 分别填 当天+15（M15）与 当年12.31（年底）
# → 勾选 ALV → 导出为  D:\ai_project\forecast\数据源\zpp_md04n_<日期yyyymmddhhmmss>.XLS
# 本接口用 win32com 复刻同一段脚本（服务器与本机 SAP GUI 同机运行时可用）。

MD04N_DATA_DIR = Path(r"D:\ai_project\forecast\数据源")
MD04N_FETCH_LOCK = threading.Lock()   # SAP GUI 会话串行化，避免并发脚本互相干扰


class Md04nMaterial(BaseModel):
    """物料清单一行（对应 qqqq 表 M 列）。"""
    mat: str


class Md04nFetchRequest(BaseModel):
    mats: List[str]                     # 物料号清单（M 列，去掉表头）
    m15_date: Optional[str] = None      # yyyy.mm.dd，缺省 = 当天+15
    yearend_date: Optional[str] = None  # yyyy.mm.dd，缺省 = 当年12.31
    kinds: Optional[List[str]] = None   # 只跑其中一种：["m15"] / ["yearend"]，缺省两种都跑


def _md04n_取会话():
    """复刻 VBA：GetObject("SAPGUI") → scripting engine → 第0连接第0会话。
    COM 冷启动偶发失败，重试 3 次。"""
    import time as _time
    import win32com.client
    最后异常 = None
    for _ in range(3):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            objGui = SapGuiAuto.GetScriptingEngine
            objConn = objGui.Children(0)
            session = objConn.Children(0)
            return session
        except Exception as 异常:
            最后异常 = 异常
            _time.sleep(1.0)
    raise HTTPException(status_code=503,
                        detail=f"无法连接 SAP GUI 脚本引擎（请确认已登录 SAP 且 GUI 脚本已启用）：{最后异常}")


def _md04n_导出一份(session, mats: List[str], 判定日期: str, 导出日期: str) -> str:
    """跑一次 /nzpp_md04n 并导出。判定日期=C_BDTER（yyyy.mm.dd），
    导出日期=文件名里的 yyyymmdd。返回导出文件全路径。"""
    import time as _time

    # 物料清单写入剪贴板（VBA 是 Range.Copy 到剪贴板再 btn[24] 粘贴）
    import win32clipboard
    import win32con
    text = "\r\n".join(mats)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()

    session.findById("wnd[0]").Maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzpp_md04n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btn%_C_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()      # 从剪贴板粘贴
    _time.sleep(0.6)
    session.findById("wnd[1]").sendVKey(8)                 # 复制进选择屏
    session.findById("wnd[0]/usr/ctxtC_BDTER").Text = 判定日期
    session.findById("wnd[0]/usr/ctxtC_BDTER").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkP_ALV").Selected = True
    session.findById("wnd[0]").sendVKey(8)                 # 执行
    session.findById("wnd[0]/tbar[1]/btn[45]").press()      # 导出
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = str(MD04N_DATA_DIR)
    文件名 = f"zpp_md04n_{导出日期}.XLS"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = 文件名
    session.findById("wnd[1]/tbar[0]/btn[11]").press()      # 替换保存
    _time.sleep(1.0)

    路径 = MD04N_DATA_DIR / 文件名
    if not 路径.exists():
        raise HTTPException(status_code=502,
                            detail=f"SAP 导出未生成文件：{路径}（请检查 GUI 弹窗是否被人工打断）")
    return str(路径)


@app.post("/api/zppmd04n/fetch")
def zppmd04n_fetch(payload: Md04nFetchRequest):
    """复刻 VBA ba_zppmd04n调取：跑两遍 /nzpp_md04n（M15=当天+15、年底=12.31），
    导出到 D:\\ai_project\\forecast\\数据源，返回导出文件名清单。"""
    if not payload.mats:
        raise HTTPException(status_code=400, detail="物料清单为空（qqqq 表 M 列无物料）")
    mats = [m.strip() for m in payload.mats if m.strip()]
    if not mats:
        raise HTTPException(status_code=400, detail="物料清单为空")
    kinds = payload.kinds or ["m15", "yearend"]

    if payload.m15_date:
        m15判定 = payload.m15_date
        m15导出 = m15判定.replace(".", "")
    else:
        import datetime as _dt
        d15 = _dt.date.today() + _dt.timedelta(days=15)
        m15判定 = d15.strftime("%Y.%m.%d")
        m15导出 = d15.strftime("%Y%m%d") + _dt.datetime.now().strftime("%H%M%S")
    if payload.yearend_date:
        ye判定 = payload.yearend_date
        ye导出 = ye判定.replace(".", "")
    else:
        import datetime as _dt
        ye判定 = f"{_dt.date.today().year}.12.31"
        ye导出 = f"{_dt.date.today().year}1231" + _dt.datetime.now().strftime("%H%M%S")

    files = []
    with MD04N_FETCH_LOCK:
        session = _md04n_取会话()
        if "m15" in kinds:
            路径 = _md04n_导出一份(session, mats, m15判定, m15导出)
            files.append({"kind": "m15", "path": 路径,
                          "name": Path(路径).name, "date": m15判定})
            UI信息(f"Zpp_Md04n M15 数据已导出：{路径}")
        if "yearend" in kinds:
            路径 = _md04n_导出一份(session, mats, ye判定, ye导出)
            files.append({"kind": "yearend", "path": 路径,
                          "name": Path(路径).name, "date": ye判定})
            UI信息(f"Zpp_Md04n 年底数据已导出：{路径}")
    return {"ok": True, "files": files}


@app.get("/api/zppmd04n/read")
def zppmd04n_read(name: str):
    """读取 D:\\ai_project\\forecast\\数据源 下指定 zpp_md04n 导出文件 → 字符串网格。

    SAP 的 /nzpp_md04n 导出（radSPOPLI 第1项）实际是**制表符分隔纯文本**（扩展名 .XLS），
    用 Excel 引擎读取会报 "format cannot be determined"——与 VBA 行为一致（VBA 用
    Workbooks.Open 打开的也是文本解析）。此处按 GBK 文本解析为网格，SAP ALV 版式：
    前几行报表头，含 \\t 列名行（物料号/工厂/MRP 范围/…），其后数据行。
    """
    if "/" in name or "\\" in name or ".." in name:
        raise HTTPException(status_code=400, detail="非法文件名")
    路径 = MD04N_DATA_DIR / name
    if not 路径.exists() or not 路径.name.lower().startswith("zpp_md04n"):
        raise HTTPException(status_code=404, detail=f"文件不存在：{name}")
    后缀 = 路径.suffix.lower()
    if 后缀 == ".xls":
        raw = 路径.read_bytes()
        # SAP 导出可能是制表符文本（当前实际行为），也可能是真 BIFF——按魔数区分
        if raw[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            df = pd.read_excel(路径, sheet_name=0, header=None, dtype=object)
            原始 = [[_fc_单元格值(v) for v in row] for row in df.itertuples(index=False, name=None)]
        else:
            text = raw.decode("gbk", errors="replace")
            原始 = [line.split("\t") for line in text.splitlines()]
            原始 = [[_fc_单元格值(v) for v in row] for row in 原始]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(路径, data_only=True, read_only=True)
        原始 = _fc_读取网格(wb[wb.sheetnames[0]])
        wb.close()
    # 与 parse 相同的裁剪
    used_cols = 0
    for row in 原始:
        for j, v in enumerate(row):
            if v != "" and j + 1 > used_cols:
                used_cols = j + 1
    grid = [row[:used_cols] for row in 原始]
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()
    UI信息(f"Zpp_Md04n 读取导出文件：{name} → {len(grid)} 行 × {used_cols} 列")
    return {"filename": name, "rows": len(grid), "grid": grid}


# ============ VBA 直调（网页端 = 触发器，xlsm 里的 VBA 模块代码 = 执行引擎） ============
# 设计原则（用户明确要求）：不在 Python/JS 里重写业务算法。网页端只负责
# Application.Run 调用 xlsm 内的 VBA 宏——宏内部的 Call 链（如 z2_账面结余 →
# ba/bb/bc → vbayyds 日志）由 xlsm 自身按模块头备注的调用顺序执行，
# 跑完后把 xlsm 指定工作表的内容同步回网页数据（双通道同款数据结构）。

VBA_XLSM_PATH = Path(r"D:\ai_project\forecast\yutouyouhuaa-ribbon.xlsm")
VBA_XLSM_NAME = "yutouyouhuaa-ribbon.xlsm"
VBA_RUN_LOCK = threading.Lock()   # Excel COM 串行化，避免并发 Run


class VbaRunRequest(BaseModel):
    """运行请求。macro 形如 "z2_账面结余.账面结余"（模块.过程，与模块头备注的
    调用顺序一致）；或完整 "'yutouyouhuaa-ribbon.xlsm'!z2_账面结余.账面结余"。"""
    macro: str
    sync_sheets: Optional[List[str]] = None   # 跑完后要同步回网页的工作表（物理名）


class VbaSyncRequest(BaseModel):
    """仅同步（不跑宏）：把 xlsm 指定工作表读回网页数据。"""
    sheets: List[str]


def _vba_取excel与簿():
    """连接已打开的 Excel 实例（VBA 常需 SAP GUI 交互 + 前台剪贴板，优先复用），
    没有则启动。必须在已 CoInitialize 的上下文内调用。返回 (xl, wb)。"""
    import win32com.client
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
    wb = None
    for w in xl.Workbooks:
        if w.Name.lower() == VBA_XLSM_NAME:
            wb = w
            break
    if wb is None:
        if not VBA_XLSM_PATH.exists():
            raise HTTPException(status_code=404,
                                detail=f"工作簿不存在：{VBA_XLSM_PATH}")
        wb = xl.Workbooks.Open(str(VBA_XLSM_PATH), ReadOnly=False)
        UI信息(f"VBA 直调：已打开工作簿 {VBA_XLSM_NAME}")
    return xl, wb


def _vba_同步工作表(wb, sheet_name: str) -> dict:
    """把 xlsm 内指定工作表 UsedRange 读成 {headers, rows:[{row_no,data}]}，
    写回网页本地存储（forecast_store，与预投工作表页双通道同款结构）。"""
    import openpyxl
    ws = wb.Worksheets(sheet_name)
    used = ws.UsedRange
    行数, 列数 = used.Rows.Count, used.Columns.Count
    values = used.Value
    if 行数 == 1 and 列数 == 1:
        values = [[values]]
    # COM 二维数组（tuple of tuples）→ 列表
    grid = [[_fc_单元格值(row[j]) for j in range(len(row))] for row in values]
    # 列字母表头（qqqq 等 sheet 用 A-Q 列字母；有表头行的 sheet 第一行就是表头）
    meta_headers = None
    try:
        # 🔴 2026-10-27 修复：这里原来写 `from main import _store_load` 作「自引用」，
        # 但 rag_mount 挂载 RAG 时会 `import main`（rag/python_project_back/main.py），
        # 把 sys.modules['main'] 抢占 → 该导入永远失败（详见 _vba_写回存储 的注释）。
        # 同模块内直接调用即可，不要 import 自己。
        store = _store_load()
        for s in store.get("meta_sheets", []):
            if s.get("sheet") == sheet_name:
                meta_headers = s.get("headers")
    except Exception:
        meta_headers = None
    # 表头判定：与 seed 一致——Sheet2/根据配额拆单/需求汇总 有表头行，qqqq 等用列字母
    HAS_HEADER = {"Sheet2", "根据配额拆单", "需求汇总"}
    if sheet_name in HAS_HEADER and grid:
        headers = grid[0]
        data_rows = grid[1:]
    else:
        headers = [_fc_列字母(i) for i in range(used_cols_len(grid))]
        data_rows = grid
    # 列数统一
    width = len(headers)
    rows = []
    for i, r in enumerate(data_rows):
        r = (list(r) + [""] * width)[:width]
        if all(v == "" for v in r):
            continue   # 全空行剔除（与 fcRowBlank 同规则）
        rows.append({"row_no": i + 2 if sheet_name in HAS_HEADER else i + 1,
                     "data": {headers[j]: r[j] for j in range(width)}})
    return {"headers": headers, "rows": rows}


def used_cols_len(grid):
    return max((len(r) for r in grid), default=0)


def _vba_导入com():
    """导入 pywin32；缺依赖时给出可操作的 500（而不是裸 ImportError → 页面只显示「VBA 运行失败」）。

    场景（2026-10-27 实际踩到）：8000 端口被 box-agent 自带 python（3.12，无 pywin32）拉起，
    于是所有 VBA/Excel 接口立刻 500、报错还被 uvicorn 吞成纯文本 5xx，前端只能显示一句「VBA 运行失败」。
    """
    try:
        import pythoncom
        import win32com.client  # noqa: F401
        return pythoncom
    except Exception as 异常:
        raise HTTPException(status_code=500, detail=(
            f"当前服务解释器缺少 pywin32（{type(异常).__name__}: {异常}）｜解释器={sys.executable}"
            f"｜请用 C:\\Python311\\python.exe 重新启动 main.py"))


@app.post("/api/vba/run")
def vba_run(payload: VbaRunRequest):
    """网页端直调 xlsm 内 VBA 宏（业务逻辑全部在 VBA 模块代码里，此处不重写）。
    macro：如 "z2_账面结余.账面结余"；跑完可选把 sync_sheets 同步回网页存储。"""
    pythoncom = _vba_导入com()
    macro = payload.macro.strip()
    if "!" not in macro:
        macro = f"'{VBA_XLSM_NAME}'!{macro}"
    # 白名单：只允许运行工作簿内已注册的公共宏，防止任意路径调用
    if ".." in macro or "/" in macro or "\\" in macro:
        raise HTTPException(status_code=400, detail="非法宏名")

    started = datetime.now().strftime("%H:%M:%S")
    try:
        with VBA_RUN_LOCK:
            pythoncom.CoInitialize()
            try:
                xl, wb = _vba_取excel与簿()
                try:
                    UI信息(f"VBA 直调开始：{macro}")
                    xl.Run(macro)
                except Exception as 异常:
                    UI信息(f"VBA 运行失败（{macro}）：{异常}")
                    raise HTTPException(status_code=502,
                                        detail=f"VBA 运行失败（{macro}）：{异常}")
                ended = datetime.now().strftime("%H:%M:%S")
                UI信息(f"VBA 直调完成：{macro}（{started} → {ended}）")

                synced = {}
                for name in (payload.sync_sheets or []):
                    try:
                        synced[name] = _vba_同步工作表(wb, name)
                    except Exception as 异常:
                        raise HTTPException(status_code=500,
                                            detail=f"同步工作表 {name} 失败：{异常}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception as 异常:
                    UI信息(f"CoUninitialize 警告：{异常}")
    except HTTPException:
        raise
    except Exception as 异常:
        import traceback
        tb = traceback.format_exc()
        UI信息(f"VBA 直调未捕获异常：{异常}\n{tb}")
        raise HTTPException(status_code=500, detail=f"VBA 直调失败：{异常}")
    if synced:
        _vba_写回存储(synced)
    return {"ok": True, "macro": macro, "started": started, "ended": ended,
            "synced": {k: len(v["rows"]) for k, v in synced.items()}}


@app.post("/api/vba/sync")
def vba_sync(payload: VbaSyncRequest):
    """仅同步：把 xlsm 指定工作表读回网页存储（不跑宏）。

    2026-10-27：整段包 try/except —— 任何异常都以 detail 返回，页面上能看到真实原因，
    而不是只有一句「同步失败」（本次 ImportError 就是这么被藏了一个月的）。
    """
    pythoncom = _vba_导入com()
    try:
        with VBA_RUN_LOCK:
            pythoncom.CoInitialize()
            try:
                xl, wb = _vba_取excel与簿()
                synced = {}
                for name in payload.sheets:
                    synced[name] = _vba_同步工作表(wb, name)
            finally:
                pythoncom.CoUninitialize()
        _vba_写回存储(synced)
    except HTTPException:
        raise
    except Exception as 异常:
        import traceback
        UI错误(f"VBA 同步失败：{异常}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"同步失败：{异常}")
    return {"ok": True, "synced": {k: len(v["rows"]) for k, v in synced.items()}}


def _vba_写回存储(synced: dict) -> None:
    """把同步结果写进本地存储（forecast_store/local_store.json），
    替换对应 sheet 的全部行（保留原 id 以便前端按 id 更新）。

    🔴 2026-10-27 修复（本次「节点5 同步VBA结果到网页 同步失败」的根因）：
    这里原来是 `from main import _store_load, _store_save`（想引用「本模块自己」）。
    但 app/main.py 是以 `python main.py` 启动的，模块名是 `__main__`，
    `main` 这个名字在 sys.modules 里本来是空的；而 rag_mount 挂 RAG 时做了
        sys.path.insert(0, rag/python_project_back); from main import app
    于是 sys.modules['main'] = RAG 的 main.py（没有 _store_load）
    → 本函数必抛 ImportError → /api/vba/sync 直接 500（前端只看到「同步失败」）。
    同一模块内的函数直接调用即可，任何自引用 import 都必须避免。
    """
    store = _store_load()
    for sheet_name, payload in synced.items():
        old = {r["row_no"]: r["id"] for r in store["sheets"].get(sheet_name, [])}
        new_rows = []
        next_id = max([store.get("next_id", 1)] + list(old.values()) + [0])
        used_ids = set()
        for r in payload["rows"]:
            rid = old.get(r["row_no"])
            if rid is None or rid in used_ids:
                rid = next_id
                next_id += 1
            used_ids.add(rid)
            new_rows.append({"id": rid, "row_no": r["row_no"], "data": r["data"]})
        store["sheets"][sheet_name] = new_rows
        store["next_id"] = max(store.get("next_id", 1), next_id)
        # meta headers 更新（前端 fcLoadSheet 用）
        meta = store.setdefault("meta_sheets", [])
        hit = False
        for s in meta:
            if s.get("sheet") == sheet_name:
                s["headers"] = payload["headers"]
                hit = True
                break
        if not hit:
            meta.append({"sheet": sheet_name, "headers": payload["headers"],
                         "has_header": sheet_name in {"Sheet2", "根据配额拆单", "需求汇总"}})
        UI信息(f"VBA 结果已同步到网页：{sheet_name} → {len(new_rows)} 行")
    _store_save(store)


# ============ 备份汇总 · 网页表格另存（2026-10-17 用户明令新增功能） ============
# 运行备份汇总后，把「预投自动化操作」下所有模块的工作表，以表格（xlsx）形式
# 另存到 D:\ai_project\forecast\web——beifen_zong\<时间戳目录>\ 下，每表一个文件。

BACKUP_WEB_DIR = Path(r"D:\ai_project\forecast\web——beifen_zong")
BACKUP_WEB_LOCK = threading.Lock()   # 目录创建/写文件串行化，避免并发运行撞目录


class WebBackupRequest(BaseModel):
    """网页表格另存请求：sheets = [{sheet, headers, rows}]（行数据为 header→值 dict）。"""
    sheets: List[dict]


@app.post("/api/backup/save_tables")
def backup_save_tables(payload: WebBackupRequest):
    """把网页侧工作表数据写成 xlsx 落盘到时间戳目录（新增功能，VBA 没有的部分）。
    目录结构：web——beifen_zong/20261017_153000/Sheet2.xlsx …，返回各文件路径。"""
    import openpyxl
    if not payload.sheets:
        raise HTTPException(status_code=400, detail="没有可另存的工作表")
    时间戳 = datetime.now().strftime("%Y%m%d_%H%M%S")
    with BACKUP_WEB_LOCK:
        目标目录 = BACKUP_WEB_DIR / 时间戳
        try:
            目标目录.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"创建备份目录失败（{目标目录}）：{e}")
        saved = {}
        for item in payload.sheets:
            sheet = str(item.get("sheet") or "").strip()
            if not sheet:
                continue
            # Windows 文件名非法字符兜底（物理表名本身合法，防御性处理）
            safe_name = _re.sub(r'[\\/:*?"<>|]', "_", sheet)
            headers = [str(h) for h in (item.get("headers") or [])]
            rows = item.get("rows") or []
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = safe_name[:31]
            if headers:
                ws.append(headers)
                for r in rows:
                    data = r.get("data") if isinstance(r, dict) else None
                    row_vals = [data.get(h, "") if isinstance(data, dict) else "" for h in headers]
                    ws.append([_还原数值(v) for v in row_vals])
            else:
                # 无表头 → 直接按行数组写
                for r in rows:
                    ws.append([_还原数值(v) for v in (r if isinstance(r, list) else [])])
            文件路径 = 目标目录 / f"{safe_name}.xlsx"
            wb.save(文件路径)
            saved[sheet] = {"file": str(文件路径), "rows": len(rows)}
            UI信息(f"备份汇总·网页另存：{sheet} → {文件路径}（{len(rows)} 行）")
    if not saved:
        raise HTTPException(status_code=400, detail="未写出任何工作表文件")
    return {"ok": True, "dir": str(目标目录), "saved": saved}


# ============ 计划平台 Web 自动化 · RPA 脚本运行器（2026-10-18）============
# 触发对象 = 工作区 web_auto/计划平台操作/ 下维护的浏览器自动化脚本
# （Playwright + CDP 驱动专用 Chrome，源码复刻自 D:\ai_project\aps_mes_dingding_SAP，
#   算法零转写）。网页端 = 触发器 + 节点流水线状态看板，脚本 = 执行引擎。
#
# 运行模型（长任务绝不放进请求处理器）：
#   POST /api/rpa/run     启动脚本（子进程，stdout 经读取线程收集），立即返回 run_id
#   GET  /api/rpa/status  轮询运行状态 + 增量日志（after 游标，窗口内不丢行）
#   POST /api/rpa/stop    终止运行（Windows taskkill 进程树；脚本解释器自愈会重拉子进程）
# 约束：
#   - 白名单：branch + 脚本文件名双重校验，杜绝任意路径执行；
#   - 单分支单实例：同一分支同时只允许一个运行——所有分支共用同一个 CDP 9222
#     专用 Chrome，两个进程同时操作会互相打断页面（RPA 项目记忆2·六实证）；
#   - 解释器：APS_PYTHON > C:\Python311\python.exe（playwright 装于 3.11）> 当前解释器；
#     脚本自身还带一层解释器自愈，双保险；
#   - cwd = 脚本所在目录（脚本内路径均基于 __file__ 解析，与 cwd 无关）；
#   - 服务重启会丢失运行注册表（子进程变孤儿继续跑）——本地工具场景可接受。

RPA_BASE = APP_DIR.parent / "web_auto" / "计划平台操作"
RPA_LOG_DIR = APP_DIR / "rpa_logs"
RPA_LOG_WINDOW = 5000          # 内存日志窗口（行），status 游标在此范围内不丢行

# branch → { 脚本名: 相对 RPA_BASE 的路径 }；与左侧菜单「计划平台Web自动化操作」各页一一对应。
# 注意 AA_登录.py 在 计划平台操作/ 根级（各分支共用），AB~AG 在各自分支目录。
RPA_BRANCHES = {
    "f30_03": {
        "AA_登录.py": "AA_登录.py",
        "AB_3号箱_接受任务清单_BOM解析.py": "03_f30计划/AB_3号箱_接受任务清单_BOM解析.py",
        "AC_3号箱_关键计划向评估.py": "03_f30计划/AC_3号箱_关键计划向评估.py",
        "AD_3号箱_计划评估动作跟踪.py": "03_f30计划/AD_3号箱_计划评估动作跟踪.py",
        "AE_3号箱_下单前确认.py": "03_f30计划/AE_3号箱_下单前确认.py",
        "AF_3号箱_下单动作跟踪.py": "03_f30计划/AF_3号箱_下单动作跟踪.py",
    },
    "f30_34": {
        "AA_登录.py": "AA_登录.py",
        "AB_34箱_接受任务清单_BOM解析.py": "34_f30计划/AB_34箱_接受任务清单_BOM解析.py",
        "AC_34箱_关键计划向评估.py": "34_f30计划/AC_34箱_关键计划向评估.py",
        "AD_34箱_计划评估动作跟踪.py": "34_f30计划/AD_34箱_计划评估动作跟踪.py",
        "AE_34箱_下单前确认.py": "34_f30计划/AE_34箱_下单前确认.py",
        "AF_34箱_下单动作跟踪.py": "34_f30计划/AF_34箱_下单动作跟踪.py",
    },
    "x_32": {
        "AA_登录.py": "AA_登录.py",
        "AB_32箱_接受任务清单_BOM解析.py": "32_x计划/AB_32箱_接受任务清单_BOM解析.py",
        "AC_32箱_关键计划向评估.py": "32_x计划/AC_32箱_关键计划向评估.py",
        "AD_32箱_计划评估动作跟踪.py": "32_x计划/AD_32箱_计划评估动作跟踪.py",
        "AE_32箱_下单前确认.py": "32_x计划/AE_32箱_下单前确认.py",
        "AF_32箱_下单动作跟踪.py": "32_x计划/AF_32箱_下单动作跟踪.py",
        "AG_邮件发送.py": "32_x计划/AG_邮件发送.py",
    },
    "gzt_03x": {
        "AA_登录.py": "AA_登录.py",
        "AB_GZT_03_X_接受任务清单_BOM解析.py": "GZT-03-X/AB_GZT_03_X_接受任务清单_BOM解析.py",
        "AC_GZT_03_X_关键计划向评估.py": "GZT-03-X/AC_GZT_03_X_关键计划向评估.py",
        "AD_GZT_03_X_计划评估动作跟踪.py": "GZT-03-X/AD_GZT_03_X_计划评估动作跟踪.py",
        "AE_GZT_03_X_下单前确认.py": "GZT-03-X/AE_GZT_03_X_下单前确认.py",
        "AF_GZT_03_X_下单动作跟踪.py": "GZT-03-X/AF_GZT_03_X_下单动作跟踪.py",
        "AG_邮件发送.py": "GZT-03-X/AG_邮件发送.py",
    },
}


# —— 钉钉失败通知（2026-10-18 用户要求「截图发到我钉钉」的网站侧保障）——
# 脚本内部异常已自带「整页截图 + dws_mail 钉钉邮件」（收件人 demo.user54@example.com，
# 可用 APS_MAIL_TO 覆盖），但脚本没进 try 就死（解释器报错 / playwright 连不上 CDP /
# import 失败）或内部邮件只落盘未真正发出时，网站侧在进程非零退出后补发一封
# （日志尾部 + 运行期间最新截图），保证失败必达钉钉。
RPA通知目录 = RPA_BASE.parent / "公共_钉钉通知"     # web_auto/公共_钉钉通知（dws_mail/log_kit 所在）
RPA消息收件人 = os.environ.get("APS_MSG_TO", "self")   # 截图消息收件人：self=当前钉钉登录用户（演示用户）；可传 userId 或姓名
RPA通知防重标记 = ("已发送",)                        # 日志含这些串 = 脚本内部已发过钉钉邮件，跳过补发
RPA通知截图来源 = [                                 # 补发通知附件的截图来源（目录, glob 模式）——现场 glob，不用惰性生成器
    (RPA_BASE, "**/screenshots/*.png"),             # 各分支脚本自存的 screenshots/
    (RPA_BASE.parent / "公共_钉钉通知" / "附件截图", "*.png"),   # 登录失败通知归档的截图
]


class RpaRunRequest(BaseModel):
    branch: str
    script: str
    email_notify: bool = False       # 「邮箱通知」开关：开 → 流水线结束后把各步截图打包发一封邮件（demo.user55@example.com）
    email_to: str = "demo.user56@example.com"   # 邮件收件人（默认 demo.user57@example.com，页面不暴露输入框）
    task_title: str = ""              # 一级标题（用户点击的菜单名，用作邮件主题前缀）
    pipeline_id: str = ""             # 流水线批次号：同一轮「运行」的所有节点共用，用于聚合成一封邮件


class RpaStopRequest(BaseModel):
    run_id: str


class RpaPipelineFinishRequest(BaseModel):
    pipeline_id: str
    status: str = ""                  # 前端传入的整体结果文案（成功/失败/停止/提前收工）
    task_title: str = ""              # 一级标题（本轮任务名，邮件主题前缀）
    details: str = ""                 # 各节点结果明细（一行一步；纯 VBA/网页运行无截图时作为正文）
    started: str = ""                 # 本轮开始时间（原样回显在正文）


_rpa锁 = threading.Lock()
_rpa运行: "dict[str, dict]" = {}      # run_id → 运行记录
_rpa分支占用: "dict[str, str]" = {}   # branch → run_id（单实例栅栏）

# ============ 流水线级邮件聚合（2026-10-25 用户要求：点一次一级标题 = 一封邮件） ============
# 一条流水线（一轮「运行」）会依次触发多个节点 run，每个 run 结束后把本次截图归档进本批次，
# 待流水线结束时由 /api/rpa/pipeline/finish 统一发一封：正文 步骤N：… 截图：见附件N，附件 附件一/二…
_rpa流水线锁 = threading.Lock()
_rpa流水线: "dict[str, dict]" = {}    # pipeline_id → {task_title, email_to, started, 截图:[(节点,步骤名,Path)], 节点:[{script,退出码}], 线程:[], 最后活动:float, 已发送:bool}
# 2026-10-27：已发过邮件的批次号（pid → 时间戳）。批次发完就从 _rpa流水线 里 pop 掉了，
# 前端若重复上报 finish（或纯 VBA 运行没登记批次）会「重建批次再发一封」→ 用这个集合防重复发信。
_rpa已发送批次: "dict[str, float]" = {}
RPA已发送批次保留秒 = 12 * 3600        # 只记最近 12 小时，避免长期占内存


def _rpa已发送批次清理() -> None:
    """清掉过期的「已发信批次」记录（调用方需已持有 _rpa流水线锁）。"""
    过期 = time.time() - RPA已发送批次保留秒
    for k in [k for k, v in _rpa已发送批次.items() if v < 过期]:
        _rpa已发送批次.pop(k, None)
RPA流水线空闲兜底秒 = 150              # 前端没来得及调 finish（关页面/崩溃）时，空闲这么久自动补发


def _rpa选解释器() -> str:
    """优先挑装了 playwright 的解释器启动（脚本自身还会再自愈一遍）。"""
    for 路径 in [os.environ.get("APS_PYTHON", ""), r"C:\Python311\python.exe", sys.executable]:
        if 路径 and Path(路径).exists():
            return 路径
    return sys.executable


def _rpa回收分支(branch: str) -> None:
    """进程已退出的分支释放占用（调用方须持 _rpa锁）。"""
    run_id = _rpa分支占用.get(branch)
    if not run_id:
        return
    记录 = _rpa运行.get(run_id)
    if not 记录 or 记录["proc"].poll() is not None:
        _rpa分支占用.pop(branch, None)


def _rpa读取循环(记录: dict) -> None:
    """后台线程：逐行收集子进程输出 → 内存 deque + 落盘日志文件（父子单一写者，无共享冲突）。"""
    行 = 记录["日志行"]
    总 = 记录["总行数"]
    proc = 记录["proc"]
    try:
        for 行文本 in iter(proc.stdout.readline, ""):
            行.append(行文本.rstrip("\n"))
            总[0] += 1
            try:
                记录["日志文件"].write(行文本)
                记录["日志文件"].flush()
            except Exception:
                pass
    except Exception:
        pass
    finally:
        try:
            记录["日志文件"].close()
        except Exception:
            pass


def _rpa记一行(记录: dict, 文本: str) -> None:
    """监视线程向运行记录补一行（deque + 总行数计数 + 落盘），保证 status 游标能取到。"""
    记录["日志行"].append(文本)
    记录["总行数"][0] += 1
    try:
        with open(记录["log"], "a", encoding="utf-8", errors="replace") as f:
            f.write(文本 + "\n")
    except Exception:
        pass


def _rpa导dws():
    """按需 import web_auto/公共_钉钉通知/dws_mail；失败返回 None（通知降级为跳过，不抛异常）。"""
    try:
        目录 = str(RPA通知目录)
        if 目录 not in sys.path:
            sys.path.insert(0, 目录)
        import dws_mail
        return dws_mail
    except Exception as exc:
        UI错误(f"导入 dws_mail 失败（钉钉通知不可用）：{exc}")
        return None


def _rpa导dws_msg():
    """按需 import 同目录 dws_msg（钉钉单聊消息封装）；失败返回 None（通知降级为跳过，不抛异常）。"""
    try:
        目录 = str(RPA通知目录)
        if 目录 not in sys.path:
            sys.path.insert(0, 目录)
        import dws_msg
        return dws_msg
    except Exception as exc:
        UI错误(f"导入 dws_msg 失败（钉钉消息不可用）：{exc}")
        return None


def _rpa导msg():
    """按需 import web_auto/公共_钉钉通知/dws_msg（钉钉截图消息）；失败返回 None。"""
    try:
        目录 = str(RPA通知目录)
        if 目录 not in sys.path:
            sys.path.insert(0, 目录)
        import dws_msg
        return dws_msg
    except Exception as exc:
        UI错误(f"导入 dws_msg 失败（截图消息不可用）：{exc}")
        return None


def _rpa最近截图(起始epoch: float, 最多: int = 3) -> list:
    """取 起始epoch 之后修改过的截图（RPA通知截图来源 两个目录），新的在前，最多 最多 张。"""
    图: list = []
    try:
        for 根, 模式 in RPA通知截图来源:
            for p in 根.glob(模式):
                try:
                    mtime = p.stat().st_mtime
                    if mtime >= 起始epoch - 10:      # 留 10s 裕量（脚本启动略早于记录建立）
                        图.append((mtime, p))
                except Exception:
                    continue
        图.sort(key=lambda x: x[0], reverse=True)
    except Exception:
        pass
    return [p for _, p in 图[:最多]]


def _rpa失败通知(记录: dict, 退出码) -> None:
    """进程非零退出且脚本内部未发过钉钉邮件时，网站侧补发失败通知（绝不抛异常打断主流程）。"""
    try:
        if 记录.get("手动停止"):
            return                      # 用户主动停止不算异常，不打扰
        行 = list(记录["日志行"])
        if any(标记 in 行文本 for 行文本 in 行 for 标记 in RPA通知防重标记):
            _rpa记一行(记录, "【钉钉】脚本内部已发送钉钉邮件，网页端不重复通知")
            return
        dws_mail = _rpa导dws()
        if dws_mail is None:
            _rpa记一行(记录, "【钉钉】dws_mail 不可用（web_auto/公共_钉钉通知 拷贝是否完整），跳过网页端补发")
            return
        截图 = _rpa最近截图(记录.get("起始epoch") or time.time())
        主题 = f"【计划平台操作】{记录['script']} 运行失败（网页端补发，exit={退出码}）"
        正文行 = [
            "提醒：网页端「计划平台Web自动化」节点运行失败，流程已中断！！！",
            "",
            f"分支：{记录['branch']}",
            f"脚本：{记录['script']}",
            f"run_id：{记录.get('run_id', '?')}",
            f"开始时间：{记录.get('started', '?')}",
            f"退出码：{退出码}（脚本内部未发钉钉邮件，故由网页端补发）",
            "",
            "日志尾部（最后 60 行）：",
            *["  " + 行文本 for 行文本 in 行[-60:]],
            "",
        ]
        if 截图:
            正文行 += ["页面截图（本次运行期间最新）："]
            正文行 += [f"[inline:{p.name}]" for p in 截图]
        else:
            正文行 += ["（本次运行未产生截图文件）"]
        状态, 说明 = dws_mail.发送邮件(主题, "\n".join(正文行), 内联图片=截图 or None)
        _rpa记一行(记录, f"【钉钉】失败通知：{状态}｜{说明}")
    except Exception as exc:
        try:
            _rpa记一行(记录, f"【钉钉】失败通知异常（不影响主流程）：{exc}")
        except Exception:
            pass


def _rpa监视线程(记录: dict) -> None:
    """等读取线程收尾 → 进程退出 → 需要补发时发送钉钉通知（通知期间 status 的 running 保持 True）。"""
    try:
        读取 = 记录.get("reader")
        if 读取:
            读取.join(timeout=60)
        proc = 记录["proc"]
        for _ in range(150):                 # 最多再等 30s（taskkill 进程树后退出可能略慢）
            if proc.poll() is not None:
                break
            time.sleep(0.2)
        退出码 = proc.poll()
        if 退出码 is None or 退出码 == 0 or 记录.get("手动停止"):
            return
        # 2026-10-25：计划平台Web自动化已改为「邮箱通知」开关控制，钉钉不再自动发送。
        # email_notify 开 → 失败也由流水线批次邮件统一发出（含失败结果文案）；
        # email_notify 关 → 按用户「关的状态不要发」不发送任何通知。二者都不走钉钉失败补发。
        if 记录.get("email_notify"):
            _rpa记一行(记录, "【钉钉】已停用钉钉失败补发（改用邮箱通知开关，失败会在流水线结束统一发邮件）")
        else:
            _rpa记一行(记录, "【钉钉】已按用户要求停用钉钉通知（邮箱开关关闭=不发送任何通知）")
        return
    except Exception as exc:
        UI错误(f"RPA 监视线程异常：{exc}")
    finally:
        记录["通知中"] = False


def _rpa截图步骤名(路径) -> str:
    """20261018_184205_2.1-X为0-重置后结束.png → 2.1-X为0-重置后结束（剥掉 日期_时间_ 两段前缀）。"""
    名 = Path(路径).stem
    部分 = 名.split("_")
    if len(部分) >= 3 and len(部分[0]) == 8 and 部分[0].isdigit() and len(部分[1]) == 6 and 部分[1].isdigit():
        return "_".join(部分[2:])
    return 名


def _rpa扫描新截图(记录: dict, 已推送: set) -> list:
    """扫 RPA通知截图来源 两个目录，返回本次运行期间新出现且未推送过的截图列表。"""
    新图: list = []
    起始 = 记录.get("起始epoch") or 0
    for 根, 模式 in RPA通知截图来源:
        try:
            for p in 根.glob(模式):
                try:
                    st = p.stat()
                except Exception:
                    continue
                if st.st_mtime < 起始 - 10:      # 只推本次运行期间产生的（留 10s 裕量）
                    continue
                键 = (str(p), st.st_mtime_ns, st.st_size)
                if 键 in 已推送:
                    continue
                已推送.add(键)
                新图.append(p)
        except Exception:
            continue
    新图.sort(key=lambda p: p.name)
    return 新图


def _rpa推送截图(记录: dict, 路径) -> None:
    """把一张整页截图推送到钉钉（叠加说明条 + file 消息）；结果写运行日志，绝不抛异常。"""
    try:
        dws_msg = _rpa导msg()
        if dws_msg is None:
            _rpa记一行(记录, "【钉钉】dws_msg 不可用（web_auto/公共_钉钉通知 是否完整），跳过截图推送")
            return
        脚本 = 记录["script"].split("_")[0]      # AB_34箱_xxx.py → AB
        步骤 = _rpa截图步骤名(路径)
        状态, 说明 = dws_msg.发送截图消息(路径, 脚本, 步骤)
        if 状态 == "已发送":
            _rpa记一行(记录, f"【钉钉】截图已推送：{脚本} · {步骤}（{Path(路径).name}）")
        else:
            _rpa记一行(记录, f"【钉钉】截图推送失败：{Path(路径).name}｜{说明}")
    except Exception as exc:
        _rpa记一行(记录, f"【钉钉】截图推送异常（不影响主流程）：{exc}")


# 中文数字 1~10（附件命名用）；超过 10 个回退为阿拉伯数字「附件N」
_中文数字 = "一二三四五六七八九十"



def _rpa节点截图归档(记录: dict) -> None:
    """单个节点 run 结束后：把本次运行新产生的截图收进所属流水线批次（不发送）。

    真正的邮件由 _rpa流水线发送() 在整条流水线结束时统一发一封，
    正文「步骤N：…　截图：见附件N」与附件「附件一/二…」严格一一对应。
    """
    pid = 记录.get("pipeline_id") or ""
    try:
        读取 = 记录.get("reader")
        if 读取:
            读取.join(timeout=60)
        proc = 记录["proc"]
        for _ in range(150):                 # 最多再等 30s（taskkill 进程树后退出可能略慢）
            if proc.poll() is not None:
                break
            time.sleep(0.2)
        time.sleep(5)                          # 退出瞬间可能还有收尾截图写出，等 5s 兜底
        # 去重：相邻节点时间窗有 10s 裕量重叠，同一张图不能进两遍（否则附件会重复）
        with _rpa流水线锁:
            线 = _rpa流水线.get(pid)
            已归档键 = 线.setdefault("已归档键", set()) if 线 else set()
            截图 = _rpa扫描新截图(记录, 已归档键)
        截图.sort(key=lambda p: (p.stat().st_mtime, p.name))

        节点名 = _rpa脚本步骤名(记录.get("script", ""))
        目录 = RPA_LOG_DIR / f"pl_{pid}"
        目录.mkdir(parents=True, exist_ok=True)
        新增: list = []
        for p in 截图:
            步骤名 = _rpa截图步骤名(p)
            目标 = 目录 / f"{len(新增):03d}_{uuid.uuid4().hex[:6]}{p.suffix}"
            try:
                shutil.copy2(p, 目标)
                新增.append((节点名, 步骤名, 目标))
            except Exception as exc:
                _rpa记一行(记录, f"【邮箱】截图归档失败（{p.name}）：{exc}")

        with _rpa流水线锁:
            线 = _rpa流水线.get(pid)
            if 线 is not None:
                线["截图"].extend(新增)
                线["节点"].append({"script": 记录.get("script", ""), "退出码": proc.poll(),
                                   "手动停止": bool(记录.get("手动停止"))})
                线["最后活动"] = time.time()
                线程数 = len(线["线程"])
            else:
                线程数 = 0
        _rpa记一行(记录, f"【邮箱】已归档本次截图 {len(新增)} 张到本轮邮件批次（累计第 {线程数} 个节点）")
        if not 新增:
            _rpa记一行(记录, "【邮箱】本次运行未产生截图文件")
    except Exception as exc:
        try:
            _rpa记一行(记录, f"【邮箱】截图归档异常（不影响主流程）：{exc}")
        except Exception:
            pass


def _rpa脚本步骤名(脚本名: str) -> str:
    """AB_34箱_接受任务清单_BOM解析.py → AB 34箱 接受任务清单 BOM解析（正文步骤前缀）。"""
    名 = Path(脚本名 or "").stem
    return 名.replace("_", " ").strip() or "节点"


def _rpa流水线取(pid: str) -> "dict | None":
    with _rpa流水线锁:
        return _rpa流水线.get(pid)


def _rpa流水线发送(pid: str, 结果文案: str = "", 标题: str = "",
                  明细: str = "", 开始时间: str = "") -> dict:
    """把一条流水线批次里收集到的全部截图拼成一封邮件发出，并清理批次。

    - 主题：<一级标题> 计划平台Web自动化
    - 正文：步骤1：<节点> <截图步骤>　截图：见附件一 …（与附件严格对应）
    - 附件：附件一.png / 附件二.png …

    2026-10-27：新增「无 RPA 节点」分支——像「预投物料需求调取及预测」这种整轮只有
    VBA/网页节点的流水线，从来不调 /api/rpa/run，批次根本没登记过；以前 finish 时
    直接回「该批次不存在或已发送」，页面上看着像发信失败。现在改为补一个空批次照发一封
    （正文用前端传上来的各节点结果明细，无附件），保证「点一次一级标题 = 一封邮件」。
    """
    有待发批次 = False
    with _rpa流水线锁:
        线 = _rpa流水线.get(pid)
        if 线 is not None and 线.get("已发送"):
            return {"ok": True, "sent": False, "说明": "该批次邮件已发送过"}
        if 线 is None and pid in _rpa已发送批次:
            return {"ok": True, "sent": False, "说明": "该批次邮件已发送过"}
        if 线 is None:
            # 纯 VBA/网页运行：没有任何截图可归档，补一个空批次（无附件，正文走 明细）
            线 = {"task_title": 标题, "email_to": "demo.user58@example.com",
                 "started": 开始时间, "截图": [], "节点": [], "线程": [],
                 "已归档键": set(), "最后活动": time.time(), "已发送": False,
                 "无RPA节点": True, "明细": 明细}
            _rpa流水线[pid] = 线
        else:
            有待发批次 = True
            if 明细:
                线["明细"] = 明细
        线程: list = [t for t in 线.get("线程", []) if t is not None]

    # 等所有节点的归档线程收尾（各节点已结束，通常几秒内）
    for t in 线程:
        try:
            if t.is_alive():
                t.join(timeout=45)
        except Exception:
            pass

    with _rpa流水线锁:
        线 = _rpa流水线.get(pid)
        if not 线:
            return {"ok": True, "sent": False, "说明": "该批次已清理"}
        线["已发送"] = True
        截图 = list(线.get("截图", []))
        节点列表 = list(线.get("节点", []))
        主题前缀 = 线.get("task_title") or ""
        收件人 = 线.get("email_to") or "demo.user59@example.com"
        开始 = 线.get("started") or ""
        步骤明细 = (线.get("明细") or "").strip()
        _rpa已发送批次[pid] = time.time()
        _rpa已发送批次清理()

    主题 = f"{主题前缀} 计划平台Web自动化" if 主题前缀 else "计划平台Web自动化"
    if not 结果文案:
        失败 = [n for n in 节点列表 if n.get("退出码") not in (0,)]
        结果文案 = "成功（全部节点 exit 0）" if 节点列表 and not 失败 else (
            f"中断（{len(节点列表)} 个节点已执行，最后退出码 {失败[-1].get('退出码')}）" if 失败 else "—")

    临时目录 = RPA_LOG_DIR / f"email_attach_{pid}"
    临时附件: list = []
    try:
        临时目录.mkdir(parents=True, exist_ok=True)
        for i, (_, _, p) in enumerate(截图):
            序号 = _中文数字[i] if i < len(_中文数字) else str(i + 1)
            目标 = 临时目录 / f"附件{序号}{p.suffix}"      # 文件名即附件名，dws 以文件名作附件名
            try:
                shutil.copy2(p, 目标)
                临时附件.append(目标)
            except Exception as exc:
                UI错误(f"RPA 邮件附件准备失败（{p.name}）：{exc}")

        行 = [
            f"「{主题前缀 or '计划平台Web自动化'}」已执行完毕。",
            "",
            f"运行结果：{结果文案}",
        ]
        if 开始:
            行.append(f"开始时间：{开始}")
        行.append("")
        if 截图:
            行.append(f"执行步骤与截图（共 {len(截图)} 步，截图见附件，序号一一对应）：")
            for i, (节点名, 步骤名, _) in enumerate(截图):
                序号 = _中文数字[i] if i < len(_中文数字) else str(i + 1)
                标题 = f"{节点名} — {步骤名}" if 步骤名 else 节点名
                行.append(f"步骤{i + 1}：{标题}　截图：见附件{序号}")
        elif 步骤明细:
            # 纯 VBA/网页运行（无截图）：正文列出前端上报的各节点结果，邮件仍然有内容可读
            行.append("执行步骤与结果（本轮无浏览器截图，故无附件）：")
            行.extend(步骤明细.splitlines())
        else:
            行.append("（本次运行未产生截图文件）")

        dws_mail = _rpa导dws()
        if dws_mail is None:
            UI错误("RPA 邮件发送失败：dws_mail 不可用（web_auto/公共_钉钉通知 拷贝是否完整）")
            return {"ok": False, "sent": False, "说明": "dws_mail 不可用"}
        状态, 说明 = dws_mail.发送邮件(主题, "\n".join(行), 收件人=收件人, 普通附件=临时附件 or None)
        UI信息(f"RPA 流水线邮件：{状态}｜{说明}（{主题}，{len(截图)} 步 / {len(临时附件)} 个附件）")
        结果 = {"ok": 状态 == "已发送", "sent": 状态 == "已发送", "状态": 状态, "说明": 说明,
                "主题": 主题, "步骤数": len(截图), "附件数": len(临时附件)}
        if 状态 != "已发送" and 临时附件:
            # 没真发出去（dws 未登录/网络异常）→ 邮件原文已落盘待发邮件/，
            # 这里把附件按同样的 附件一/二… 名另存到它旁边，人工补发时正文与附件仍然对得上。
            件目录 = RPA通知目录 / "待发邮件" / f"附件_{time.strftime('%Y%m%d_%H%M%S')}_{pid}"
            try:
                件目录.mkdir(parents=True, exist_ok=True)
                for t in 临时附件:
                    shutil.copy2(t, 件目录 / t.name)
                结果["附件备份"] = str(件目录)
                UI错误(f"RPA 邮件未真正发出（{状态}）：正文已落盘待发邮件/，附件已另存 {件目录}")
            except Exception as exc:
                UI错误(f"RPA 邮件附件另存失败：{exc}")
        return 结果
    except Exception as exc:
        UI错误(f"RPA 流水线邮件发送异常：{exc}")
        return {"ok": False, "sent": False, "说明": str(exc)}
    finally:
        for t in 临时附件:
            try:
                t.unlink(missing_ok=True)
            except Exception:
                pass
        shutil.rmtree(临时目录, ignore_errors=True)
        with _rpa流水线锁:
            批次 = _rpa流水线.pop(pid, None)
        if 批次:
            shutil.rmtree(RPA_LOG_DIR / f"pl_{pid}", ignore_errors=True)


def _rpa流水线兜底定时器(pid: str) -> None:
    """前端没调 finish（页面关闭/崩溃）时的兜底：批次内已无存活进程且空闲超阈值，才自动发一封。"""
    try:
        # ① 本批次还有节点进程在跑（长任务可能一两个小时）→ 绝不能提前发，续期等待
        仍有进程 = False
        with _rpa锁:
            for 记录 in _rpa运行.values():
                if 记录.get("pipeline_id") == pid:
                    try:
                        if 记录["proc"].poll() is None:
                            仍有进程 = True
                            break
                    except Exception:
                        pass
        if 仍有进程:
            续 = threading.Timer(20, _rpa流水线兜底定时器, args=(pid,))
            续.daemon = True
            续.start()
            return

        线 = _rpa流水线取(pid)
        if not 线 or 线.get("已发送"):
            return
        剩余 = RPA流水线空闲兜底秒 - (time.time() - (线.get("最后活动") or 0))
        if 剩余 > 0:
            续 = threading.Timer(min(剩余 + 1, RPA流水线空闲兜底秒), _rpa流水线兜底定时器, args=(pid,))
            续.daemon = True
            续.start()
            return
        UI信息(f"RPA 流水线邮件兜底发送（前端未上报结束）：{pid}")
        _rpa流水线发送(pid, "（前端未上报结束，按空闲超时自动发送）")
    except Exception as exc:
        UI错误(f"RPA 流水线兜底定时器异常：{exc}")


def _rpa流水线登记(记录: dict) -> threading.Thread:
    """把本节点 run 登记进所属流水线批次，返回（未启动的）截图归档线程。

    首批（本轮流水线第一个节点）时额外挂一个空闲兜底定时器：前端若没上报结束（关页面/崩溃），
    批次内已无存活进程且空闲超阈值后自动补发一封，避免「跑了却没收到邮件」。
    """
    pid = 记录["pipeline_id"]
    线程 = threading.Thread(target=_rpa节点截图归档, args=(记录,), daemon=True)
    with _rpa流水线锁:
        线 = _rpa流水线.get(pid)
        首批 = 线 is None
        if 线 is None:
            线 = {"task_title": 记录.get("task_title", ""), "email_to": 记录.get("email_to", ""),
                 "started": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "截图": [], "节点": [], "线程": [], "已归档键": set(),
                 "最后活动": time.time(), "已发送": False}
            _rpa流水线[pid] = 线
        else:
            if 记录.get("task_title"):
                线["task_title"] = 记录["task_title"]
            线["最后活动"] = time.time()
        线["线程"].append(线程)
    if 首批:
        兜底 = threading.Timer(RPA流水线空闲兜底秒 + 5, _rpa流水线兜底定时器, args=(pid,))
        兜底.daemon = True
        兜底.start()
    return 线程


@app.post("/api/rpa/pipeline/finish")
def rpa_pipeline_finish(payload: RpaPipelineFinishRequest):
    """流水线结束（成功/失败/停止/提前收工）→ 统一发一封汇总邮件。"""
    pid = (payload.pipeline_id or "").strip()
    if not pid:
        raise HTTPException(status_code=400, detail="pipeline_id 不能为空")
    return _rpa流水线发送(pid, (payload.status or "").strip(),
                        标题=(payload.task_title or "").strip(),
                        明细=(payload.details or "").strip(),
                        开始时间=(payload.started or "").strip())


@app.post("/api/rpa/run")
def rpa_run(payload: RpaRunRequest):
    """启动一个 RPA 脚本（branch + script 白名单校验；单分支单实例）。"""
    branch = (payload.branch or "").strip()
    script = (payload.script or "").strip()
    配置 = RPA_BRANCHES.get(branch)
    if not 配置:
        raise HTTPException(status_code=400,
                            detail=f"未知分支 {branch!r}（可用：{'、'.join(RPA_BRANCHES)}）")
    相对路径 = 配置.get(script)
    if not 相对路径:
        raise HTTPException(status_code=400, detail=f"脚本 {script!r} 不在分支 {branch} 的白名单")
    脚本路径 = (RPA_BASE / 相对路径).resolve()
    if not str(脚本路径).startswith(str(RPA_BASE.resolve())):   # 防路径穿越双保险
        raise HTTPException(status_code=400, detail=f"脚本路径越界：{script!r}")
    if not 脚本路径.is_file():
        raise HTTPException(status_code=500, detail=f"脚本文件不存在：{脚本路径}（web_auto 拷贝是否完整）")

    with _rpa锁:
        _rpa回收分支(branch)
        已有 = _rpa分支占用.get(branch)
        if 已有:
            raise HTTPException(status_code=409,
                                detail=f"分支 {branch} 已有运行中的任务（run_id={已有}）："
                                       f"所有分支共用同一个已登录 Chrome，请先停止或等待完成")

    解释器 = _rpa选解释器()
    run_id = uuid.uuid4().hex[:12]
    RPA_LOG_DIR.mkdir(parents=True, exist_ok=True)
    日志路径 = RPA_LOG_DIR / f"{run_id}.log"
    环境 = dict(os.environ)
    环境["PYTHONIOENCODING"] = "utf-8"       # 子进程输出统一 UTF-8，避免 Windows 控制台 GBK 报错
    环境.setdefault("DEBUG_MODE", "1")       # RPA 脚本调试输出默认开（网页日志看板需要）
    try:
        日志文件 = open(日志路径, "w", encoding="utf-8", errors="replace")
        proc = subprocess.Popen(
            [解释器, str(脚本路径)],
            cwd=str(脚本路径.parent),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            env=环境,
            encoding="utf-8",
            errors="replace",
            bufsize=1,                       # 行缓冲：日志近实时可见
            creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        )
    except Exception as 异常:
        try:
            日志文件.close()
        except Exception:
            pass
        UI错误(f"RPA 启动失败（{script}）：{异常}")
        raise HTTPException(status_code=500, detail=f"RPA 启动失败：{异常}")

    记录 = {
        "branch": branch, "script": script, "proc": proc, "log": 日志路径,
        "日志行": deque(maxlen=RPA_LOG_WINDOW), "总行数": [0], "日志文件": 日志文件,
        "started": datetime.now().strftime("%H:%M:%S"),
        "email_notify": bool(payload.email_notify),
        "email_to": (payload.email_to or "demo.user60@example.com").strip() or "demo.user61@example.com",
        "task_title": (payload.task_title or "").strip(),
        "pipeline_id": (payload.pipeline_id or "").strip() or run_id,
    }
    读取线程 = threading.Thread(target=_rpa读取循环, args=(记录,), daemon=True)
    读取线程.start()
    记录["reader"] = 读取线程
    记录["run_id"] = run_id
    记录["手动停止"] = False
    记录["通知中"] = False
    记录["起始epoch"] = time.time()
    threading.Thread(target=_rpa监视线程, args=(记录,), daemon=True).start()
    # 2026-10-25：计划平台Web自动化通知改为「邮箱通知」开关控制，且按「点一次一级标题 = 一封邮件」聚合。
    # 开 → 本节点截图归档进 pipeline_id 批次，流水线结束时由 /api/rpa/pipeline/finish 统一发一封；
    # 关 → 既不发钉钉也不发邮件。
    if 记录["email_notify"]:
        _rpa流水线登记(记录).start()
    with _rpa锁:
        _rpa运行[run_id] = 记录
        _rpa分支占用[branch] = run_id
    UI信息(f"RPA 启动：{script}（run_id={run_id}，解释器={解释器}）")
    return {"ok": True, "run_id": run_id, "branch": branch, "script": script,
            "interpreter": 解释器}


@app.get("/api/rpa/status")
def rpa_status(run_id: str = Query(..., description="运行 ID"), after: int = Query(0, description="已消费日志行数（游标）")):
    """查询运行状态；after 游标之后的增量日志随本次返回（窗口内不丢行）。"""
    with _rpa锁:
        记录 = _rpa运行.get(run_id)
        if 记录:
            _rpa回收分支(记录["branch"])
    if not 记录:
        raise HTTPException(status_code=404, detail=f"运行不存在或已清理：{run_id}")
    退出码 = 记录["proc"].poll()
    通知中 = bool(记录.get("通知中"))       # 钉钉补发通知发送期间，前端应继续轮询（running 保持 True）
    行 = list(记录["日志行"])
    总行数 = 记录["总行数"][0]
    窗口起点 = 总行数 - len(行)
    截断 = after > 0 and after < 窗口起点
    增量 = 行[max(0, after - 窗口起点):] if not 截断 else 行
    return {
        "run_id": run_id,
        "branch": 记录["branch"],
        "script": 记录["script"],
        "running": 退出码 is None or 通知中,
        "notifying": 通知中,
        "exit_code": 退出码,
        "started": 记录["started"],
        "lines": 增量,
        "next": 总行数,
        "truncated": 截断,
    }


@app.post("/api/rpa/stop")
def rpa_stop(payload: RpaStopRequest):
    """终止运行：Windows 杀进程树（脚本自愈重拉的子进程一并杀掉）。"""
    with _rpa锁:
        记录 = _rpa运行.get(payload.run_id)
    if not 记录:
        raise HTTPException(status_code=404, detail=f"运行不存在或已清理：{payload.run_id}")
    proc = 记录["proc"]
    退出码 = proc.poll()
    if 退出码 is not None:
        with _rpa锁:
            _rpa回收分支(记录["branch"])
        return {"ok": True, "already_exited": True, "exit_code": 退出码}
    记录["手动停止"] = True                  # 主动停止不算异常：监视线程不再补发钉钉通知
    try:
        subprocess.run(["taskkill", "/T", "/F", "/PID", str(proc.pid)],
                       capture_output=True, timeout=15)
    except Exception:
        try:
            proc.terminate()
        except Exception:
            pass
    for _ in range(30):        # 最多等 6s 确认退出
        if proc.poll() is not None:
            break
        time.sleep(0.2)
    with _rpa锁:
        _rpa回收分支(记录["branch"])
    UI信息(f"RPA 停止：{记录['script']}（run_id={payload.run_id}，exit={proc.poll()}）")
    return {"ok": True, "exit_code": proc.poll()}


class RpaTestNotifyRequest(BaseModel):
    branch: str = "f30_34"


@app.post("/api/rpa/test_notify")
def rpa_test_notify(payload: RpaTestNotifyRequest):
    """发送一封测试邮件，验证「异常截图 → 钉钉邮件」链路（dws 可用性 / 登录态 / 收件人）。"""
    branch = (payload.branch or "").strip() or "f30_34"
    if branch not in RPA_BRANCHES:
        raise HTTPException(status_code=400,
                            detail=f"未知分支 {branch!r}（可用：{'、'.join(RPA_BRANCHES)}）")
    dws_mail = _rpa导dws()
    if dws_mail is None:
        raise HTTPException(status_code=500,
                            detail="dws_mail 导入失败（web_auto/公共_钉钉通知 拷贝是否完整）")
    主题 = "【计划平台操作】网页端钉钉通知链路测试"
    正文 = ("这是一封来自「03#箱工作站 · 计划平台Web自动化」的测试邮件。\n\n"
            f"发送时间：{datetime.now():%Y-%m-%d %H:%M:%S}\n"
            f"目标分支：{branch}\n\n"
            "如果你在钉钉收到本邮件，说明异常截图通知链路（dws）工作正常：\n"
            "此后脚本运行失败时，网页端会自动把日志尾部与页面截图发送到本邮箱。")
    状态, 说明 = dws_mail.发送邮件(主题, 正文)
    UI信息(f"RPA 测试通知：{状态}｜{说明}")
    return {"ok": 状态 == "已发送", "状态": 状态, "说明": 说明}


def _rpa生成测试图() -> "Path | None":
    """PIL 画一张通道测试卡（与运行期截图消息同款 file 消息样式）；PIL 不可用返回 None。"""
    try:
        from PIL import Image, ImageDraw, ImageFont
        字体 = next((p for p in [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf"]
                     if Path(p).exists()), None)
        if 字体 is None:
            return None
        图 = Image.new("RGB", (960, 540), (13, 20, 33))
        画 = ImageDraw.Draw(图)
        大 = ImageFont.truetype(字体, 34)
        小 = ImageFont.truetype(字体, 20)
        画.text((40, 56), "计划平台截图消息通道测试", font=大, fill=(233, 243, 255))
        画.text((40, 150), "「计划平台Web自动化」运行期整页截图即按此样式推送", font=小, fill=(169, 191, 228))
        画.text((40, 200), f"收件人：{RPA消息收件人}（self=当前登录用户演示用户）", font=小, fill=(169, 191, 228))
        画.text((40, 250), f"时间：{datetime.now():%Y-%m-%d %H:%M:%S}", font=小, fill=(169, 191, 228))
        画.text((40, 310), "收到本图 = 上传/发送/说明条叠加全链路正常", font=小, fill=(46, 230, 168))
        目录 = RPA_LOG_DIR / "dingtalk_push"
        目录.mkdir(parents=True, exist_ok=True)
        路径 = 目录 / f"_test_{datetime.now():%Y%m%d_%H%M%S}.png"
        图.save(路径)
        return 路径
    except Exception as exc:
        UI错误(f"测试图生成失败：{exc}")
        return None


@app.post("/api/rpa/test_msg")
def rpa_test_msg(payload: RpaTestNotifyRequest):
    """按运行期截图推送的同款链路发一条测试消息（文本+带说明条的图片），验证「整页截图 → 钉钉消息」通道。"""
    branch = (payload.branch or "").strip() or "f30_34"
    if branch not in RPA_BRANCHES:
        raise HTTPException(status_code=400,
                            detail=f"未知分支 {branch!r}（可用：{'、'.join(RPA_BRANCHES)}）")
    dws_msg = _rpa导msg()
    if dws_msg is None:
        raise HTTPException(status_code=500,
                            detail="dws_msg 导入失败（web_auto/公共_钉钉通知 拷贝是否完整）")
    图 = _rpa生成测试图()
    if 图 is None:
        raise HTTPException(status_code=500, detail="测试图生成失败（PIL 或中文字体不可用）")
    状态1, 说明1 = dws_msg.发送文本消息(
        "【计划平台Web自动化】截图消息通道测试：下面这条文件消息即运行期整页截图的推送样式")
    状态2, 说明2 = dws_msg.发送截图消息(图, "测试", "截图消息通道测试")
    UI信息(f"RPA 测试截图消息：文本={状态1} 截图={状态2}")
    return {"ok": 状态2 == "已发送", "文本状态": 状态1, "截图状态": 状态2, "说明": 说明2}


@app.post("/api/rpa/page_shot")
async def rpa_page_shot(file: UploadFile = File(...), caption: str = Form("")):
    """接收「网站任务页截图」（前端 html2canvas 实时截 #autoPage：节点状态+运行日志），
    叠加说明条后以钉钉文件消息发给演示用户（dws_msg 同款链路）。
    用途：①手动点任务页「截图发钉钉」②运行中每个节点结束时自动截推（远程盯进度）。"""
    内容 = await file.read()
    if not 内容 or len(内容) < 500:
        raise HTTPException(status_code=400, detail="截图内容为空或过小（前端截图失败）")
    目录 = RPA通知目录 / "page_shots"
    目录.mkdir(parents=True, exist_ok=True)
    文件名 = f"page_{datetime.now():%Y%m%d_%H%M%S_%f}.png"
    (目录 / 文件名).write_bytes(内容)
    dws_msg = _rpa导msg()
    if dws_msg is None:
        raise HTTPException(status_code=500,
                            detail="dws_msg 导入失败（web_auto/公共_钉钉通知 拷贝是否完整）")
    说明 = (caption or "").strip() or "网站任务页截图"
    状态, 摘要 = dws_msg.发送截图消息(str(目录 / 文件名), "APD-F30-34#箱自动化", 说明)
    UI信息(f"RPA 页面截图推送：{状态}｜{摘要[-120:]}")
    return {"ok": 状态 == "已发送", "状态": 状态, "说明": 摘要, "文件": 文件名}


# ============ 预投自动化状态图 · 运行快照（2026-10-17 用户明令新增） ============
# 每次自动化任务运行结束后，前端把「六功能块节点状态总览」上报存档到
# statusmap_store/<时间戳>.json；状态图页下拉框可查看不同时间快照，双击功能块
# 跳转对应任务页。

STATUSMAP_DIR = APP_DIR / "statusmap_store"


class StatusMapSnapshot(BaseModel):
    """运行快照：taskId/name/timestamp/ok + blocks（六功能块各自节点状态摘要）+ image（整页截图 dataURL，可选）。"""
    taskId: str
    taskName: str
    ok: bool
    cost: str = ""
    meta: str = ""
    blocks: List[dict]          # [{page, name, state, nodes:[{name,state}], log}]
    log: str = ""
    image: str = ""             # data:image/png;base64,... 整页截图（2026-10-17 用户要求：快照=页面完整截图）


def _statusmap_读(文件: Path) -> dict:
    try:
        return json.loads(文件.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _statusmap_存(数据: dict) -> str:
    """快照落盘：image 是 dataURL 时拆成 <时间戳>.png 单独存（JSON 里只留图片文件名），避免 base64 撑爆 JSON。"""
    STATUSMAP_DIR.mkdir(parents=True, exist_ok=True)
    时间戳 = datetime.now().strftime("%Y%m%d_%H%M%S")
    image = 数据.pop("image", "") or ""
    if isinstance(image, str) and image.startswith("data:image/"):
        ext = image[11:].split(";", 1)[0].split("/", 1)[-1] or "png"
        ext = "".join(ch for ch in ext if ch.isalnum())[:5] or "png"
        (STATUSMAP_DIR / f"{时间戳}.{ext}").write_bytes(base64.b64decode(image.split(",", 1)[1]))
        数据["imageFile"] = f"{时间戳}.{ext}"
    else:
        数据.pop("imageFile", None)
    数据["timestamp"] = 时间戳
    (STATUSMAP_DIR / f"{时间戳}.json").write_text(json.dumps(数据, ensure_ascii=False), encoding="utf-8")
    return 时间戳


def _statusmap_补图URL(数据: dict) -> dict:
    """读取侧：把 imageFile 补成可直接 <img> 用的 /api/statusmap/image?timestamp=… 地址。"""
    f = 数据.get("imageFile")
    if f:
        数据["imageUrl"] = "/api/statusmap/image?timestamp=" + str(f).rsplit(".", 1)[0]
    return 数据


@app.post("/api/statusmap/save")
def statusmap_save(payload: StatusMapSnapshot):
    """保存一次运行快照（文件名=时间戳，保留全部历史供下拉回看；带整页截图时一并落盘 PNG）。"""
    时间戳 = _statusmap_存(payload.model_dump())
    UI信息(f"状态图快照已保存：{时间戳}（{payload.taskName}）")
    return {"ok": True, "timestamp": 时间戳}


@app.get("/api/statusmap/list")
def statusmap_list():
    """全部快照清单（新→旧），供下拉框。"""
    if not STATUSMAP_DIR.exists():
        return {"ok": True, "snapshots": []}
    快照 = []
    for f in sorted(STATUSMAP_DIR.glob("*.json"), reverse=True):
        数据 = _statusmap_读(f)
        if not 数据:
            continue
        快照.append({
            "timestamp": 数据.get("timestamp", f.stem),
            "taskId": 数据.get("taskId", ""),
            "taskName": 数据.get("taskName", ""),
            "ok": bool(数据.get("ok", False)),
            "cost": 数据.get("cost", ""),
        })
    return {"ok": True, "snapshots": 快照}


@app.get("/api/statusmap/get")
def statusmap_get(timestamp: str):
    """读取指定时间戳的快照详情（带截图时附 imageUrl）。"""
    if not _re.fullmatch(r"[0-9]{8}_[0-9]{6}", timestamp or ""):
        raise HTTPException(status_code=400, detail="非法时间戳")
    文件 = STATUSMAP_DIR / f"{timestamp}.json"
    if not 文件.exists():
        raise HTTPException(status_code=404, detail="快照不存在")
    return _statusmap_补图URL(_statusmap_读(文件))


@app.get("/api/statusmap/image")
def statusmap_image(timestamp: str):
    """读取快照整页截图 PNG（2026-10-17 用户要求：快照=页面完整截图）。"""
    if not _re.fullmatch(r"[0-9]{8}_[0-9]{6}", timestamp or ""):
        raise HTTPException(status_code=400, detail="非法时间戳")
    for ext in ("png", "jpeg", "webp"):
        f = STATUSMAP_DIR / f"{timestamp}.{ext}"
        if f.exists():
            media = "image/png" if ext == "png" else f"image/{ext}"
            return FileResponse(str(f), media_type=media, headers={"Cache-Control": "no-cache"})
    raise HTTPException(status_code=404, detail="该快照无截图")


# ============ SAPGUI 自动化脚本（VBS 经 cscript 执行） ============
# 与 RPA（web_auto 里的 .py）互补：SAP GUI 前端脚本是 .vbs，
# 由网页「SAPGUI自动化脚本」菜单各页触发，白名单控制可执行范围。

VBS_BASE = Path(r"D:\ai_project\GUI\ZPP_PO")     # VBS 脚本根目录（白名单边界）
VBS_LOG_DIR = APP_DIR / "vbs_logs"
VBS_TIMEOUT = 600                                  # 单次运行超时（秒）：SAP 慢/多屏数据时留足（2026-10-23 300→600）

# 脚本名 → 相对 VBS_BASE 的文件名；与左侧菜单「SAPGUI自动化脚本 → CO28」页对应
VBS_SCRIPTS = {
    # 2026-10-23 v3：由 v2 复制改造（v2 原样保留）。
    # v2：先点全选按钮 btn[42]，再跨屏（OK 码 P+ 翻页）反选不符合的行
    #     （物料 *-*-*-03* / 描述含「安全钳」或含「框架」但描述含"~-"或"反绳轮框架"的框架件豁免不反选 / 基本开始距今<=1天）。
    # v3：在 v2 基础上，下达（wnd[1] sendVKey 8）结果总览渲染稳定后，
    #     按 CO28_MAIL_MODE 开关（后端按次注入为 1）跨滚动条多屏截图，供「执行[且发邮件]」发送结果邮件。
    "co28": "co28_有问题复查_v3_下达截图.vbs",
}

# ============ 2026-10-23：三个「选择屏查询」页接入执行链（用户明令） ============
# 用户原话：「看一下实际输入的条件，然后接入到执行按钮；我希望跟 CO28 一样，
#           我在 UI 界面输入什么内容，相应的自动更改输入的条件，然后在 SAP 当中
#           同步更改后的条件，然后并执行」
# 做法与 CO28 完全同构：原脚本不动 → 每次执行按页面现值改写「运行时副本」→ cscript 跑副本。
# 与 CO28 的差别只有一处：这三个脚本放在网站目录 app/vbs_scripts/（用户要求的副本位置），
# 所以白名单需要第二个 base；字段映射不写死在 Python 里，而是读 VBS 行尾的 [key:xxx] 标记。
VBS_SITE_BASE = APP_DIR / "vbs_scripts"
VBS_SCRIPTS_SITE = {
    # 键: (文件名, 显示名, 事务码, 页面锚点, 邮件主题动作词)
    "zb02":    ("Zb02.vbs",    "Zb02（ZLSTMA）",  "/nzb02", "#zb02SapScreen", "查询"),
    "zb03":    ("Zb03.vbs",    "Zb03（ZLSTBM）",  "/nzb03", "#zb03SapScreen", "查询"),
    "zpp_po":  ("Zpp_Po.vbs",  "Zpp_Po（ZPP_PO）", "/nse16", "#zppScreen",    "查询"),
    # 2026-10-24：Zpp_Pur（SE16 表 ZPP_PUR 选择屏复刻）
    "zpp_pur": ("Zpp_Pur.vbs", "Zpp_Pur（ZPP_PUR）", "/nse16", "#zppPurScreen", "查询"),
    # 2026-10-24：照用户截图复刻的 SE16 查询页（字段 ID = 只读探针实测，脚本 gen_sap_vbs.py 规格驱动生成）
    "marc":    ("Marc.vbs",    "Marc（MARC）",    "/nse16", "#marcSapScreen", "查询"),
    "mara":    ("Mara.vbs",    "Mara（MARA）",    "/nse16", "#maraSapScreen", "查询"),
    "eord":    ("Eord.vbs",    "Eord（EORD）",    "/nse16", "#eordSapScreen", "查询"),
    "zbom":    ("Zpp_Bom.vbs", "Zpp_Bom（ZPP_BOM）", "/nse16", "#zbomSapScreen", "查询"),
    "resb":    ("Resb.vbs",    "Resb（RESB）",    "/nse16", "#resbSapScreen", "查询"),
    "plaf":    ("Zpp_Plaf.vbs", "Zpp_Plaf（ZPP_PLAF）", "/nse16", "#plafSapScreen", "查询"),
    "eina":    ("Eina.vbs",    "Eina（EINA）",    "/nse16", "#einaSapScreen", "查询"),
    "zlstma":  ("Zlstma_Log_02.vbs", "Zlstma_Log_02（ZLSTMA_LOG_02）", "/nse16", "#zlstmaSapScreen", "查询"),
    "zlstbm":  ("Zlstbm_Log_02.vbs", "Zlstbm_Log_02（ZLSTBM_LOG_02）", "/nse16", "#zlstbmSapScreen", "查询"),
    # 2026-10-24 用户：「zpp_cs12 我需要你复刻，并且 VBS 代码集成到里面，参考表查询的 UI 结构 不要参考 CO28」
    # + 纠正「/nzpp_cs12 是事务码，不是 SE16 表查询」→ 直进事务码，不写表名
    "zppcs12": ("Zpp_Cs12.vbs", "Zpp_Cs12（/nzpp_cs12）", "/nzpp_cs12", "#zppcs12SapScreen", "查询"),
    # 2026-10-24 用户：「请你继续把这个内部表复刻出来」——ZPP_ZLSTMA（16 条件 + 输出控制，只读探针实测 ID）
    "zppzlstma": ("Zpp_Zlstma.vbs", "Zpp_Zlstma（ZPP_ZLSTMA）", "/nse16", "#zppzlstmaSapScreen", "查询"),
    # 2026-10-25 用户：「SQ00 改名为 SQ00配方查询，把 SQ00PEIF.vbs 集成进去」
    #   SQ00 = 查询事务（查询 ZB001「配方查询」）：初始屏 → 显示 → 标题格式 → F8 选择屏（12 条件 + 格式）
    #   选择屏字段 ID 全部来自只读探针 dump（_probe_sq00_b/e.txt），未凭猜测。
    "sq00":    ("Sq00.vbs",    "SQ00配方查询（ZB001）", "/nsq00", "#sq00SapScreen", "查询"),
    # 2026-10-29 用户：「复刻 Zpp_Pr2po 事务码，自动集成 VBS 代码」——采购申请批量转换采购订单（/nzpp_pr2po）
    #   选择屏字段 ID 来自只读探针（_probe_pr2po.txt；探针_Zpp_Pr2po取字段ID.vbs 产出），未凭猜测。
    "zpppr2po": ("Zpp_Pr2po.vbs", "Zpp_Pr2po（/nzpp_pr2po）", "/nzpp_pr2po", "#zpppr2poSapScreen", "转换"),
    # 2026-11-03：限速器物料号每日同步专用脚本（Zb02.vbs 拷贝改造；不走网页页面，只给触发器用）
    "zlstma_sync": ("Zlstma_Sync.vbs", "Zlstma_Sync（ZLSTMA 限速器物料同步）", "/nzb02", "", "导出"),
    # 2026-11-03：安全钳物料号每日同步专用脚本（同 Zlstma_Sync 工艺）
    "safclamp_sync": ("SafClamp_Sync.vbs", "SafClamp_Sync（ZLSTMA 安全钳物料同步）", "/nzb02", "", "导出"),
}

# 2026-10-24：批量语法（A,B / A..B / !A / !A..B）不再一律跳过 —— 这几个站点脚本尾部有
# BatchWrite 代码块（打开 SAP「多重选择」框写入，见 vbs_scripts/batch_write.inc.vbs），
# 注入器把该行改写成 BatchWrite "<字段ID>", "<语法串>"。没打这个补丁的站点仍按原样跳过并记账。
VBS_批量写入支持 = {"zpp_pur", "zb02", "zb03", "zpp_po", "marc", "mara", "eord", "zbom", "resb", "plaf", "eina",
                    "zlstma", "zlstbm", "zppzlstma", "zppcs12", "sq00", "zpppr2po",
                    # 2026-11-03 补：两个同步脚本（Zb02 拷贝版带 BatchWrite 块）——
                    # 22:00 首跑事故：漏登记 → 注入器把 *限速器*,!*通用件*… 原样留在输入框（! 没翻译进排除页签）、
                    # 32 个图号也被跳过 → 0 命中停在选择屏 → 导出链空跑失败
                    "zlstma_sync", "safclamp_sync"}

# 必填条件（2026-10-24 用户现场：「这个空在那儿不行的，必须要输入一个东西」）
#   物料号在 SAP 屏上是黄色必填字段（C_MATNR-LOW），空着按 F8 直接报错。
#   这里登记「脚本键 → 必填 data-key 集合」：注入时若页面值为空，保留脚本里的非空默认值，
#   绝不把空的必填项带进 SAP。页面侧也会在点「执行查询」前拦住空值（双重保险）。
VBS_必填键 = {
    "zppcs12": {"zppcs12CmatnrLow"},
}


def _是必填行(脚本键: str, key: str) -> bool:
    return key in VBS_必填键.get(脚本键, set())


def _vbs站点截图目录(脚本键: str) -> Path:
    """每个页面的截图目录（每次运行前清空，供结果邮件内联）。"""
    return VBS_LOG_DIR / f"{脚本键}_shots"


# ---------- 报表下载（2026-10-23 用户需求）----------
# 用户原话：「报表的命名和保存位置，参考的下载报表的 VBS 代码（D:\ai_project\GUI\ZPP_PO\下载报表同用.vbs），
#            我固定写成了 ZB02 加时间轴，我希望你根据功能 + 时间戳来保存」「C:\Users\u3854\Desktop\web_xls」。
# 命名 = <功能名>_<YYYYMMDD_HHMMSS>.xls —— 由后端按次生成并注入运行时 VBS，
# 保证「磁盘上的文件」与「邮件附件」同名同源（用户要求：两个开关都开时报表作为附件发邮件）。
SAPQ_下载目录 = Path(r"C:\Users\u3854\Desktop\web_xls")
SAPQ_报表后缀 = ".xls"
SAPQ_报表前缀 = {"zb02": "Zb02", "zb03": "Zb03", "zpp_po": "Zpp_Po", "zpp_pur": "Zpp_Pur", "co28": "CO28",
                 "marc": "Marc", "mara": "Mara", "eord": "Eord", "zbom": "Zpp_Bom", "resb": "Resb",
                 "plaf": "Zpp_Plaf", "eina": "Eina", "zlstma": "Zlstma_Log_02", "zlstbm": "Zlstbm_Log_02",
                 "zppzlstma": "Zpp_Zlstma", "zppcs12": "Zpp_Cs12", "sq00": "Sq00",
                 "zpppr2po": "Zpp_Pr2po",
                 "zlstma_sync": "Zlstma_Sync",
                 "safclamp_sync": "SafClamp_Sync",
                 # 2026-10-24：修改模块「zpp_cpgl」按钮（不带事务码页，脚本在 vbs_scripts/Zpp_Cpgl.vbs）
                 "zpp_cpgl": "Zpp_Cpgl"}


def _sapq报表文件名(脚本键: str, 时刻: "datetime | None" = None) -> str:
    """报表文件名 = 功能名 + 时间戳（用户 2026-10-23 口径）。"""
    前缀 = SAPQ_报表前缀.get(脚本键) or 脚本键
    return f"{前缀}_{(时刻 or datetime.now()):%Y%m%d_%H%M%S}{SAPQ_报表后缀}"


def _sapq解析下载目录(文本: "str | None") -> Path:
    """页面「下载路径」框的值 → 落点目录（用户 2026-10-23：可自定义，优先级高于代码里的默认路径）。

    留空 = 用代码默认（SAPQ_下载目录 = C:/Users/u3854/Desktop/web_xls）。
    支持环境变量（%USERPROFILE%）与 ~；相对路径或建不出来的路径一律**回落默认**并记账，
    绝不因为路径写错就让整次执行失败（用户口径：默认路径兜底，自定义优先）。
    """
    原 = str(文本 or "").strip().strip('"').strip("'")
    if not 原:
        return _sapq确保下载目录()
    try:
        展开 = os.path.expandvars(os.path.expanduser(原))
        路径 = Path(展开)
        if not 路径.is_absolute():
            UI错误(f"下载路径必须是绝对路径，已回落默认：{原}")
            return _sapq确保下载目录()
        return _sapq确保下载目录(路径)
    except Exception as exc:  # noqa: BLE001
        UI错误(f"下载路径无法使用（{原}）：{exc}；已回落默认 {SAPQ_下载目录}")
        return _sapq确保下载目录()


def _sapq确保下载目录(目录: "Path | str | None" = None) -> Path:
    """报表下载落点（不存在则建；建不了也不抛，交调用方兜底报错）。"""
    目标 = Path(目录) if 目录 else SAPQ_下载目录
    try:
        目标.mkdir(parents=True, exist_ok=True)
    except Exception as exc:  # noqa: BLE001
        UI错误(f"报表下载目录创建失败（{目标}）：{exc}；已回落默认 {SAPQ_下载目录}")
        if 目标 != SAPQ_下载目录:
            目标 = SAPQ_下载目录
            try:
                目标.mkdir(parents=True, exist_ok=True)
            except Exception:  # noqa: BLE001
                pass
    return 目标


def _sapq解析收件人(文本: str) -> List[str]:
    """收件人文本 → 邮箱列表（多个用 , ; 空格 / 换行 分隔；非法项丢弃并记账）。

    用户 2026-10-23：「我希望可以自己填写收件人邮箱，具体你参考触发器里面邮箱触发器，
    我可以自定义发送邮箱，可以发送给多个人」。
    与「触发器 · 运行错误处理」的多收件人（收件箱 chips）同口径：
    发送方式固定 = 钉钉邮箱（dws，发件人 = 当前钉钉账号自己的企业邮箱），收件人可自定义多个。
    """
    原 = str(文本 or "").strip()
    if not 原:
        return []
    结果: List[str] = []
    for x in _re.split(r"[,;，；\s]+", 原):
        x = x.strip()
        if not x:
            continue
        if not _re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", x):
            UI错误(f"收件人格式非法已忽略：{x}")
            continue
        if x not in 结果:
            结果.append(x)
    return 结果


def _sapq取最新报表(脚本键: str, 目录: "Path | str | None" = None) -> "Path | None":
    """最近一次下载出来的报表文件：按「功能名_时间戳」匹配，取最新一个。

    先看页面自定义的「下载路径」（有则优先），再看代码默认目录 —— 用户改过路径又改回来时
    两边都能取到，不会「明明下载过却说没有」。
    """
    前缀 = SAPQ_报表前缀.get(脚本键) or 脚本键
    候选目录 = []
    if 目录:
        候选目录.append(Path(目录))
    if Path(SAPQ_下载目录) not in 候选目录:
        候选目录.append(SAPQ_下载目录)
    for 一个 in 候选目录:
        try:
            if not 一个.is_dir():
                continue
            候选 = sorted([p for p in 一个.glob(f"{前缀}_*{SAPQ_报表后缀}") if p.is_file()],
                          key=lambda x: x.stat().st_mtime, reverse=True)
            if 候选:
                return 候选[0]
        except Exception:  # noqa: BLE001
            continue
    return None


def _sapq读报表行(报表路径) -> List[List[str]]:
    """把 SAP 下载下来的报表（.xls = GBK+Tab 文本，或 .xlsx）读成「行 → 列」二维列表。失败返回空列表。"""
    p = Path(报表路径)
    if not p.is_file():
        return []
    try:
        if p.suffix.lower() == ".xlsx":
            import openpyxl
            # read_only=True 会一直持有文件句柄，不 close 会导致随后「把这份 xlsx 作为邮件附件」
            # 或「自动打开」时因文件被本进程锁住而失败；务必读完就关。
            wb = openpyxl.load_workbook(str(p), data_only=True, read_only=True)
            try:
                ws = wb.active
                行 = []
                for r in ws.iter_rows(values_only=True):
                    行.append(["" if c is None else str(c) for c in r])
            finally:
                wb.close()
        else:
            # SAP 报表本质是 GBK 编码的 Tab 分隔文本，却叫 .xls（见 _是伪装的文本报表）
            文本 = p.read_text(encoding="gbk", errors="replace")
            行 = [ln.split("\t") for ln in 文本.splitlines()]
        return [r for r in 行 if any((c or "").strip() for c in r)]
    except Exception as exc:  # noqa: BLE001
        UI错误(f"读取报表失败（{p.name}）：{exc}")
        return []


def _sapq提取成功行(报表路径) -> "Tuple[List[str], int]":
    """从报表里挑出「采购订单生成成功」的行，整理成可直接进邮件正文 / 粘贴 Excel 的文本块。

    返回 (文本行列表, 成功行数)。
      · 无成功行 → 文本块只有一句「无采购订单生成成功」，行数 0；
      · 读不到报表 → (None, -1)，由调用方决定怎么提示。
    """
    行 = _sapq读报表行(报表路径)
    if not 行:
        return None, -1
    # 表头行：含「采购申请号 / 物料号 / 供应商 / 采购组织」等稳定列名的第一行
    # （SAP 列表导出常在真正表头前塞一个标题行，所以不能直接拿第 1 行当表头）
    表头下标 = next((i for i, r in enumerate(行)
                   if any(k in (c or "") for c in r
                          for k in ("采购申请号", "物料号", "供应商", "采购组织"))), 0)
    表头 = 行[表头下标]
    数据 = 行[表头下标 + 1:]
    成功 = [r for r in 数据 if any("采购订单生成成功" in (c or "") for c in r)]
    if not 成功:
        return ["无采购订单生成成功"], 0
    块 = [f"采购订单生成成功的行（共 {len(成功)} 行；下方为制表符分隔，可直接粘贴到 Excel）：", ""]
    块.append("\t".join(表头))
    for r in 成功:
        块.append("\t".join(str(c) for c in r))
    return 块, len(成功)


def _归一化代号(代号: str) -> str:
    """供应商代号归一化：去前后空格 + 去全部前导 0（与「停产供方」写入口径一致）。"""
    return (代号 or "").strip().lstrip("0")


def _停产供方_代号集合() -> set:
    """读取「基础数据 > 供应商基础数据 > 停产供方」里所有供应商代号（已去前导 0），返回集合；读不到返回空集合。"""
    try:
        store = _safety_store_load()
        rows = (store.get("sheets") or {}).get(停产供方_SHEET) or []
        s = set()
        for r in rows:
            d = r.get("data") or {}
            c = str(d.get(停产供方_CODE) or "").strip()
            if c:
                s.add(_归一化代号(c))
        return s
    except Exception:  # noqa: BLE001
        return set()


def _sapq定位表头(行: List[List[str]]):
    """在报表二维行里定位表头行下标与关键列下标。

    返回 (表头下标, {关键词: 列下标})；找不到表头返回 (-1, {})。
    SAP 列表导出常在真正表头前塞一个标题行，所以不能直接拿第 1 行当表头。
    """
    表头下标 = next((i for i, r in enumerate(行)
                   if any(k in (c or "") for c in r
                          for k in ("采购申请号", "物料号", "供应商", "采购组织", "处理结果"))), -1)
    if 表头下标 < 0:
        return -1, {}
    表头 = 行[表头下标]
    列 = {}
    for 关键词 in ("供应商", "物料号", "处理结果"):
        for j, c in enumerate(表头):
            if 关键词 in (c or ""):
                列[关键词] = j
                break
    return 表头下标, 列


def _sapq筛选行(报表路径) -> "Optional[List[List[str]]]":
    """2026-11-06 用户需求：从 Zpp_Pr2po 报表里挑出「需重点关注」的行，整理成邮件正文表格。

    筛选条件（任一命中即纳入）：
      ① 物料号以 P125 开头 且 供应商 = 201253；
      ② 物料号以 P183 开头 且 供应商 = 201007；
      ③ 供应商属于「停产供方」（基础数据 > 供应商基础数据 > 停产供方 的供应商代号）。
    返回 [表头] + 命中数据行（全列）；无命中返回 [表头]（调用方据此提示「无符合条件数据」）；
    读不到报表或找不到表头返回 None。
    """
    行 = _sapq读报表行(报表路径)
    if not 行:
        return None
    表头下标, 列 = _sapq定位表头(行)
    if 表头下标 < 0:
        return None
    表头 = 行[表头下标]
    数据 = 行[表头下标 + 1:]
    供应_idx = 列.get("供应商", 0)
    物料_idx = 列.get("物料号", 5)
    停产 = _停产供方_代号集合()
    命中 = []
    for r in 数据:
        sup = (r[供应_idx] if 供应_idx < len(r) else "").strip()
        mat = (r[物料_idx] if 物料_idx < len(r) else "").strip()
        sup_n = _归一化代号(sup)
        if (mat.upper().startswith("P125") and sup_n == "201253"
                or mat.upper().startswith("P183") and sup_n == "201007"
                or sup_n in 停产):
            命中.append(r)
    return [表头] + 命中


def _sapq正文表格文本(表格: List[List[str]]) -> List[str]:
    """把 [表头] + 数据行 渲染成纯文本表格（│ 分隔 + 表头分隔线），用于邮件正文（邮件通道为纯文本）。"""
    if not 表格:
        return []
    表头 = 表格[0]
    数据 = 表格[1:]
    出 = []
    出.append(" │ ".join((c or "") for c in 表头))
    出.append("─" * max(24, sum(len(str(c or "")) + 3 for c in 表头)))
    if not 数据:
        出.append("（本次没有符合筛选条件的数据）")
        return 出
    for r in 数据:
        出.append(" │ ".join((c or "") for c in r))
    return 出


def _sapq整理报表xlsx(报表路径) -> dict:
    """2026-11-06 用户需求：把 Zpp_Pr2po 的报表 xlsx 整理成「好看的 Excel」再作附件。

    ① 删除「处理结果」列值以「已转换数量不为0，无法全量转换」开头的行（整批未转换的提示行，无意义）；
    ② 列宽自适应（按内容最大显示宽度，CJK 记 2 宽）；
    ③ 冻结表头行（字段抬头冻结）。
    输出到「原名_已整理.xlsx」，绝不覆盖用户原来的报表文件。失败如实返回，绝不抛异常。
    """
    源 = Path(报表路径)
    if not 源.is_file() or 源.suffix.lower() != ".xlsx":
        return {"ok": False, "说明": f"需要 xlsx 才能整理（当前：{源.name}）", "路径": str(源)}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(源))
        try:
            ws = wb.active
            行 = []
            for r in ws.iter_rows(values_only=True):
                行.append(["" if c is None else str(c) for c in r])
            表头下标, 列 = _sapq定位表头(行)
            if 表头下标 < 0:
                return {"ok": False, "说明": "找不到表头行，无法整理（仍用原文件作附件）", "路径": str(源)}
            处理_idx = 列.get("处理结果")
            删 = []
            if 处理_idx is not None:
                for ri in range(表头下标 + 1, len(行)):
                    v = (行[ri][处理_idx] if 处理_idx < len(行[ri]) else "")
                    if (v or "").strip().startswith("已转换数量不为0"):
                        删.append(ri + 1)   # openpyxl 行号从 1 开始
                for ri in sorted(删, reverse=True):
                    ws.delete_rows(ri, 1)
            # 列宽自适应（CJK 记 2 宽）
            _字母 = openpyxl.utils.get_column_letter
            for col_idx in range(1, ws.max_column + 1):
                mx = 0
                for cell in ws[_字母(col_idx)]:
                    v = "" if cell.value is None else str(cell.value)
                    w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in v)
                    mx = max(mx, w)
                ws.column_dimensions[_字母(col_idx)].width = min(max(mx + 2, 8), 60)
            # 冻结表头（表头行号 + 1）
            ws.freeze_panes = f"A{表头下标 + 2}"
            目标 = 源.parent / (源.stem + "_已整理.xlsx")
            wb.save(str(目标))
            return {"ok": True,
                    "说明": f"{源.name} → {目标.name}（已删 {len(删)} 行、列宽自适应、冻结表头）",
                    "路径": str(目标), "删行数": len(删)}
        finally:
            wb.close()
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "说明": f"整理异常：{exc}", "路径": str(源)}


def _vbs注入常量(文本: str, 名: str, 值: str) -> str:
    """把运行时 VBS 里「Const <名> = ...」那一行整行替换为新值（找不到则原样返回）。

    与 CO28 的 MAIL_MODE/SHOT_DIR 注入同一手法，但用正则按常量名定位，
    常量行怎么改写（补注释、改对齐）都不会失配。
    """
    # 注意 [^\r\n]* 而不是 .*$：后者会把行尾的 \r 一起吃掉，注出行就变成裸 LF（CRLF 被破坏）。
    模式 = _re.compile(r'(?m)^Const[ \t]+' + _re.escape(名) + r'[ \t]*=[ \t]*[^\r\n]*')
    if not 模式.search(文本):
        return 文本
    return 模式.sub(lambda _m: f"Const {名} = {值}", 文本, count=1)


def 生成选择屏运行时脚本(脚本键: str, 条件: dict, 邮件模式: bool = False,
                         下载模式: bool = False, 报表文件: str = "",
                         跳过查询: bool = False, 下载目录: "Path | str | None" = None):
    """把网页选择屏现值改写成运行时 VBS 副本（原脚本一行不动）。

    注入依据 = 脚本每行尾的 `[key:网页data-key]` 标记（由 verify_gen_vbs_scripts.py 生成）：
        session.findById("wnd[0]/usr/ctxtI1-LOW").text = "DM*"   ' 物料·下限 [key:zb03MatLow]
    后端只替换 `[key:xxx]` 对应那一行的 `=` 右侧字面值，其余（含注释、未启用行）原样保留。
    行首带 ' 的未启用行不匹配（正则要求行首即 session），所以永远不会被误改。

    返回 (运行时文件路径, 摘要 dict)。
    """
    文件名, 显示名, _, _, _ = VBS_SCRIPTS_SITE[脚本键]
    源 = (VBS_SITE_BASE / 文件名).resolve()
    文本 = 源.read_bytes().decode("gbk", errors="replace")
    摘要 = {"script": 脚本键, "fields": [], "skipped": [], "missing": [], "unmatched": [], "mail": bool(邮件模式),
            "download": bool(下载模式), "skipQuery": bool(跳过查询), "报表文件": 报表文件,
            "下载目录": str(_sapq确保下载目录(下载目录)) if 下载模式 else ""}

    def _校验(fid, key, 值, 回读=True):
        """写入后立刻自证（2026-10-24 ZPP_ZLSTMA 复盘）：出错或回读不符就打成 [fill-fail] 行。
        以前界面只报「已写入 N 项」（按页面现值数出来的），SAP 里其实一条都没落上，用户看到的却是「查询完成」。"""
        缩 = "    "
        换行 = chr(13) + chr(10)
        out = ('%sIf Err.Number <> 0 Then WScript.Echo "[fill-fail] %s 写入报错: " & Err.Description : Err.Clear'
               % (缩, key))
        if 回读:
            out += (换行 + '%sIf Err.Number = 0 Then If Trim(session.findById("wnd[0]/usr/%s").text) <> Trim("%s") '
                    'Then WScript.Echo "[fill-fail] %s 回读不符（期望 %s）"' % (缩, fid, 值, key, 值))
        return 换行 + out

    def _替换(m):
        key = m.group("key")
        if key not in 条件:
            摘要["missing"].append(key)
            return m.group(0)          # 页面没这个键 → 保留脚本自带默认值
        新值 = _vbs清值(条件.get(key))
        # 🔴 必填字段不许被空值覆盖（2026-10-24 用户现场：「这个空在那儿不行的，必须要输入一个东西」）。
        #    物料号在 SAP 屏上是黄色必填字段，空着 F8 会直接报错；脚本里该行带了一个非空默认占位，
        #    所以页面对应的键若是空串 → 保留脚本自己的默认值，并把这件事明确记进摘要。
        if not 新值 and _是必填行(脚本键, key):
            摘要["skipped"].append(key + "（必填，页面为空 → 保留脚本默认值）")
            return m.group(0)
        if _vbs含批量语法(新值):
            # 2026-10-24：有 BatchWrite 补丁的站点改走「多重选择」框写入（用户现场：网页填了多值 SAP 里却是空白）
            if 脚本键 in VBS_批量写入支持:
                摘要["fields"].append((key, f"{新值}（多重选择框）"))
                # 第三参 = 网页 data-key：VBS 里写失败时按 [batch-fail] <key> 点名上报，
                # 后端再如实告诉用户「N 项未写入 SAP」（2026-10-27 用户现场：报成功其实没写进去）
                return ('BatchWrite "{0}", "{1}", "{2}"'.format(m.group("fid"), 新值, key) + m.group("tail")
                        + _校验(m.group("fid"), key, 新值, 回读=False))
            # 未打补丁的站点：SAP 单选屏放不下 → 跳过该行并记账（原行为）
            摘要["skipped"].append(key)
            return m.group(0)
        摘要["fields"].append((key, 新值))
        return f'{m.group("head")}{新值}"{m.group("tail")}' + _校验(m.group("fid"), key, 新值)

    行正则 = _re.compile(
        r'(?m)^(?P<head>[ \t]*session\.findById\("wnd\[0\]/usr/(?P<fid>[^"]+)"\)\.text = ")'
        r'(?P<val>[^"]*)"'
        r'(?P<tail>[^\r\n]*\[key:(?P<key>[A-Za-z0-9_]+)\])')
    文本 = 行正则.sub(_替换, 文本)

    # 🔴 勾选类字段（复选框 / 单选项）单独注入（2026-10-24 Zpp_Cs12 新增）。
    #    它们写的是 `.Selected = CBool("0" = "1")` 而不是 `.text = "..."`，上面那条「写文本」正则天生匹配不到；
    #    若不处理，页面勾了「匹配汇总」、SAP 里永远不勾 —— 又是「界面说成功、SAP 里没生效」。
    #    页面侧 data-key 存的是 "1"/"0"，这里翻译成 CBool("1" = "1")=True / CBool("0" = "1")=False。
    勾选正则 = _re.compile(
        r'(?m)^(?P<head>[ \t]*session\.findById\("wnd\[0\]/usr/(?P<fid>[^"]+)"\)\.Selected = )CBool\([^)]*\)'
        r'(?P<tail>[^\r\n]*\[key:(?P<key>[A-Za-z0-9_]+)\])')

    def _替换勾选(m):
        key = m.group("key")
        if key not in 条件:
            摘要["missing"].append(key)
            return m.group(0)
        原 = _vbs清值(条件.get(key)).strip()
        真 = 原 in ("1", "true", "True", "TRUE", "on", "是", "y", "Y")
        摘要["fields"].append((key, "勾选" if 真 else "不勾选"))
        # 复选框（GuiCheckBox）可以自由设 True/False，不受单选框那种「同组互斥」限制
        return f'{m.group("head")}CBool("{"1" if 真 else "0"}" = "1"){m.group("tail")}'

    文本 = 勾选正则.sub(_替换勾选, 文本)

    # 🔴 单选项（radio）：写法是「仅当值为 1 时才点选」——
    #     If "0" = "1" Then session.findById(".../radXXX").Selected = True    ' 标签 [key:xxx]
    #     为什么这么做：同一组里对未选中项设 .Selected = False 会抛
    #     "SAP Frontend Server: The method got an invalid argument."（2026-10-24 实测）。
    #     单选框只能靠「选中某一项」互斥，所以值为 0 时整条语句不执行、保持原状。
    单选正则 = _re.compile(
        r'(?m)^(?P<head>[ \t]*If )"(?P<val>[01]?)"(?P<mid> = "1" Then session\.findById\("wnd\[0\]/usr/(?P<fid>[^"]+)"\)'
        r'\.Selected = True)'
        r'(?P<tail>[^\r\n]*\[key:(?P<key>[A-Za-z0-9_]+)\])')

    def _替换单选(m):
        key = m.group("key")
        if key not in 条件:
            摘要["missing"].append(key)
            return m.group(0)
        原 = _vbs清值(条件.get(key)).strip()
        真 = 原 in ("1", "true", "True", "TRUE", "on", "是", "y", "Y")
        摘要["fields"].append((key, "选中" if 真 else "不选中（保持原状）"))
        return f'{m.group("head")}"{"1" if 真 else "0"}"{m.group("mid")}{m.group("tail")}'

    文本 = 单选正则.sub(_替换单选, 文本)

    # 防御（2026-10-24 ZPP_ZLSTMA 翻车复盘）：行尾带 [key:xxx] 却匹配不上上面「写字段」正则的行 = 字段 ID 形态异常
    # （典型：探针只回裸 ID ctxtI1-LOW，少了 wnd[0]/usr/ 前缀）。这种条件既注入不了、也落不到 SAP 屏上，
    # 必须大声报出来，绝不能静默当成功——否则用户看到的就是「填了条件却说查询完成、SAP 里一片空白」。
    #
    # 2026-10-27 修两处误报：
    #   ① 注释行不算 —— 脚本头部说明文字里写着「行尾 [key:xxx] 是后端注入标记」，
    #      那是散文不是字段。以前全文扫描把注释里的字面量 [key:xxx] 也算成「1 项未能注入」，
    #      用户为此白查了半天。判定：整行去掉空白后以 ' 开头 = 注释行，跳过。
    #   ② skipped 里的 key 带后缀（如「xxx（必填，页面为空 → 保留脚本默认值）」），
    #      取「（」之前的部分才是真正的 key。
    已认领 = {k for k, _ in 摘要["fields"]}
    已认领 |= {str(_s).split("（")[0].strip() for _s in 摘要["skipped"]}
    已认领 |= set(摘要["missing"])
    for _行 in 文本.splitlines():
        if _行.strip().startswith("'"):
            continue                      # 注释行里的 [key:xxx] 是说明文字，不是字段
        for _m in _re.finditer(r"\[key:(?P<key>[A-Za-z0-9_]+)\]", _行):
            if _m.group("key") not in 已认领 and _m.group("key") not in 摘要["unmatched"]:
                摘要["unmatched"].append(_m.group("key"))

    # 邮件模式：打开截图开关 + 把截图目录指向本次运行目录（VBS 运行时会自行清空该目录）
    if 邮件模式:
        目录 = _vbs站点截图目录(脚本键)
        目录.mkdir(parents=True, exist_ok=True)
        if "MAIL_MODE" in 文本:
            文本 = _vbs注入常量(文本, "MAIL_MODE", "1")
        else:
            摘要["skipped"].append("未找到 MAIL_MODE 常量，截图可能不生效")
        旧目录 = _re.search(r'(?m)^Const[ \t]+SHOT_DIR[ \t]*=[ \t]*".*?"', 文本)
        if 旧目录:
            文本 = 文本[:旧目录.start()] + (
                f'Const SHOT_DIR   = "{str(目录.resolve()).replace("/", chr(92))}"'
                "   ' 截图输出目录（后端按次注入运行目录）") + 文本[旧目录.end():]
        else:
            摘要["skipped"].append("未找到 SHOT_DIR 常量，截图目录未注入")

    # 略过查询模式（「报表下载」单独点）：SAP 当前屏已经是查询结果，只走导出链
    文本 = _vbs注入常量(文本, "SKIP_QUERY", "1" if 跳过查询 else "0")

    # 报表下载模式：打开导出开关 + 注入落点目录与「功能名_时间戳.xls」文件名
    # 文件名由后端生成（不是 VBS 自己拼），保证磁盘文件与邮件附件同名同源。
    if 下载模式:
        目录 = _sapq确保下载目录(下载目录)   # 页面「下载路径」优先，空则代码默认
        文本 = _vbs注入常量(文本, "DL_MODE", "1")
        文本 = _vbs注入常量(文本, "DL_DIR", '"' + str(目录).replace("/", chr(92)) + '"')
        文本 = _vbs注入常量(文本, "DL_FILE", '"' + _vbs清值(报表文件) + '"')
    else:
        文本 = _vbs注入常量(文本, "DL_MODE", "0")

    运行时 = VBS_SITE_BASE / f"{脚本键}_runtime.vbs"
    运行时.write_bytes(文本.encode("gbk", errors="replace"))
    return 运行时, 摘要


class VbsRunRequest(BaseModel):
    script: str = "co28"
    params: Optional[dict] = None   # 网页选择屏现值；每次执行改写进运行时 VBS（原脚本不动）
    mail: bool = False              # 「发邮件」开关：执行后发送结果邮件（2026-10-23 由「执行[且发邮件]」拆出）
    收件人: str = ""                # 收件人（多个用 , / ; / 空格分隔；留空 = 默认 demo.user62@example.com）
    download: bool = False          # 「报表下载」开关：执行后把报表导出（功能名_时间戳.xls）
    skipQuery: bool = False         # 单独点「报表下载」：跳过查询，直接导出当前 SAP 屏幕上的报表
    下载目录: str = ""              # 页面「下载路径」框：留空 = 代码默认；自定义优先（用户 2026-10-23）
    钉钉: bool = False              # 「钉钉消息」开关（2026-10-24 用户需求）：执行后把结果摘要发钉钉单聊
    钉钉收件人: str = "演示用户"          # 页面「钉钉消息」框：留空/演示用户 = 当前钉钉登录用户本人（组织内 3 个重名，按本人最稳）
    报表转换: bool = False           # 「报表转换为xlsx」开关（2026-10-24）：下载完成后把 .xls 用 xlwings 另存为同名 .xlsx
    报表转换路径: str = ""            # 页面「转换路径」框：留空 = 单向跟随「报表下载」路径（下载路径变了它跟着变）
    自动打开: bool = False                  # 「自动打开」开关：执行+下载后自动用本机默认程序打开报表（2026-10-24 用户需求）
    # 2026-11-05 用户：SAP 相关「计划任务」在 SAP 执行前先连按 2 次 Enter 热键。
    # 默认 0 = 网页手动点「执行查询」完全不受影响；只有下面 4 个触发器链路传 2。
    前置回车: int = 0
    # 2026-11-05 用户：「我只要最新这份邮件，不要两份」——计划任务关掉 SAP 脚本自带的
    # 「仅下载报表」结果邮件，只留任务最后那封汇总通知；默认 True = 网页手动执行照旧。
    报表邮件: bool = True
    # 🔴🔴 2026-11-06 用户明令：「MESE限速器以及MESE安全钳 我不需要你发邮件给我」
    #   True = 本次执行一律不调用发信（结果邮件 / 「仅报表下载也发一封给 songjian」的兜底分支
    #          / 失败邮件，全部掐掉），只由 MESE 两条线内部驱动 SAP 下载时传入。
    #   🔴 默认 False：SAP 查询页手动执行、计划任务、CO28 的邮件口径一律不变，绝不扩散。
    不发本人邮件: bool = False


class Co28FailMailRequest(BaseModel):
    """CO28 执行失败时，前端把「当前网站页面」截图 + 日志一起回传，后端合并 SAP 当前屏发失败邮件。"""
    网页截图: Optional[str] = None    # data:image/png;base64,... 或 data/...;base64,
    日志: str = ""
    输出: str = ""
    主题: Optional[str] = None
    收件人: str = ""                 # 多个用 , / ; / 空格分隔；留空 = dws 默认


# CO28「执行[且发邮件]」：邮件主题（成功 / 失败），收件人固定 demo.user63@example.com（dws_mail 默认）
CO28_MAIL_主题成功 = "CO28下级零件(24*) 基本开工时间>=当天 【非安全钳组件及安全钳框架】，下达成功"
CO28_MAIL_主题失败 = "CO28下级零件(24*) 基本开工时间>=当天 【非安全钳组件及安全钳框架】，下达失败"
CO28_SHOTS_DIR = VBS_LOG_DIR / "co28_shots"      # 同一目录：VBS 每次运行前清空、重新截图


VBS_PROC = None                 # 当前运行中的 cscript 进程句柄（供 /api/vbs/stop 强杀）
VBS_STOP_REQUESTED = False      # 用户点了「重置」→ 下次 /api/vbs/stop 生效标记
VBS_PROC_LOCK = threading.Lock()

# CO28 选择屏字段关键名 → 网页参数键（Low/High）；VBS 注入段按关键名扫描 -LOW/-HIGH 输入框
CO28_FIELD_MAP = [
    ("MATNR", "mat",   "物料"),
    ("WERKS", "werks", "工厂"),
    ("DISPO", "dispo", "MRP控制器"),
    ("FEVOR", "fevor", "生产调度员"),
    ("ECKST", "gstrp", "基本开始日期"),   # CO28 屏上「表头级别的绝对日期」字段实际名 S_ECKST-LOW/HIGH
    ("ECKEN", "gltrp", "基本完成日期"),   # 对应 SAP 屏上 S_ECKEN-LOW/HIGH（GSTRP/GLTRP 在 CO28 不出现）
]


def _vbs清值(v) -> str:
    """页面值清洗：去首尾空白、去双引号（防 VBScript 语句注入）。"""
    return str(v or "").strip().replace('"', "'")


def _是千分位数字(v: str) -> bool:
    """判断是不是「带千分位分隔符的一个数」（如 9,000,000 / 1,234），而不是「多个值」。

    2026-10-27 现场坑：网页「最大命中数量」默认 9,000,000，被逗号判定当成批量语法，
    于是走进多重选择框 —— 而那个字段根本没有多选箭头 → 必然失败 → 字段空着。
    规则（全为数字才可能是千分位）：
      · 第一段 1~3 位，且不许有前导零（"002,115" 是两个值，不是千分位；
        真千分位不会写成 002,115）；
      · 后面每一段必须正好 3 位。
    """
    段 = str(v or "").strip().split(",")
    if len(段) < 2:
        return False
    if not all(p.isdigit() for p in 段):
        return False
    if len(段[0]) > 3:
        return False
    if len(段[0]) > 1 and 段[0][0] == "0":
        return False
    return all(len(p) == 3 for p in 段[1:])


def _vbs含批量语法(v: str) -> bool:
    """批量箭头序列化语法（A,B / A..B / !A / !A..B / ~A）——单值输入框放不下，改走 SAP 的多重选择框。

    千分位数字不算批量语法（见 _是千分位数字）：9,000,000 是一个数，不是九个值。
    ~A 是「强制按单一值」的转义（网页在单值本身含「..」或开头是「!」时自动加），
    必须也算批量语法——否则会被当成普通单值塞进输入框，落屏的是带 ~ 的错误值。
    """
    s = str(v or "").strip()
    if not s:
        return False
    段 = [x.strip() for x in s.split(",")]
    if any(x.startswith("~") or x.startswith("!~") for x in 段):
        return True
    if ".." in s or "!" in s:
        return True
    if "," in s:
        return not _是千分位数字(s)
    return False


# ---------- 收尾子过程（2026-10-23 用户需求）----------
# 用户原话：「自执行完 GUI 脚本以后，右下角这边一直处于运行状态，我希望每次指令完成以后，这个立即给关闭。」
# 现象：脚本跑完（或半途退出）后，SAP GUI 右下角仍挂着"运行中/忙"小指示。
# 成因：① 残留弹窗把会话挂在模态态；② SAP 端仍在异步收尾（清单导出/保存转换）；
#       ③ 脚本进程结束时 COM 引用（含 ConnectObject 事件挂接）没释放。
# 做法（无害，不碰业务数据）：关残留弹窗（查询类）/ 短暂等待 / 断开事件挂接并释放全部 COM 对象。
# 与 app/vbs_scripts 里三个报表脚本的 EndClean 保持同一份逻辑（改一处要同步另一处）。
VBS_CLEAN_SUB = """
' ============================================================
' 收尾（2026-10-23 用户需求）
' ------------------------------------------------------------
' 用户原话：「自执行完 GUI 脚本以后，右下角这边一直处于运行状态，我希望每次指令完成以后，
'           这个立即给关闭。」
' 成因：残留弹窗 / SAP 异步收尾未完成 / COM 引用（含事件挂接）未释放。
' 参数/变量名必须纯 ASCII（cscript 按 ANSI 读取，中文标识符会报「无效字符」）。
' ============================================================
Sub EndClean(closePopups)
    Dim i, n, w
    On Error Resume Next
    If closePopups And IsObject(session) Then
        n = 0
        n = session.Children.Count
        If n > 1 Then
            For i = n - 1 To 1 Step -1
                Set w = Nothing
                Set w = session.Children(i)
                If Not (w Is Nothing) Then
                    WScript.Echo "[clean] 关闭残留窗口：" & w.Text
                    w.Close
                    WScript.Sleep 300
                End If
            Next
        End If
    End If
    WScript.Sleep 400
    If IsObject(WScript) Then
        WScript.DisconnectObject session
        WScript.DisconnectObject application
    End If
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
    WScript.Echo "[clean] 收尾完成：COM 引用已释放，SAP GUI 的忙指示随即复位"
    On Error GoTo 0
End Sub
"""

def 生成co28运行时脚本(条件: dict, 邮件模式: bool = False):
    """把网页条件改写进 CO28 运行时 VBS 副本（原脚本不动，每次执行重新生成）。

    改动 1：P_AUTFP / P_PROFID 直接替换脚本里的写死值；
    改动 2：整段移除「打开变式搜索 → 选变式 u2418」流程（2026-10-19 用户明令：
            前端筛选条件已写全，运行时不再依赖变式），替换为直接写选择屏字段
            （-LOW/-HIGH，按 MATNR/WERKS/DISPO/FEVOR/ECKST/ECKEN 关键名扫描定位，
            找不到则回显警告），然后照常点执行。
    改动 3（v3，2026-10-23「执行[且发邮件]」）：邮件模式时把 CO28_MAIL_MODE 置 1
            （下达后跨滚动条多屏截图）并把 CO28_SHOT_DIR 指向本次运行目录。
    返回 (运行时文件路径, 注入摘要 dict)。
    """
    源 = (VBS_BASE / VBS_SCRIPTS["co28"]).resolve()
    文本 = 源.read_bytes().decode("gbk", errors="replace")
    摘要 = {"autfp": None, "prof": None, "fields": [], "skipped": [], "mail": bool(邮件模式)}

    # v3：邮件模式 → 打开截图开关并把截图目录指向本次运行目录（VBS 每次运行前会自行清空该目录）
    if 邮件模式:
        CO28_SHOTS_DIR.mkdir(parents=True, exist_ok=True)
        旧开关 = "Const CO28_MAIL_MODE = 0"
        if 旧开关 in 文本:
            文本 = 文本.replace(旧开关, "Const CO28_MAIL_MODE = 1", 1)
        else:
            摘要["skipped"].append("未找到 CO28_MAIL_MODE 常量，邮件模式可能不生效")
        旧目录 = _re.search(r'(?m)^Const[ \t]+CO28_SHOT_DIR[ \t]*=[ \t]*".*?"',
                           文本)
        if 旧目录:
            截图目录VBS = str(CO28_SHOTS_DIR.resolve()).replace("/", "\\")
            文本 = 文本[:旧目录.start()] + (
                f'Const CO28_SHOT_DIR  = "{截图目录VBS}"'
                "   ' 截图输出目录（后端按次注入运行目录）") + 文本[旧目录.end():]
        else:
            摘要["skipped"].append("未找到 CO28_SHOT_DIR 常量，截图目录未注入")

    autfp = _vbs清值(条件.get("autfp"))
    prof = _vbs清值(条件.get("prof"))
    if autfp:
        旧 = 'session.findById("wnd[0]/usr/ctxtP_AUTFP").text = "2"'
        if 旧 in 文本:
            文本 = 文本.replace(旧, f'session.findById("wnd[0]/usr/ctxtP_AUTFP").text = "{autfp}"', 1)
            摘要["autfp"] = autfp
    if prof:
        旧 = 'session.findById("wnd[0]/usr/ctxtP_PROFID").text = "000001"'
        if 旧 in 文本:
            文本 = 文本.replace(旧, f'session.findById("wnd[0]/usr/ctxtP_PROFID").text = "{prof}"', 1)
            摘要["prof"] = prof

    # 选择屏字段：生成 Dim 声明 + 赋值（插在执行按钮之前，保证先执行）+ Sub 定义（附在文件尾，VBScript 过程定义解析期提升）
    声明 = []
    赋值 = []
    调用 = []
    for 关键名, 参数键, 中文名 in CO28_FIELD_MAP:
        lo = _vbs清值(条件.get(参数键 + "Low"))
        hi = _vbs清值(条件.get(参数键 + "High"))
        if _vbs含批量语法(lo) or _vbs含批量语法(hi):
            摘要["skipped"].append(f"{中文名}（含批量语法，该条件不生效）")
            lo, hi = "", ""
        if lo or hi:
            摘要["fields"].append({"name": 中文名, "key": 关键名, "low": lo, "high": hi})
        声明.append(f"OV_{参数键}_LO, OV_{参数键}_HI")
        赋值.append(f'OV_{参数键}_LO = "{lo}" : OV_{参数键}_HI = "{hi}"')
        调用.append(f'    Co28SetSel sess, "{关键名}", OV_{参数键}_LO, OV_{参数键}_HI, "{中文名}"')

    头注释 = ("' ==== 动态条件（网页 CO28 页自动生成，勿手改） ====\n"
              f"' 生成时间：{datetime.now():%Y-%m-%d %H:%M:%S} · "
              f"P_AUTFP={摘要['autfp'] or '(原值)'} · P_PROFID={摘要['prof'] or '(原值)'}\n"
              "' 已按需求移除「变式搜索/选变式」段：筛选条件全部由页面直接写入选择屏\n")
    注入头 = (头注释
              + "Dim " + ", ".join(声明) + "\n"
              + "\n".join(赋值) + "\n"
              + "Co28ApplyOverrides session\n")

    尾注 = ("' ==== 动态条件子过程（网页条件自动生成，勿手改） ====\n"
            "Sub Co28ApplyOverrides(sess)\n"
            + "\n".join(调用) + "\n"
            + "End Sub\n"
            "\n"
            "' 按字段关键名在选择屏上找 xxx-LOW / xxx-HIGH 输入框并写入；找不到回显警告（变式段已移除，条件以页面为准）\n"
            "Sub Co28SetSel(sess, key, lo, hi, zh)\n"
            "    Dim usr, c, nm, hitLo, hitHi\n"
            "    If Len(Trim(lo)) = 0 And Len(Trim(hi)) = 0 Then Exit Sub\n"
            "    Set usr = sess.findById(\"wnd[0]/usr\")\n"
            "    hitLo = False : hitHi = False\n"
            "    On Error Resume Next\n"
            "    For Each c In usr.Children\n"
            "        nm = UCase(CStr(c.Id))\n"
            "        If InStr(nm, UCase(key)) > 0 Then\n"
            "            If Right(nm, 4) = \"-LOW\" And Len(Trim(lo)) > 0 Then\n"
            "                c.Text = lo : hitLo = True\n"
            "                WScript.Echo \"[ov] \" & zh & \" 下限 <- \" & lo\n"
            "            ElseIf Right(nm, 5) = \"-HIGH\" And Len(Trim(hi)) > 0 Then\n"
            "                c.Text = hi : hitHi = True\n"
            "                WScript.Echo \"[ov] \" & zh & \" 上限 <- \" & hi\n"
            "            End If\n"
            "        End If\n"
            "    Next\n"
            "    On Error GoTo 0\n"
            "    If Not hitLo And Len(Trim(lo)) > 0 Then WScript.Echo \"[ov] 警告：选择屏未找到 \" & key & \"-LOW，\" & zh & \" 该条件未生效\"\n"
            "    If Not hitHi And Len(Trim(hi)) > 0 Then WScript.Echo \"[ov] 警告：选择屏未找到 \" & key & \"-HIGH，\" & zh & \" 该条件未生效\"\n"
            "End Sub\n")

    # 整段移除「变式搜索 → 选变式」流程（到执行按钮之前），换成直接写选择屏字段。
    # 锚点用纯代码定位（不依赖中文注释，原 .vbs 注释乱码时也能命中）：
    #   起点 = 打开变式搜索那行 Co28PressBtn ... "btn[17]"（含其上方紧邻的注释行）
    #   终点 = 执行按钮 session.findById("wnd[0]/tbar[1]/btn[8]").press
    段起代码 = 'Co28PressBtn session, "wnd[0]/tbar[1]", "btn[17]"'
    锚点 = 'session.findById("wnd[0]/tbar[1]/btn[8]").press'
    i段 = 文本.find(段起代码)
    i锚 = 文本.find(锚点)
    i1 = -1
    if i段 > 0:
        i1 = 文本.rfind("\n", 0, i段) + 1
        # 往上吃掉紧邻的注释行（以单引号开头，且中间没有空行）
        while i1 > 0:
            上一行首 = 文本.rfind("\n", 0, i1 - 1) + 1
            if 上一行首 >= i1 - 1:
                break
            上一行 = 文本[上一行首:i1].strip()
            if 上一行.startswith("'") and 上一行:
                i1 = 上一行首
            else:
                break
    if 0 < i1 < i锚:
        文本 = 文本[:i1] + 注入头 + 文本[i锚:]
    else:
        摘要["skipped"].append(
            f"未找到变式段/执行锚点（btn[17]@{i段} / btn[8]@{i锚}），选择屏字段覆盖未注入（仅改 P_AUTFP/P_PROFID）")
    # 彻底清除变式残留：常量 VARIANT_NAME + 已不再调用的 Co28PickVariant 函数定义
    # （用户 2026-10-19 明令：运行时脚本里不许再有任何「选变式」痕迹）
    文本 = _re.sub(r"(?m)^[ \t]*Const[ \t]+VARIANT_NAME[ \t]*=.*\r?\n", "", 文本)
    文本 = _re.sub(r"(?ms)^[ \t]*Sub[ \t]+Co28PickVariant\b.*?^[ \t]*End Sub[ \t]*\r?\n", "", 文本)
    文本 = 文本 + "\n" + 尾注
    # 收尾（2026-10-23 用户需求）：执行完把 SAP GUI 右下角的「忙/运行中」指示立即复位。
    # closePopups=False —— CO28 是写入类（下达）流程，绝不在收尾时乱关确认框。
    文本 = 文本 + "\n" + VBS_CLEAN_SUB + "\nEndClean False\n"
    # 「无数据行」提前退出那条路也要收尾（否则 GUI 一样挂着忙指示）
    # 注意 \r?：VBS 是 CRLF 文件，$ 在 (?m) 下停 \n 之前，不吞 \r 就永远匹配不上（2026-10-23 踩过）
    文本 = _re.sub(r"(?m)^([ \t]*)WScript\.Quit 0[ \t]*\r?$",
                   r"\1EndClean False\r\n\1WScript.Quit 0", 文本)

    运行时 = VBS_BASE / "co28_runtime.vbs"
    运行时.write_bytes(文本.encode("gbk", errors="replace"))
    return 运行时, 摘要


def _co28缩放图片(p: "Path", 最大宽: int = 1080, 最大高: int = 1500) -> "Path":
    """把截图等比缩到适合邮件正文阅读的尺寸（原地覆盖）。

    用户 2026-10-23：「邮件里面可以缩小一点，能看清楚就行」——dws 内联图按**原始像素**渲染，
    1936x1056 的 SAP 全屏图在邮箱正文区只露得出一部分（右侧/下方被裁），所以发信前统一缩宽。
    """
    try:
        from PIL import Image
        with Image.open(p) as im:
            宽, 高 = im.size
            比 = min(最大宽 / 宽 if 宽 > 最大宽 else 1.0,
                     最大高 / 高 if 高 > 最大高 else 1.0)
            if 比 >= 1.0:
                return p
            _ = im.convert("RGB").resize((max(1, int(宽 * 比)), max(1, int(高 * 比))),
                                         Image.LANCZOS)
            _.save(p, "PNG", optimize=True)
    except Exception as exc:
        UI错误(f"CO28 截图缩放失败：{exc}")
    return p


def _co28暂存邮件图片(paths: list) -> list:
    """把待发邮件用的截图复制到独立暂存目录，返回副本路径列表。

    2026-10-31：失败邮件的内联截图与 VBS 运行截图同在 co28_shots/ 下，新运行开跑
    （或 VBS Co28ShotAllScreens）会清目录 —— 上一封邮件还在 dws 重试时文件就被删了。
    发信前先拷一份到 mail_stage/<时间戳>/，此后无论 co28_shots 怎么清，邮件都能发出。
    复制失败的项跳过（返回里不带），宁可少图也不让整封邮件挂掉。
    """
    结果 = []
    try:
        暂存根 = VBS_LOG_DIR / "mail_stage"
        暂存目录 = 暂存根 / f"{datetime.now():%Y%m%d_%H%M%S_%f}"
        暂存目录.mkdir(parents=True, exist_ok=True)
        # 顺手清掉 7 天前的旧暂存目录，防长期堆积
        try:
            截止 = datetime.now().timestamp() - 7 * 86400
            for 旧目录 in 暂存根.glob("*"):
                if 旧目录.is_dir() and 旧目录.stat().st_mtime < 截止:
                    shutil.rmtree(str(旧目录), ignore_errors=True)
        except Exception:
            pass
        for p in paths or []:
            try:
                p = Path(p)
                if not p.is_file():
                    continue
                目标 = 暂存目录 / p.name
                shutil.copy2(str(p), str(目标))
                结果.append(目标)
            except Exception as exc:
                UI错误(f"CO28 邮件截图暂存失败（跳过 {p}）：{exc}")
    except Exception as exc:
        UI错误(f"CO28 邮件截图暂存目录创建失败（退回原路径）：{exc}")
        return list(paths or [])
    return 结果 or list(paths or [])


def _co28收集截图() -> list:
    """收集本次运行的截图（co28_shots 目录，按文件名排序，过滤过小文件与失败邮件专用图）。

    2026-10-28 用户：「第 1 屏 · 执行结果屏（含左下角状态栏）这一屏不要，结果已经进邮件标题」
    → VBS 已不再生成 co28_shot_00_result.png；这里再兜一层，历史残留的 _result 图也不发。
    """
    try:
        return sorted([p for p in CO28_SHOTS_DIR.glob("co28_shot_*.png")
                       if p.stat().st_size > 2000 and "_result" not in p.name],
                      key=lambda x: x.name)
    except Exception:
        return []


def _co28抓当前SAP屏(前缀: str = "co28_fail_sap") -> "Path | None":
    """失败时抓当前 SAP 窗口一屏（调 sap_shot.py 子进程）；失败返回 None。"""
    try:
        CO28_SHOTS_DIR.mkdir(parents=True, exist_ok=True)
        目标 = CO28_SHOTS_DIR / f"{前缀}_{datetime.now():%H%M%S}.png"
        助手 = APP_DIR / "sap_shot.py"
        pyw = Path(r"C:\Python311\pythonw.exe")
        解释器 = str(pyw) if pyw.is_file() else sys.executable
        subprocess.run([解释器, str(助手), str(目标)],
                       capture_output=True, timeout=40,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return 目标 if 目标.is_file() and 目标.stat().st_size > 2000 else None
    except Exception as exc:
        UI错误(f"CO28 失败截图失败：{exc}")
        return None


def _co28存网页截图(数据: str, 前缀: str = "co28_fail_web") -> "Path | None":
    """前端 html2canvas 的 data:image/png;base64,... → 落盘 PNG；失败返回 None。"""
    try:
        原始 = str(数据 or "").strip()
        if not 原始:
            return None
        if "," in 原始 and 原始.lower().startswith("data:"):
            原始 = 原始.split(",", 1)[1]
        import base64
        字节 = base64.b64decode(原始)
        if len(字节) < 2000:
            return None
        CO28_SHOTS_DIR.mkdir(parents=True, exist_ok=True)
        目标 = CO28_SHOTS_DIR / f"{前缀}_{datetime.now():%H%M%S}.png"
        目标.write_bytes(字节)
        return 目标
    except Exception as exc:
        UI错误(f"CO28 网页截图落盘失败：{exc}")
        return None


def _co28日志尾部(n: int = 45) -> str:
    """本次运行日志尾部（vbs_logs 最新 co28 日志；读不到返回空串）。"""
    try:
        日志 = sorted(VBS_LOG_DIR.glob("*_co28.log"),
                     key=lambda x: x.stat().st_mtime, reverse=True)
        if 日志:
            行 = 日志[0].read_text(encoding="utf-8", errors="replace").splitlines()
            return "\n".join(行[-n:])
    except Exception:
        pass
    return ""


def _co28取执行结果(输出: str) -> "tuple[str, str]":
    """从本次运行输出里取 SAP 状态栏执行结果、勾选汇总（取不到返回空串）。

    用户 2026-10-23 19:17（附 SAP 截图红圈）：「截图标记的执行的结果，你要带上你发给我的邮件截图，
    里面没有这么个结果」——邮件必须带上左下角那句「已执行批量处理 - 已创建2日志(0/0/2)」。

    VBS v3.2 会在 P-- 覆盖状态栏之前打印 `[co28][result] 状态栏：xxx`，
    这里优先用本次 stdout，取不到再回退最新日志尾部。
    """
    结果 = ""
    汇总 = ""
    for 源 in (输出 or "", _co28日志尾部(120)):
        for 行 in 源.splitlines():
            行 = 行.strip()
            if not 结果 and 行.startswith("[co28][result]"):
                值 = 行.split("：", 1)[-1].strip()
                if 值 and "未取到" not in 值:
                    结果 = 值
            elif not 汇总 and 行.startswith("[co28] 全选+反选完成"):
                汇总 = 行[len("[co28] "):].strip()
        if 结果 and 汇总:
            break
    return 结果, 汇总


def _co28发送结果邮件(正常: bool, 被中断: bool, 输出: str, 开始: datetime,
                     收件人: str = "") -> dict:
    """CO28 结果邮件分发（2026-10-23 从「执行[且发邮件]」拆出的独立发邮件功能）。

    成功 → 收集 VBS 下达后跨滚动条多屏截图，发「下达成功」邮件（正文按屏引用内联图）。
    失败 → 不在此处发送（缺「当前网站的页面」截图），交回前端经 /api/vbs/co28_failmail 合并发送。

    2026-10-23 修复：VBS 在「无数据行」时会 `WScript.Quit 0` 提前收尾，cscript 仍返回 0、
    输出里也没有「错误」二字 → 此前被误判为「下达成功」并发出成功邮件（实际一行都没下达）。
    现在显式识别无数据行，交给失败邮件链路处理。
    """
    无数据 = bool(_re.search(r"无数据行|未找到数据行|无数据", 输出 or ""))
    if 无数据:
        return {"已发": False, "待失败邮件": True, "无数据": True,
                "说明": "当前筛选条件下无数据行，未执行下达；由前端回传网页截图后发送失败邮件"}
    if not 正常:
        return {"已发": False, "待失败邮件": True,
                "说明": f"{'被强行中断' if 被中断 else '执行异常'}：由前端回传网页截图后发送失败邮件"}
    dws = _rpa导dws()
    if dws is None:
        return {"已发": False, "说明": "dws_mail 不可用（web_auto/公共_钉钉通知 是否完整）"}
    # 2026-10-31：暂存副本再发，防 dws 重试期间 co28_shots 被并发运行清空
    截图 = _co28暂存邮件图片([_co28缩放图片(p) for p in _co28收集截图()])
    状态栏结果, 勾选汇总 = _co28取执行结果(输出)
    行 = ["CO28 下达完成。", ""]
    if 状态栏结果:
        行.append(f"执行结果（SAP 左下角状态栏）：{状态栏结果}")
    else:
        行.append("执行结果（SAP 左下角状态栏）：未取到（下达后状态栏可能已无提示）")
    if 勾选汇总:
        行.append(f"勾选汇总：{勾选汇总}")
    行.append(f"执行时间：{开始:%Y-%m-%d %H:%M:%S}")
    行.append(f"结果总览截图：共 {len(截图)} 屏（按垂直滚动条逐屏截取，含勾选列）")
    if not 截图:
        行.append("（本次未取到截图：mail 模式下 SAP 下达结果总览是否正常渲染？）")
    for i, p in enumerate(截图, 1):
        行.append("")
        行.append(f"第 {i} 屏 [inline:{p.name}]")
    正文 = "\n".join(行)
    收件 = _sapq解析收件人(收件人)
    收件人文本 = ", ".join(收件) if 收件 else ""
    # 2026-10-28 用户：「已创建796日志(0 / 0 / 0 / 796) 放在邮件标题里面」——
    # 状态栏结果随主题一起发，收件人在列表里不点开正文也能一眼看到下达结果。
    主题 = CO28_MAIL_主题成功
    结果短 = " ".join((状态栏结果 or "").split())
    if 结果短:
        if len(结果短) > 60:
            结果短 = 结果短[:60] + "…"
        主题 += "｜" + 结果短
    try:
        kwargs = {"内联图片": 截图}
        if 收件人文本:
            kwargs["收件人"] = 收件人文本
        状态, 说明 = dws.发送邮件(主题, 正文, **kwargs)
    except Exception as exc:
        状态, 说明 = "彻底失败", f"发送异常：{exc}"
    UI信息(f"CO28 成功邮件：{状态}｜{说明}｜截图 {len(截图)} 张｜收件人 {收件人文本 or '（默认）'}"
           f"｜主题：{主题}")
    return {"已发": 状态 == "已发送", "状态": 状态, "说明": 说明, "主题": 主题,
            "截图数": len(截图), "收件人": 收件人文本 or "demo.user64@example.com（默认）"}


def _查询页条件文字(条件: dict) -> list:
    """把本次注入的条件按页签整理成邮件正文的「查询条件」段（只列填了值的项）。

    键名来自 VBS 行尾标记，形如 zb03MatLow / zb03MatHigh / listBre / zb02ChangedAtLow 等。
    """
    if not 条件:
        return []
    显示名 = {
        # Zb02（ZLSTMA）
        "zb02Mat": "物料", "zb02Pir": "品目", "zb02PartName": "部件名称",
        "zb02DrawVar": "图号/变量", "zb02BrandGrp": "品/群", "zb02Lno": "L 番",
        "zb02Erdat": "创建日期", "zb02Ernam": "创建者", "zb02MatTrFlag": "物料传输标记",
        "zb02BomTrFlag": "BOM 传输标志", "zb02Fertw": "生产科使用",
        "zb02Variant": "变量", "zb02MatLongParam": "物料号长参数表",
        "zb02TechEff": "工艺生效日期", "zb02ChangedAt": "Changed at",
        # Zb03（ZLSTBM）
        "zb03Mat": "物料", "zb03Posnr": "项目号", "zb03Comp": "组件", "zb03Menge": "数量",
        "zb03MatSupply": "物料供应标识", "zb03FactoryState": "出厂状态的标记",
        "zb03FactoryBox": "出厂箱号", "zb03BoxQty": "箱数", "zb03Fertw": "生产科使用",
        "zb03BoxNo": "箱号", "zb03Mat2": "物料(第2个)", "zb03Pir": "品目",
        "zb03BrandGrp": "品/群", "zb03Lno": "L 番", "zb03Variant": "变量",
        "zb03Erdat": "创建日期", "zb03Ernam": "创建者", "zb03Mfbom": "制造BOM标记",
        "zb03ListBre": "输出列表的宽度", "zb03MaxSel": "最大命中数量",
        # Zpp_Po（ZPP_PO）
        "delflagA": "删除标志A", "delflagB": "删除标志B", "fevor": "生产调度员",
        "aufnr": "订单", "plnum": "计划订单", "posnr": "订单项目编号",
        "matnr": "物料编号", "matkl": "物料组", "dispo": "MRP控制者",
        "stlan": "BOM 展开编号", "pstart": "计划开始日期", "pdelv": "计划订单交货日",
        "gstrp": "基本开始日期", "gltrp": "基本完成日期", "freiz": "实际下达",
        "fstrt": "实际开始", "fenden": "实际完成", "ladat": "实际交货日期",
        "wzeit": "收货处理时间", "lgort": "库存地点", "tecko": "技术完成",
        "lifsk": "交货已完成", "psmng": "项目数量", "wemng": "GR 数量",
        "meins": "基本单位", "quan": "确认数量", "ueber": "收货的过量/不足",
        "maktx": "描述", "spras": "语言", "plnnum": "计划工艺路线号",
        "plnnr": "生产/检验备忘录", "ltxsp": "长文本存在", "erdat": "创建日期",
        "listBre": "输出列表的宽度", "maxSel": "最大命中数量",
        # Zpp_Pur（ZPP_PUR，SE16 选择屏）
        "zpurEbeln": "采购凭证", "zpurEbelp": "项目", "zpurLoekz": "删除标识",
        "zpurPstyp": "项目类别", "zpurLifnr": "供应商", "zpurMatkl": "物料组",
        "zpurDispo": "MRP 控制员", "zpurEkgrp": "采购组", "zpurMatnr": "物料",
        "zpurEmatn": "物料(EMATN)", "zpurNetpr": "净价", "zpurEkorg": "采购组织",
        "zpurWerks": "工厂", "zpurLgort": "库存地点", "zpurBednr": "跟踪编号",
        "zpurInfnr": "信息记录", "zpurBedat": "采购订单日期", "zpurMenge": "采购订单数量",
        "zpurMeins": "订单单位", "zpurEindt": "交货日期", "zpurWemng": "交货",
        "zpurElikz": "交货已完成", "zpurSernr": "BOM 展开编号", "zpurEtenr": "计划行",
        "zpurDrdat": "上一次传输", "zpurLewed": "最近的收货日期", "zpurEvers": "装运介绍",
        "zpurMaktx": "描述", "zpurSpras": "语言", "zpurListBre": "输出列表的宽度",
        "zpurMaxSel": "最大命中数量",
        # Zpp_Cs12（/nzpp_cs12 事务码；探针实测后按实测标签自动更新此处）
        "zppcs12I1": "物料", "zppcs12I2": "工厂", "zppcs12I3": "BOM 用途",
        "zppcs12I4": "BOM 类别", "zppcs12I5": "有效起始日", "zppcs12I6": "更改编号",
        "zppcs12I7": "项目", "zppcs12I8": "组件", "zppcs12I9": "物料组",
        "zppcs12I10": "创建者", "zppcs12I11": "创建日期", "zppcs12I12": "更改日期",
    }
    单值键 = {"zb03ListBre", "zb03MaxSel", "listBre", "maxSel", "zpurListBre", "zpurMaxSel"}
    行 = []
    已列 = set()
    for k, v in 条件.items():
        if k in 单值键:
            if str(v or "").strip():
                行.append(f"  · {显示名.get(k, k)} = {v}")
            已列.add(k)
            continue
        if k.endswith("Low"):
            基 = k[:-3]
            高 = str(条件.get(基 + "High", "") or "").strip()
            低 = str(v or "").strip()
            if not 低 and not 高:
                continue
            名 = 显示名.get(基, 基)
            行.append(f"  · {名} = {低 + ('..' + 高 if 高 else '')}")
            已列.add(基 + "High")
    行.sort()
    return 行


def _发送查询结果邮件(脚本键: str, 正常: bool, 被中断: bool, 输出: str,
                     开始: datetime, 条件: dict,
                     收件人: str = "", 附件: Optional[List[Path]] = None,
                     下载目录: "Path | str | None" = None, 报表回落: bool = True,
                     正文表格: "Optional[List[List[str]]]" = None, 附截图: bool = True) -> dict:
    """Zb02 / Zb03 / Zpp_Po / Zpp_Pur 的结果邮件（2026-10-23 从「执行[且发邮件]」拆出的独立发邮件功能）。

    正文 = 执行结果（SAP 左下角状态栏）+ 查询条件（本次注入的页面现值）
         + 执行时间 + 逐屏内联截图（第 1 屏为执行结果屏，左下角保留命中行数）
         + 报表附件（下载目录里最新的「功能名_时间戳.xls」作为普通附件随邮件发出）。

    报表附件口径（2026-10-23 晚修正 —— 用户反馈「明明下载开关开着却显示开关未开」）：
        ① 本次执行刚下载的那份（附件入参）优先；
        ② 没有就回落「下载目录里最新的一份」（于是"上次下载的、这次只点发邮件"也能带上）；
           注意：② 只在「本次执行成功」时生效；执行异常 / 被中断时绝不拿旧文件冒充本次结果。
        ③ 两份都没有才写「无」，并且正文里如实写清**实际原因与目录**，不再拿开关状态当借口。

    收件人 = 页面「收件人」框（可多个，逗号/分号/空格分隔）；留空回落 dws 默认 demo.user65@example.com。
    下载目录 = 页面「下载路径」框（自定义优先，留空 = 代码默认 C:/Users/u3854/Desktop/web_xls）。
    注意：本函数里另有一个局部变量 目录 = 截图目录，所以参数特意叫 下载目录，避免被覆盖。
    """
    _, 显示名, 事务码, _, _ = VBS_SCRIPTS_SITE[脚本键]
    dws = _rpa导dws()
    if dws is None:
        return {"已发": False, "说明": "dws_mail 不可用（web_auto/公共_钉钉通知 是否完整）"}

    目录 = _vbs站点截图目录(脚本键)
    前缀 = f"{脚本键}_shot_"
    截图 = []
    if 附截图:
        try:
            截图 = sorted([p for p in 目录.glob(f"{前缀}*.png") if p.stat().st_size > 2000],
                         key=lambda x: x.name)
        except Exception:
            截图 = []
        # 2026-10-31：暂存副本再发，防 dws 重试期间截图目录被并发运行清空
        截图 = _co28暂存邮件图片([_co28缩放图片(p) for p in 截图])

    状态栏 = ""
    for 行_ in (输出 or "").splitlines():
        行_ = 行_.strip()
        if 行_.startswith("[result] 状态栏："):
            值 = 行_.split("：", 1)[-1].strip()
            if 值 and "无提示" not in 值:
                状态栏 = 值
                break

    行 = [f"{显示名} 查询完成（事务码 {事务码}）。", ""]
    if 被中断:
        行.append("执行状态：被用户强行中断（结果不完整）")
    elif not 正常:
        行.append("执行状态：执行异常（详见网站「自动化任务 → 运行日志」）")
    行.append(f"执行结果（SAP 左下角状态栏）：{状态栏 or '未取到（请看第 1 屏截图左下角）'}")
    条件行 = _查询页条件文字(条件 or {})
    行.append("")
    行.append(f"查询条件（网页输入现值，共 {len(条件行)} 项）：" if 条件行 else "查询条件：（全部为空）")
    行.extend(条件行)
    行.append("")
    行.append(f"执行时间：{开始:%Y-%m-%d %H:%M:%S}")
    # 2026-11-06 用户需求：Zpp_Pr2po 正文改为「按条件筛选的纯文本表格」（替代旧的制表符堆砌 blob）
    #   筛选条件：①物料号 P125 开头且供应商 201253；②物料号 P183 开头且供应商 201007；③供应商属「停产供方」。
    #   邮件通道为纯文本，这里用 │ 分隔 + 表头分隔线，字段对齐、一眼能看懂；真正的 Excel 样式在附件 xlsx。
    if 正文表格:
        行.append("")
        行.append("—— 需重点关注的数据（筛选条件：①物料号 P125 开头且供应商 201253；"
                  "②物料号 P183 开头且供应商 201007；③供应商属于「停产供方」）——")
        行.extend(_sapq正文表格文本(正文表格))
        行.append("")
        行.append("（完整报表见附件 xlsx：已删除「已转换数量不为0，无法全量转换」等提示行、列宽自适应、表头冻结）")
    if 附截图:
        行.append("")
        行.append(f"结果截图：共 {len(截图)} 屏（第 1 屏为执行完成瞬间，左下角保留命中行数；其余按滚动条逐屏截取）")
        if not 截图:
            行.append("（本次未取到截图：请确认 SAP 已最大化、截图助手 sap_shot.py 可用）")
        else:
            # 2026-10-24 用户三次反馈「执行完 GUI 脚本还处于运行状态」——实为截图右下角那条红白条纹图标
            #   = SAP GUI 的「脚本正在运行」指示；这条截图是脚本自己在运行中拍的，必然亮着。收尾与否看脚本日志。
            行.append("说明：截图右下角的红白条纹小图标 = SAP GUI 的「脚本正在运行」指示，"
                      "脚本自己给自己拍照时必然亮着，不代表脚本没停；收尾是否完成看网站日志里的"
                      "「[clean] 收尾完成」一行。")
        for i, p in enumerate(截图, 1):
            行.append("")
            if p.name.endswith("_00_result.png"):
                行.append(f"第 {i} 屏 · 执行结果屏（含左下角状态栏）[inline:{p.name}]")
            else:
                行.append(f"第 {i} 屏 [inline:{p.name}]")
    报表附件 = [Path(p) for p in (附件 or []) if p and Path(p).is_file()]
    附件来历 = "本次下载" if 报表附件 else ""
    # 回落「下载目录里最新的一份」只在「本次执行成功」时才允许：
    # 执行异常 / 被中断时绝不再拿目录里旧文件冒充本次结果（避免 2026-10-27 失败的 Zb03 还挂上午的报表）。
    # 「只点发邮件」入口永远传 正常=True，所以那个便利依然保留。
    if not 报表附件 and 报表回落 and 正常 and not 被中断:
        最新 = _sapq取最新报表(脚本键, 下载目录)
        if 最新:
            报表附件 = [最新]
            # 来历写清「哪個目录里最新的」，免得用户以为是本次下载的
            附件来历 = f"{最新.parent} 里最新的一份"
    实际目录 = str(_sapq确保下载目录(下载目录))
    前缀 = SAPQ_报表前缀.get(脚本键) or 脚本键
    行.append("")
    if 报表附件:
        行.append("报表附件：" + "、".join(p.name for p in 报表附件)
                 + f"（{附件来历}，已随本邮件附带）")
    else:
        行.append(f"报表附件：无 —— {实际目录} 里还没有 {前缀}_*{SAPQ_报表后缀} 报表。"
                 "要带上报表：把「报表下载」开关打开再点「执行查询」。"
                 "（2026-10-24 用户要求删除「⬇ 报表下载」按钮：下载/发邮件统一由开关 + 执行查询驱动）")
    正文 = "\n".join(行)
    主题 = f"{显示名} 查询结果【{'成功' if 正常 and not 被中断 else '异常'}】{开始:%Y-%m-%d %H:%M}"
    收件 = _sapq解析收件人(收件人)
    收件人文本 = ", ".join(收件) if 收件 else ""
    try:
        kwargs = {"内联图片": 截图}
        if 收件人文本:
            kwargs["收件人"] = 收件人文本
        if 报表附件:
            kwargs["普通附件"] = 报表附件
        状态, 说明 = dws.发送邮件(主题, 正文, **kwargs)
    except Exception as exc:
        状态, 说明 = "彻底失败", f"发送异常：{exc}"
    UI信息(f"{脚本键} 结果邮件：{状态}｜{说明}｜截图 {len(截图)} 张"
           f"｜报表附件 {len(报表附件)} 个｜收件人 {收件人文本 or '（默认）'}")
    return {"已发": 状态 == "已发送", "状态": 状态, "说明": 说明,
            "截图数": len(截图), "报表附件": [p.name for p in 报表附件],
            "收件人": 收件人文本 or "demo.user66@example.com（默认）"}


def _发送co28结果钉钉(正常: bool, 被中断: bool, 输出: str, 开始, 收件人: str = "") -> dict:
    """CO28 结果摘要发钉钉单聊（并把「执行结果屏」截图作为文件消息补发一条）。

    口径与 _co28发送结果邮件 完全一致（同一份状态栏 / 勾选汇总 / co28_shots 截图目录），
    只是通道从邮件换成钉钉单聊 —— CO28 与各查询页行为对齐，用户不必记两套。
    """
    模块 = _rpa导dws_msg()
    if 模块 is None:
        return {"已发": False, "状态": "失败", "说明": "dws_msg 不可用（web_auto/公共_钉钉通知 是否完整）"}
    无数据 = bool(_re.search(r"无数据行|未找到数据行|无数据", 输出 or ""))
    状态栏结果, 勾选汇总 = _co28取执行结果(输出)
    截图 = _co28收集截图()
    行 = ["【CO28 批量下达】结果通知（事务码 CO28）"]
    行.append("执行状态：" + ("无数据行（未执行下达）" if 无数据
                          else ("被用户强行中断（结果不完整）" if 被中断
                                else ("成功" if 正常 else "执行异常（详见网站「自动化任务 → 运行日志」）"))))
    行.append(f"执行结果（SAP 左下角状态栏）：{状态栏结果 or '未取到（请看截图左下角）'}")
    if 勾选汇总:
        行.append(f"勾选汇总：{勾选汇总}")
    行.append(f"执行时间：{开始:%Y-%m-%d %H:%M:%S}　耗时：{round((datetime.now() - 开始).total_seconds(), 1)}s"
              f"　结果截图：{len(截图)} 屏")
    行.append(f"随附截图：{截图[0].name}（另发一条文件消息）" if 截图 else "（本次未取到截图）")
    正文 = chr(10).join(行)

    产出 = {"已发": False, "状态": "失败", "说明": "", "正文": 正文,
            "收件人": 收件人 or "演示用户（当前登录用户本人）", "报表附件": [], "附件名": []}
    try:
        状态, 说明 = 模块.发送文本消息(正文, 收件人)
    except Exception as exc:
        状态, 说明 = "失败", f"发送异常：{exc}"
    产出["状态"], 产出["说明"] = 状态, 说明
    产出["已发"] = 状态 in ("已发送", "部分成功")
    if 产出["已发"] and 截图:
        try:
            文件状态, 文件说明 = 模块.发送文件消息(截图[0], "", 收件人)
        except Exception as exc:
            文件状态, 文件说明 = "失败", f"发送异常：{exc}"
        产出["文件消息"] = {"状态": 文件状态, "说明": 文件说明}
        if 文件状态 in ("已发送", "部分成功"):
            产出["附件名"] = [截图[0].name]
            产出["报表附件"] = [str(截图[0])]
    UI信息(f"CO28 钉钉消息：{状态}｜{说明}｜截图附件 {len(产出['附件名'])} 个｜收件人 {产出['收件人']}")
    return 产出


def _是伪装的文本报表(路径) -> bool:
    """判断一个 .xls 其实是不是「Tab 分隔纯文本」。

    🔴 背景（2026-10-24 用户现场截图）：SAP 报表下载下来是 GBK 编码的 Tab 分隔文本，
    只是文件名被起成了 .xls。用 os.startfile 交给 Excel 时，Excel 会弹
    「Zpp_Cs12_xxx.xls 的文件格式和扩展名不匹配。文件可能已损坏或不安全…
     除非您信任其来源，否则请勿打开。是否仍要打开它？」——必须点「是」才能看。
    COM 直接解析不会弹这个框（已实测）。"""
    try:
        with open(路径, "rb") as f:
            头 = f.read(8)
    except Exception:
        return False
    # 真 xls = OLE2 复合文档；真 xlsx/xlsm = ZIP(PK)；其余一律按文本处理
    return not (头[:4] == b"\xd0\xcf\x11\xe0" or 头[:2] == b"PK")


def _关闭Excel里已打开的文件(路径) -> List[str]:
    """把 Excel 里打开着的这个文件关掉（不保存改动），用于「转成 xlsx 后同步关闭老的 xls」。

    2026-10-24 用户明令：「如果转化为xlsx，老的 xls 文件同步关闭」。
    只按完整路径**精确匹配**，绝不碰用户其它工作簿；找不到 Excel / 文件没打开 → 静默返回空列表。
    """
    目标 = ""
    try:
        目标 = str(Path(路径).resolve()).lower()
    except Exception:
        return []
    try:
        import win32com.client as wc
    except Exception:
        return []
    try:
        app = wc.GetObject(Class="Excel.Application")
    except Exception:
        return []
    簿列表 = []
    try:
        for 簿 in app.Workbooks:
            try:
                簿列表.append(簿)
            except Exception:
                continue
    except Exception:
        return []
    已关 = []
    for 簿 in 簿列表:
        try:
            名 = str(簿.Name)
            全名 = str(簿.FullName or "").lower()
        except Exception:
            continue
        if 全名 != 目标:
            continue
        try:
            簿.Close(SaveChanges=False)      # False = 不弹「是否保存」，直接关，不动原文件内容
            已关.append(名)
        except Exception:
            pass
    return 已关


def _读剪贴板文本() -> str:
    """读 Windows 剪贴板里的文本（SAP 清单走「在剪贴板中」导出时在这里取）。

    2026-10-25：SAP「本地文件 → 电子表格」导 85652 行要 ~38 秒（含 ALV→清单转换），
    而「在剪贴板中」只要 ~16 秒且不经过保存对话框 → 大数据量改走剪贴板，这里负责落盘。
    读不到（剪贴板被占/被清）就返回空串，由调用方如实报错，绝不静默当成功。
    """
    try:
        import win32clipboard as cb
    except Exception:
        return ""
    try:
        cb.OpenClipboard()
        try:
            return cb.GetClipboardData(cb.CF_UNICODETEXT) or ""
        finally:
            cb.CloseClipboard()
    except Exception:
        return ""


def _sapq剪贴板转TSV(文本: str) -> str:
    """SAP 清单剪贴板文本（定宽 + `|` 竖线 + `----` 分隔线）→ Tab 分隔文本。

    产出与「本地文件 → 电子表格」同一口径：GBK + Tab + CRLF，Excel 打开列对齐、不弹格式警告。
    规则：以 `|` 开头的行按竖线切列并 rstrip；纯分隔线丢弃；其余行（日期/标题等）原样保留。
    """
    行们 = []
    for 行 in (文本 or "").splitlines():
        s = 行.rstrip()
        if not s.strip():
            行们.append("")
            continue
        if set(s.strip()) == {"-"}:
            continue
        if "|" in s:
            段 = s.split("|")
            if 段 and not 段[0].strip():
                段 = 段[1:]
            if 段 and not 段[-1].strip():
                段 = 段[:-1]
            行们.append("\t".join(p.rstrip() for p in 段))
        else:
            行们.append(s)
    return "\r\n".join(行们)


def _xls转xlsx(xls路径, 目标目录=None) -> dict:
    """用 xlwings 把 .xls 另存为同名 .xlsx（Excel 另存 FileFormat=51 = xlOpenXMLWorkbook）。

    2026-10-24 用户需求：「报表转换为xlsx … 将报表下载对应的 xls 文件自动转换为同名的 xlsx 文件，
    如果用 Python 代码的话一定要使用 xlwings 库」——本实现即按要求用 xlwings。
    只读取并另存，不动原 .xls；任何失败都如实返回，绝不抛异常。
    """
    源 = Path(xls路径)
    目录 = Path(目标目录) if 目标目录 else 源.parent
    目标 = 目录 / (源.stem + ".xlsx")
    if not 源.is_file():
        return {"ok": False, "说明": f"找不到要转换的文件：{源}", "源": str(源)}
    if 源.suffix.lower() == ".xlsx":
        return {"ok": True, "文件": 源.name, "路径": str(源), "说明": "源文件已是 xlsx，无需转换"}
    try:
        import xlwings as xw
    except Exception as exc:
        return {"ok": False, "说明": f"未安装 xlwings（pip install xlwings）：{exc}", "源": str(源)}
    app = None
    try:
        目录.mkdir(parents=True, exist_ok=True)
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        # 🔴 SAP 报表其实是 GBK+Tab 文本却叫 .xls（见 _是伪装的文本报表）。xlwings 的
        #    books.open 不给 Format 时会走「文本导入向导」，弹 3 步对话框把流程卡死。
        #    这里显式声明 Tab 分隔（1=xlDelimited）直接进表；真 xls/xlsx 也照此打开无副作用。
        簿 = app.books.open(str(源), format=1 if _是伪装的文本报表(源) else None)
        try:
            簿.api.SaveAs(str(目标), FileFormat=51)
        finally:
            try:
                簿.close()
            except Exception:
                pass
        if 目标.is_file():
            # 用户 2026-10-24：「如果转化为xlsx，老的 xls 文件同步关闭」——
            # 转换完把还开在 Excel 里的原 .xls 关掉，别留一个旧窗口在那儿。
            已关 = _关闭Excel里已打开的文件(源)
            return {"ok": True, "文件": 目标.name, "路径": str(目标), "源": str(源),
                    "关闭原文件": 已关,
                    "说明": f"{源.name} → {目标.name}" + (f"（已关闭旧的 {源.name}）" if 已关 else "")}
        return {"ok": False, "说明": f"另存为失败：{目标} 未生成", "源": str(源)}
    except Exception as exc:
        return {"ok": False, "说明": f"转换异常：{exc}", "源": str(源)}
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def _发送查询结果钉钉(脚本键: str, 正常: bool, 被中断: bool, 输出: str, 开始,
                      条件: dict, 收件人: str = "", 附件: list | None = None) -> dict:
    """把本次查询结果摘要发到钉钉单聊（dws chat +messages-send --as user）。

    2026-10-24 用户需求：「发邮件后面再加一个钉钉消息/群通知…输入名字的话自动发送到对应聊天窗口…
    如果输入框以及报表下载都打开状态，我希望发给演示用户附带报表」。
    与结果邮件同一口径（同一份状态栏/条件/报表），只是通道换成钉钉单聊：
      ① 先发一条文本摘要（执行状态 / SAP 状态栏 / 条件 / 耗时 / 截图数）；
      ② 报表开关同时打开且报表已下载 → 再补一条文件消息把 xls 发过去
         （dws 的 text 与 file 互斥，只能分两条发）。
    收件人：留空或「演示用户/演示用户本人」= 当前钉钉登录用户本人（组织内 3 个重名，按本人最稳）；
            纯数字 = userId；其他 = 姓名（走 --user-query，重名会返回候选并判失败）。
    任何失败都不抛异常，返回 dict 供前端提示与排查。
    """
    _, 显示名, 事务码, _, _ = VBS_SCRIPTS_SITE[脚本键]
    模块 = _rpa导dws_msg()
    if 模块 is None:
        return {"已发": False, "状态": "失败", "说明": "dws_msg 不可用（web_auto/公共_钉钉通知 是否完整）"}

    状态栏 = ""
    for 行_ in (输出 or "").splitlines():
        行_ = 行_.strip()
        if 行_.startswith("[result] 状态栏："):
            值 = 行_.split("：", 1)[-1].strip()
            if 值 and "无提示" not in 值:
                状态栏 = 值
                break

    截图数 = 0
    try:
        截图数 = len([p for p in _vbs站点截图目录(脚本键).glob(f"{脚本键}_shot_*.png") if p.stat().st_size > 2000])
    except Exception:
        截图数 = 0

    报表 = [p for p in (附件 or []) if getattr(p, "is_file", None) and p.is_file()]
    行 = [f"【{显示名}】查询完成（事务码 {事务码}）"]
    行.append("执行状态：" + ("被用户强行中断（结果不完整）" if 被中断
                          else ("成功" if 正常 else "执行异常（详见网站「自动化任务 → 运行日志」）")))
    行.append(f"执行结果（SAP 左下角状态栏）：{状态栏 or '未取到（请看第 1 屏截图左下角）'}")
    条件行 = _查询页条件文字(条件 or {})
    行.append(f"查询条件（网页输入现值，共 {len(条件行)} 项）：" if 条件行 else "查询条件：（全部为空）")
    行.extend(条件行)
    行.append(f"执行时间：{开始:%Y-%m-%d %H:%M:%S}　耗时：{round((datetime.now() - 开始).total_seconds(), 1)}s"
              f"　结果截图：{截图数} 屏")
    行.append(f"报表附件：{报表[0].name}（随后另发一条文件消息）" if 报表
              else "报表附件：无 —— 要带上报表：把「报表下载」开关打开再点「执行查询」")
    正文 = chr(10).join(行)

    产出 = {"已发": False, "状态": "失败", "说明": "", "正文": 正文,
            "收件人": 收件人 or "演示用户（当前登录用户本人）", "报表附件": []}
    try:
        状态, 说明 = 模块.发送文本消息(正文, 收件人)
    except Exception as exc:
        状态, 说明 = "失败", f"发送异常：{exc}"
    产出["状态"], 产出["说明"] = 状态, 说明
    # 「部分成功」= 有收件人送达、有收件人没解析到（如填了不存在的人名）：
    # 如实算「已发」但把逐人明细带给前端，避免一个人名写错就让人以为整条都没发出去。
    产出["已发"] = 状态 in ("已发送", "部分成功")
    # 2026-10-30 用户口径：「打开钉钉消息…截图…作为附件」——把本次留下的结果截图逐张发过去。
    # 钉钉一条消息只能带一个文件/图，所以按屏逐张发（最多 12 张，避免长列表刷屏）。
    if 产出["已发"]:
        截图 = []
        try:
            截图 = sorted(p for p in _vbs站点截图目录(脚本键).glob(f"{脚本键}_shot_*.png")
                          if p.stat().st_size > 2000)[:12]
        except Exception:
            截图 = []
        if 截图:
            送达 = 0
            for 图 in 截图:
                try:
                    图状态, _图说明 = 模块.发送截图消息(图, 脚本键, 图.stem)
                except Exception:
                    图状态 = "失败"
                if 图状态 in ("已发送", "部分成功"):
                    送达 += 1
            产出["截图附件"] = [p.name for p in 截图]
            UI信息(f"{脚本键} 钉钉截图：{送达}/{len(截图)} 张已发")
    if 产出["已发"] and 报表:
        try:
            文件状态, 文件说明 = 模块.发送文件消息(报表[0], "", 收件人)
        except Exception as exc:
            文件状态, 文件说明 = "失败", f"发送异常：{exc}"
        产出["文件消息"] = {"状态": 文件状态, "说明": 文件说明}
        if 文件状态 in ("已发送", "部分成功"):
            产出["报表附件"] = [str(报表[0])]
            产出["附件名"] = [报表[0].name]
    UI信息(f"{脚本键} 钉钉消息：{状态}｜{说明}｜报表附件 {len(产出['报表附件'])} 个"
           f"｜收件人 {产出['收件人']}")
    return 产出



_可打开后缀 = {".xls", ".xlsx", ".xlsm", ".csv"}


def _用Excel打开不弹框(路径, 只读: bool = False, auto_close_sec: float = 0) -> dict:
    """用 Excel COM 打开报表，**绝不弹「格式与扩展名不匹配」**（等价于自动点「是」）。

    🔴 2026-10-24 用户明令：「我打开excel 遇到这种破坏，我需要你自动选是，强行操作打开」。
    关键三件事，缺一就会被卡住：
      ① `DisplayAlerts = False` —— 关掉一切警告框（含扩展名不匹配、更新链接）；
      ② 伪装成 .xls 的文本才传 `Format=1` —— 显式声明 Tab 分隔，**跳过文本导入向导**
         （不指定 Format 时 Excel 会弹 3 步向导，比原弹窗更烦）；真 xlsx 传它没意义；
      ③ 用 `DispatchEx` 起**独立实例**（不复用用户正在编辑的 Excel）—— 复用实例在用户
         有别的簿打开/有模态框时会报「操作无法使用」，已实测。
    打开后不 Quit，留给用户看；只把「实例是否要保留」的判断交给调用方。

    🔴 2026-10-24 二次修正（用户截图：标题栏 `Zpp_Cs12_...xlsx [只读]`，问「我要修改要保存怎么办」）：
    以前默认 `ReadOnly=True`，自动打开/打开报表出来的表**一律是只读**，改不了也存不了。
    现在**默认可写**（只读=False）；只有“文件正被别的 Excel 占用”时 Excel 才会自己降级成只读，
    那种情况如实回报「被占用」，让用户知道该关哪个窗口，而不是莫名其妙存不了。
    """
    try:
        import win32com.client as wc
    except Exception as exc:
        return {"ok": False, "说明": f"未安装 pywin32，无法静默打开：{exc}"}
    app = None
    try:
        app = wc.DispatchEx("Excel.Application")
        app.Visible = True
        app.DisplayAlerts = False          # ← 等价于「自动选是」
        app.AskToUpdateLinks = False
        app.ScreenUpdating = True
        打开参数 = {"UpdateLinks": 0, "ReadOnly": bool(只读)}
        伪文本 = _是伪装的文本报表(路径)
        if 伪文本:
            打开参数["Format"] = 1         # xlDelimited(Tab 分隔)，跳过文本导入向导
        簿 = None
        实际只读 = False
        for 第几次 in range(2):
            簿 = app.Workbooks.Open(str(路径), **打开参数)
            实际只读 = False
            try:
                实际只读 = bool(簿.ReadOnly)
            except Exception:
                实际只读 = False
            if not 实际只读:
                break
            # 第一次就只读：多半是「报表转换」刚用完的那个 Excel 实例还没彻底放手，
            # 文件被短暂锁着。关掉重开一次，别让用户莫名拿到只读表。
            if 第几次 == 0:
                try:
                    簿.Close(SaveChanges=False)
                except Exception:
                    pass
                time.sleep(1.5)
        try:
            簿.Activate()
        except Exception:
            pass

        # ---- 自动关闭模式：开一会儿让用户确认后，关掉这个【独立】Excel 实例 ----
        # 仅 zpp_cpgl 等场景需要：用户要「打开看一眼」但不想报表一直占着屏幕。
        # 这里用的是 DispatchEx 起的独立实例，Quit 只影响本报表，绝不碰用户自己打开的簿。
        # auto_close_sec<=0 时走原逻辑：打开后留着不关。
        if auto_close_sec and auto_close_sec > 0:
            try:
                time.sleep(float(auto_close_sec))
                簿.Close(SaveChanges=False)
                app.Quit()
                说明 = ("已用 Excel 打开并自动关闭（文本报表，已跳过格式告警）：" if 伪文本
                        else "已用 Excel 打开并自动关闭：") + Path(路径).name \
                       + f"（停留约 {auto_close_sec:g} 秒）"
                return {"ok": True, "说明": 说明, "文件": Path(路径).name, "路径": str(路径),
                        "行数": None, "只读": 实际只读, "已自动关闭": True}
            except Exception as exc:
                说明 = ("已用 Excel 打开（文本报表，已自动跳过格式告警）：" if 伪文本
                        else "已用 Excel 打开：") + Path(路径).name \
                       + f"（自动关闭失败：{exc}）"
                return {"ok": True, "说明": 说明, "文件": Path(路径).name, "路径": str(路径),
                        "行数": None, "只读": 实际只读, "已自动关闭": False}

        说明 = ("已用 Excel 打开（文本报表，已自动跳过格式告警）：" if 伪文本
                else "已用 Excel 打开：") + Path(路径).name
        if 实际只读:
            说明 += ("　⚠ 但它被锁成【只读】：该文件正被另一个 Excel 窗口或程序占用。"
                    "把占用它的窗口关掉，再点「打开报表」就能正常编辑保存。")
        return {"ok": True, "说明": 说明, "文件": Path(路径).name, "路径": str(路径),
                "行数": None, "只读": 实际只读}
    except Exception as exc:
        # COM 失败（没装 Excel / 被策略拦）→ 退回系统默认程序，至少别让用户点了没反应
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            os.startfile(str(路径))
            return {"ok": True, "说明": f"Excel 自动化打不开（{exc}），已改用系统默认程序打开",
                    "文件": Path(路径).name, "路径": str(路径), "退化": True}
        except Exception as exc2:
            return {"ok": False, "说明": f"打开失败：{exc}；退回默认程序也失败：{exc2}"}


def _sapq打开报表(目录, 用xlsx: bool = False, 指定文件=None) -> dict:
    """打开指定目录里【最近一次】下载/转换出来的报表（只认表格后缀，绝不启动其它类型文件）。
    2026-10-24 用户：「自动打开刚刚下载的那个报表 xls 文件；报表转换为xlsx 也开着的话，以转换后的 excel 打开」。

    🔴 2026-10-24 二次事故（用户现场：「转化失败没讲话，然后打开老的 xls」）：
    以前这里只按「目录里 mtime 最新」挑文件，一旦本次下载/转换失败，挑中的就是**上一次**的
    报表 —— 用户看到的是旧数据还以为刚跑出来的。现在传 指定文件 时只认那一个：
    它不存在就直接报失败，绝不回落到目录里的历史文件。
    """
    try:
        d = Path(目录).expanduser()
        if not d.is_dir():
            return {"ok": False, "说明": "目录不存在：" + str(d)}
        if 指定文件:
            目标 = Path(指定文件)
            if not 目标.is_file():
                return {"ok": False, "说明": f"要打开的文件不存在：{目标.name}（不会打开目录里的旧文件）"}
            if 目标.suffix.lower() in _可打开后缀:
                return _用Excel打开不弹框(目标)
            os.startfile(str(目标))
            return {"ok": True, "说明": "已打开 " + 目标.name, "文件": 目标.name, "路径": str(目标)}
        后缀 = {".xlsx"} if 用xlsx else {".xls", ".xlsm", ".csv"}
        候选 = [f for f in d.iterdir() if f.is_file() and f.suffix.lower() in 后缀]
        if not 候选 and 用xlsx:      # 转换失败/还没转 → 退回 .xls，别让用户点了没反应
            候选 = [f for f in d.iterdir() if f.is_file() and f.suffix.lower() in _可打开后缀]
        if not 候选:
            return {"ok": False, "说明": "目录里没有可打开的报表：" + str(d)}
        最新 = max(候选, key=lambda f: f.stat().st_mtime)
        # 🔴 2026-10-24 用户现场：「我打开excel 遇到这种破坏，我需要你自动选是，强行操作打开」。
        #    SAP 报表是 GBK+Tab 文本却叫 .xls，交给默认程序必然弹
        #    「文件格式和扩展名不匹配…是否仍要打开它？」——必须先点「是」。
        #    表格类一律走 Excel COM（DisplayAlerts=False + 仅文本传 Format=1），彻底不弹框。
        if 最新.suffix.lower() in _可打开后缀:
            return _用Excel打开不弹框(最新)
        os.startfile(str(最新))       # 非表格后缀才交给默认程序
        return {"ok": True, "说明": "已打开 " + 最新.name, "文件": 最新.name, "路径": str(最新)}
    except Exception as exc:
        return {"ok": False, "说明": "打开失败：" + str(exc)}


class SapqOpenRequest(BaseModel):
    script: str = ""
    用xlsx: bool = False
    下载目录: str = ""
    转换路径: str = ""


@app.post("/api/sapq/open")
def sapq_open(payload: SapqOpenRequest):
    """页面「打开报表」按钮：立刻打开最近一次下载/转换的报表（用户 2026-10-24）。"""
    目录 = (payload.转换路径 if (payload.用xlsx and str(payload.转换路径 or "").strip()) else payload.下载目录)
    结果 = _sapq打开报表(_sapq确保下载目录(目录 or None), payload.用xlsx)
    UI信息(f"打开报表（{payload.script}）：{结果.get('说明')}")
    return 结果


@app.post("/api/vbs/run")
def vbs_run(payload: VbsRunRequest):
    """运行一个 SAP GUI 自动化 VBS：cscript 同步执行，stdout/stderr 全部回传并落盘。

    进程句柄记入 VBS_PROC，供「重置」按钮经 /api/vbs/stop 强行中断（用户 2026-10-19 明令）。
    """
    global VBS_PROC, VBS_STOP_REQUESTED
    script = (payload.script or "").strip()
    站点键 = script if script in VBS_SCRIPTS_SITE else None
    if 站点键:
        # 「选择屏查询」页（zb02 / zb03 / zpp_po / zpp_pur）：脚本在 app/vbs_scripts/
        文件名 = VBS_SCRIPTS_SITE[站点键][0]
        基目录 = VBS_SITE_BASE
    else:
        文件名 = VBS_SCRIPTS.get(script)
        基目录 = VBS_BASE
    if not 文件名:
        可用 = "、".join(list(VBS_SCRIPTS) + list(VBS_SCRIPTS_SITE))
        raise HTTPException(status_code=400,
                            detail=f"脚本 {script!r} 不在白名单（可用：{可用}）")
    脚本路径 = (基目录 / 文件名).resolve()
    if not str(脚本路径).startswith(str(基目录.resolve())) or not 脚本路径.is_file():
        raise HTTPException(status_code=500, detail=f"VBS 文件不存在：{脚本路径}")
    # 动态条件：网页选择屏现值 → 每次改写运行时 VBS 副本（原脚本不动），执行副本
    注入摘要 = None
    本次报表文件 = ""          # 本次报表下载的文件名（功能名_时间戳.xls）
    本次报表路径 = None        # 磁盘上的报表（两个开关都开时作为邮件附件）
    # 页面「下载路径」框优先（用户 2026-10-23：自定义路径优先级更高），留空 = 代码默认目录
    本次下载目录 = _sapq解析下载目录(payload.下载目录) if 站点键 else SAPQ_下载目录
    if 站点键 and (payload.params or payload.download or payload.skipQuery):
        # 「UI 输入什么 → VBS 就改成什么 → SAP 同步并执行」的核心一步（用户 2026-10-23 需求）
        try:
            # 2026-10-30 用户口径：「打开钉钉消息或者邮件通知的时候」都要截图（截图要作为附件跟着走），
            # 所以截图模式 = 邮件 **或** 钉钉 任一打开，不再只认邮件。
            # 2026-11-05 用户：Zpp_Pr2po 单独退出截图（彻底取消滚动截图）——该页通知以「报表附件 + 正文成功行」为准，
            # 不再截 SAP 屏；因此 要截图 对 zpppr2po 恒为 False（钉钉/邮件也不触发截图）。
            要截图 = bool(payload.mail or payload.钉钉) and 站点键 != "zpppr2po"
            if 要截图:
                try:
                    目录 = _vbs站点截图目录(站点键)
                    目录.mkdir(parents=True, exist_ok=True)
                    for 旧图 in 目录.glob("*.png"):
                        旧图.unlink()
                except Exception as exc:
                    UI错误(f"{站点键} 截图目录预清理失败（不影响执行）：{exc}")
            if payload.download:
                本次报表文件 = _sapq报表文件名(站点键)
                本次报表路径 = _sapq确保下载目录(本次下载目录) / 本次报表文件
            else:
                本次报表文件 = ""
                本次报表路径 = None
            脚本路径, 注入摘要 = 生成选择屏运行时脚本(
                站点键, payload.params or {}, 要截图,
                bool(payload.download), 本次报表文件, bool(payload.skipQuery),
                本次下载目录)
            未注入说明 = (f" [警告]形态异常未注入 {len(注入摘要['unmatched'])} 项（字段 ID 缺 wnd[0]/usr/ 前缀？）" if 注入摘要.get('unmatched') else "")
            UI信息(f"{站点键} 运行时 VBS 已按页面条件改写：{脚本路径.name}"
                   f"（注入字段 {len(注入摘要['fields'])} 项"
                   f" 跳过 {len(注入摘要['skipped'])} 项"
                   f" 页面缺键 {len(注入摘要['missing'])} 项{未注入说明}"
                   f" 邮件模式={'开' if payload.mail else '关'}"
                   f" 报表下载={'开(' + 本次报表文件 + ')' if payload.download else '关'}"
                   f" 跳过查询={'是' if payload.skipQuery else '否'}）")
        except Exception as exc:
            UI错误(f"{站点键} 运行时 VBS 生成失败，回退原脚本：{exc}")
            脚本路径 = (基目录 / 文件名).resolve()
            注入摘要 = None
            本次报表路径 = None
    elif script == "co28" and payload.params:
        try:
            # 2026-10-23：mail 模式本次运行「开跑前」先清空截图目录。
            # 原因：VBS 只在真正走到 Co28ShotAllScreens 时才清目录；若中途「无数据行」提前
            # Quit 0，目录会残留上一次的图 → 成功邮件可能引用到旧图（图文不对应）。
            # 2026-10-31 修复：只清 VBS 的运行截图 co28_shot_*.png，不动 co28_fail_*（失败邮件
            # 正在引用的 SAP/网页截图）。今天现场：上一封失败邮件还在 dws 重试（11:53:29~44），
            # 新运行 11:53:32 开跑清空了整个目录 → 内联附件文件消失 → 邮件整封发不出去。
            本次截图目录已清 = False
            if payload.mail:
                try:
                    CO28_SHOTS_DIR.mkdir(parents=True, exist_ok=True)
                    for 旧图 in CO28_SHOTS_DIR.glob("co28_shot_*.png"):
                        旧图.unlink()
                    本次截图目录已清 = True
                except Exception as exc:
                    UI错误(f"CO28 截图目录预清理失败（不影响执行）：{exc}")
            脚本路径, 注入摘要 = 生成co28运行时脚本(payload.params, bool(payload.mail))
            注入摘要["截图目录已清"] = 本次截图目录已清
            UI信息(f"CO28 运行时 VBS 已按页面条件改写：{脚本路径.name}"
                   f"（P_AUTFP={注入摘要['autfp']} P_PROFID={注入摘要['prof']}"
                   f" 字段={len(注入摘要['fields'])} 跳过={len(注入摘要['skipped'])}"
                   f" 邮件模式={'开' if 注入摘要.get('mail') else '关'}"
                   f" 截图目录已清={'是' if 本次截图目录已清 else '否'}）")
        except Exception as exc:
            UI错误(f"CO28 运行时 VBS 生成失败，回退原脚本：{exc}")
            脚本路径 = (VBS_BASE / 文件名).resolve()
            注入摘要 = None
    # ---------- 2026-11-05 用户：SAP 计划任务「执行前先连按 2 次 Enter 热键」 ----------
    # 只对传了 前置回车>0 的计划任务生效；网页手动点「执行查询」默认 0，行为零变化。
    # 前置动作单独跑一支极小 VBS（不碰任何 .vbs 底稿），SAP 未就绪时安静跳过、绝不阻断主脚本。
    前置回车说明 = ""
    if payload.前置回车 > 0:
        try:
            if str(APP_DIR) not in sys.path:
                sys.path.insert(0, str(APP_DIR))
            import sap_pre_enter as _SPE
            前置回车说明 = _SPE.按回车(int(payload.前置回车))
        except Exception:
            前置回车说明 = "[PreEnter] 本次未发送 Enter（主脚本照常执行）"
        UI信息(f"{script} 前置回车：{前置回车说明}")
    if 站点键 and payload.download and not payload.mail and not payload.报表邮件:
        UI信息(f"{script} 计划任务口径：SAP 报表结果邮件已关闭，只保留任务汇总通知一封")
    cscript = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "cscript.exe"
    解释器 = str(cscript) if cscript.is_file() else "cscript"
    环境 = dict(os.environ)
    环境["PYTHONIOENCODING"] = "utf-8"
    开始 = datetime.now()
    with VBS_PROC_LOCK:
        if VBS_PROC is not None and VBS_PROC.poll() is None:
            raise HTTPException(status_code=409, detail="已有 VBS 脚本在运行；点「重置」可强行中断")
        VBS_STOP_REQUESTED = False
        try:
            VBS_PROC = subprocess.Popen(
                [解释器, "//nologo", "//E:vbscript", str(脚本路径)],
                cwd=str(基目录), env=环境,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as exc:
            VBS_PROC = None
            UI错误(f"VBS 启动失败（{script}）：{exc}")
            raise HTTPException(status_code=500, detail=f"VBS 启动失败：{exc}")
    被中断 = False
    原始字节 = b""
    try:
        原始字节, _ = VBS_PROC.communicate(timeout=VBS_TIMEOUT)
    except subprocess.TimeoutExpired as 超时:
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(VBS_PROC.pid)],
                           capture_output=True,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass
        # 2026-10-23：超时也必须落日志（此前超时不留任何痕迹，网页只报笼统「执行异常」无法排障）
        try:
            VBS_LOG_DIR.mkdir(parents=True, exist_ok=True)
            部分输出 = (getattr(超时, "stdout", None) or b"").decode("gbk", errors="replace")
            (VBS_LOG_DIR / f"{开始:%Y%m%d_%H%M%S}_{script}_timeout.log").write_text(
                f"[运行超时 >{VBS_TIMEOUT}s，cscript 进程树已强停]\n" + 部分输出,
                encoding="utf-8", errors="replace")
        except Exception:
            pass
        UI错误(f"VBS 运行超时（{script}，>{VBS_TIMEOUT}s），已强停；详见 vbs_logs/*_timeout.log")
        raise HTTPException(status_code=504,
                            detail=f"VBS 运行超时（>{VBS_TIMEOUT} 秒）已强停：SAP GUI 是否卡在弹窗或网络等待？可点「重置」后重试")
    except Exception as exc:
        # 2026-10-23：运行异常同样落日志
        try:
            VBS_LOG_DIR.mkdir(parents=True, exist_ok=True)
            (VBS_LOG_DIR / f"{开始:%Y%m%d_%H%M%S}_{script}_error.log").write_text(
                f"[VBS 运行异常] {exc}\n", encoding="utf-8", errors="replace")
        except Exception:
            pass
        UI错误(f"VBS 运行异常（{script}）：{exc}")
        raise HTTPException(status_code=500, detail=f"VBS 运行异常：{exc}")
    finally:
        with VBS_PROC_LOCK:
            被中断 = VBS_STOP_REQUESTED
            VBS_PROC = None
            VBS_STOP_REQUESTED = False
    原始 = 原始字节 or b""
    try:
        输出 = 原始.decode("utf-8")
    except UnicodeDecodeError:
        输出 = 原始.decode("gbk", errors="replace")
    # 前置回车说明是固定安全文本（不含错误/失败/error 字样），拼在开头不会误伤下面「有无真实错误」判据
    if 前置回车说明:
        输出 = 前置回车说明 + "\n" + 输出
    # 2026-10-23 排障：cscript 管道输出出现锟斤拷乱码（混合编码）时，落一份原始字节快照，
    # 供下次直接按字节分析真实编码（日志本身已经过解码，无法还原 raw）
    try:
        if "锟" in 输出 or "\ufffd" in 输出:
            VBS_LOG_DIR.mkdir(parents=True, exist_ok=True)
            (VBS_LOG_DIR / f"{开始:%Y%m%d_%H%M%S}_{script}.rawbin").write_bytes(原始)
    except Exception:
        pass
    # cscript 对脚本运行时错误也返回 exit 0，必须按输出里的错误文本判定成败。
    # 2026-10-31 修复：SAP COM 运行时错误（如 `co28_runtime.vbs(349, 9) (null): 完成该操作
    # 所需的数据还不可使用`）原文不含「错误/error」二字，被误判成「正常」→ 后端发了成功邮件、
    # 前端又按 .vbs(N,N) 判失败发失败邮件，一次运行双发。现与前端同一判据：
    # 错误特征 = 错误/error/exception/Err.Raise/脚本错误/执行失败/失败：/.vbs(行, 列)，
    # 但 [co28] 警告、[batch-fail]、[ov] 警告 不算错误（它们是明示的非致命提示）。
    _错误行 = _re.compile(
        r"错误|error|exception|Err\.Raise|脚本错误|执行失败|失败：|\.vbs\(\d+\s*,\s*\d+\)",
        _re.IGNORECASE)
    _非错误行 = _re.compile(r"\[co28\]\s*警告|\[batch-fail\]|\[ov\]\s*警告")
    有真实错误 = any(_错误行.search(x) and not _非错误行.search(x)
                 for x in (输出 or "").splitlines())
    正常 = (not 被中断) and (not 有真实错误)
    UI信息(f"VBS 运行结束：{script}（{'被强行中断' if 被中断 else '正常结束'}）")

    # ---------- 批量写入失败点名（2026-10-27 用户现场：界面报「已写入 N 项」，SAP 里其实空着）----------
    # VBS 侧每个走「多重选择框」的字段，写不进去时会回显 "[batch-fail] <网页data-key> <原因>"。
    # 这里解析出来如实回传前端 —— toast 必须说「N 项未写入 SAP」，不能继续报成功。
    # 注意：[batch-fail] 的文案里不能出现「错误/error」，否则上面的 正常 判定会把整次运行判成失败。
    批量失败 = []
    for _行 in (输出 or "").splitlines():
        _行 = _行.strip()
        if _行.startswith("[batch-fail]"):
            _体 = _行[len("[batch-fail]"):].strip()
            if _体 and _体 not in 批量失败:
                批量失败.append(_体)
    if 批量失败:
        UI错误(f"VBS 批量写入未完成 {len(批量失败)} 项：{'；'.join(批量失败[:6])}")
    try:
        VBS_LOG_DIR.mkdir(parents=True, exist_ok=True)
        (VBS_LOG_DIR / f"{开始:%Y%m%d_%H%M%S}_{script}.log").write_text(
            ("[被用户强行中断]\n" if 被中断 else "") + 输出,
            encoding="utf-8", errors="replace")
    except Exception:
        pass

    # ---------- 极速导出：VBS 把清单放进剪贴板，这里落盘（2026-10-25 新增） ----------
    # 触发条件 = VBS 回显 "[clip] 已把清单写入剪贴板" 且本次确有下载目标。
    # 只在 VBS 自己判定「行数够多、剪贴板更快」时才会走这条；小结果仍走文件导出（更快更稳）。
    if payload.download and 站点键 and ("[clip] 已把清单写入剪贴板" in (输出 or "")):
        try:
            剪贴 = _读剪贴板文本()
            转换后 = _sapq剪贴板转TSV(剪贴) if 剪贴 else ""
            if 本次报表路径 and len(转换后) > 100:
                本次报表路径.parent.mkdir(parents=True, exist_ok=True)
                本次报表路径.write_text(转换后, encoding="gbk", errors="replace", newline="\r\n")
                输出 += (f"\n[dl] 极速导出：剪贴板 {len(剪贴)} 字符 → {本次报表路径.name}"
                         f"（{len(转换后.splitlines())} 行）")
                UI信息(f"{站点键} 极速导出（剪贴板）：{本次报表路径.name}"
                       f" {len(转换后.splitlines())} 行")
            else:
                UI错误(f"{站点键} 剪贴板导出异常：读到 {len(剪贴)} 字符，未落盘")
                输出 += "\n[dl] 极速导出失败：剪贴板内容异常，本次没有报表文件"
        except Exception as exc:
            UI错误(f"{站点键} 剪贴板落盘失败：{exc}")
            输出 += f"\n[dl] 极速导出失败：{exc}"

    # ---------- 「报表下载」开关：核对本次报表是否真的落盘 ----------
    报表下载信息 = None
    if payload.download and 站点键:
        已存回显 = ""
        for 行_ in (输出 or "").splitlines():
            if "[dl] 已保存：" in 行_:
                已存回显 = 行_.split("：", 1)[-1].strip()
                break
        已落盘 = bool(本次报表路径 and 本次报表路径.is_file())
        报表下载信息 = {"ok": 已落盘, "文件": 本次报表文件,
                    "目录": str(本次下载目录),
                    "路径": str(本次报表路径) if 本次报表路径 else "",
                    "vbs": 已存回显}
        UI信息(f"{站点键} 报表下载：{'成功' if 已落盘 else '未生成'} → {报表下载信息['路径']}")

    # ---------- 「报表转换为xlsx」开关：把刚下载的 .xls 用 xlwings 另存为同名 .xlsx ----------
    # 用户 2026-10-24：「报表转换为xlsx 只有在报表下载的开关开启的状态才能开启/关闭；
    #                   作用是将报表下载对应的 xls 文件自动转换为同名的 xlsx 文件（用 Python 就必须用 xlwings）」。
    # 路径：页面「转换路径」留空 = 跟随「报表下载」路径；填了以页面为准（且不反向改动下载路径）。
    转换信息 = None
    if payload.报表转换 and 站点键:
        if not payload.download:
            转换信息 = {"ok": False, "说明": "「报表下载」开关没开 —— 没有可转换的文件"}
        elif not (本次报表路径 and 本次报表路径.is_file()):
            转换信息 = {"ok": False, "说明": f"本次没下载到报表，无法转换（{本次报表文件 or '未生成'}）"}
        else:
            转目录 = (_sapq解析下载目录(payload.报表转换路径) if str(payload.报表转换路径 or "").strip()
                     else 本次下载目录)
            转换信息 = _xls转xlsx(本次报表路径, 转目录)
            UI信息(f"{站点键} 报表转换：{转换信息.get('说明')}")

    # ---------- 「自动打开」开关：用本机默认程序打开刚下载的报表（2026-10-24 用户需求）----------
    # 用户原话：「自动打开刚刚下载的那个报表 xls 文件；如果报表转换为xlsx 也处于打开状态的话，
    #           请你以转换后的 excel 打开」。只在「报表下载」开着、且确实落了文件时才打开。
    打开信息 = None
    if payload.自动打开 and 站点键 and payload.download:
        # 🔴 2026-10-24 事故修正：只打开「本次」产出的文件。
        #   以前按目录里 mtime 最新挑，本次下载/转换失败时就打开上一次的旧 xls（用户当场发现
        #   「转化失败没讲话，然后打开老的 xls」）。现在：转换成功→开转换后的 xlsx；
        #   转换没开/本次 .xls 确实在→开本次 .xls；两个都没有→明确报「本次没有可打开的文件」。
        待开 = None
        if 转换信息 and 转换信息.get("ok") and 转换信息.get("路径"):
            待开 = Path(str(转换信息["路径"]))
        elif 本次报表路径 and Path(本次报表路径).is_file() and not payload.报表转换:
            待开 = Path(本次报表路径)
        if 待开 is None:
            打开信息 = {"ok": False, "说明": "本次没有可打开的报表（下载/转换都没产出文件），已跳过，未打开目录里的旧文件"}
        else:
            打开信息 = _sapq打开报表(待开.parent, bool(转换信息 and 转换信息.get("ok")), 指定文件=待开)
        UI信息(f"{站点键} 自动打开：{打开信息.get('说明')}")

    # ---------- 2026-11-06 用户需求：报表下载开关打开 → 准备 Zpp_Pr2po 的「正文筛选表」与「附件报表文件」----------
    #   · 正文筛选表：按条件（P125/201253、P183/201007、停产供方）挑行，进邮件正文（纯文本表格）；
    #   · 报表文件：优先用已转换的 xlsx，否则把 .xls 就地转一份 xlsx，供邮件附件「整理」用。
    正文表格 = None
    报表文件 = None
    if payload.download and 站点键 and 报表下载信息 and 报表下载信息.get("ok") and 本次报表路径:
        if 转换信息 and 转换信息.get("ok") and 转换信息.get("路径"):
            报表文件 = Path(str(转换信息["路径"]))
        else:
            转 = _xls转xlsx(本次报表路径, 本次下载目录)
            报表文件 = Path(str(转["路径"])) if 转.get("ok") else Path(本次报表路径)
        if 站点键 == "zpppr2po" and 报表文件.is_file():
            正文表格 = _sapq筛选行(报表文件)
            if 正文表格 is None:
                UI错误(f"zpppr2po 报表筛选失败（无法解析表头）：{报表文件}")
            elif len(正文表格) <= 1:
                UI信息(f"zpppr2po 报表：筛选后无符合条件的数据")
            else:
                UI信息(f"zpppr2po 报表：筛选命中 {len(正文表格) - 1} 行")

    # ---------- 「发邮件」开关：结果邮件（截图内联 + 报表作附件）----------
    # 用户 2026-10-23：「如果发送邮件跟报表下载都打开了，我希望报表作为附件发送邮件」
    # 🔴🔴 2026-11-06 用户明令：「MESE限速器以及MESE安全钳 我不需要你发邮件给我」
    #   MESE 两条线内部驱动 SAP 下载时传 不发本人邮件=True → 下面所有发信分支一律跳过。
    #   🔴 其它模块（SAP 查询页手动执行 / 计划任务 / CO28）默认 False，行为一字不变。
    禁发本人邮件 = bool(getattr(payload, "不发本人邮件", False))
    if 禁发本人邮件:
        UI信息(f"{script} MESE 口径：本次不发任何结果邮件给演示用户本人（用户 2026-11-06 明令）")
    邮件结果 = None
    附件 = []
    if payload.mail and 报表下载信息 and 报表下载信息.get("ok") and 本次报表路径:
        附件 = [本次报表路径]
    if script == "co28" and payload.mail and not 禁发本人邮件:
        邮件结果 = _co28发送结果邮件(正常, 被中断, 输出, 开始, 收件人=payload.收件人)
        UI信息(f"CO28 结果邮件：{邮件结果}")
    elif 站点键 and payload.mail and not 禁发本人邮件:
        # 三个查询页：执行结果 + （Zpp_Pr2po 不截图）报表附件 → 收件人由页面「收件人」框决定。
        # 2026-11-06 用户明令：只有点「发邮件」才发通知；下载表格按钮只下载、不邮件。
        # Zpp_Pr2po：附件改为「整理后的 xlsx」（删异常提示行 / 列宽自适应 / 冻结表头），正文放筛选表。
        _附件 = 附件
        if 站点键 == "zpppr2po" and 报表文件 and 报表文件.is_file():
            _整理 = _sapq整理报表xlsx(报表文件)
            if _整理.get("ok"):
                _附件 = [Path(str(_整理["路径"]))]
                UI信息(f"zpppr2po 报表整理：{_整理.get('说明')}")
            else:
                UI错误(f"zpppr2po 报表整理失败：{_整理.get('说明')}")
        邮件结果 = _发送查询结果邮件(站点键, 正常, 被中断, 输出, 开始, payload.params or {},
                                     收件人=payload.收件人, 附件=_附件, 下载目录=本次下载目录,
                                     正文表格=正文表格, 附截图=False)
        UI信息(f"{站点键} 结果邮件：{邮件结果}")

    # ---------- 「钉钉消息」开关：结果摘要发钉钉单聊（2026-10-24 用户需求）----------
    # 用户原话：「发邮件后面再加一个钉钉消息或钉钉群通知…输入名字的话自动发送到对应的聊天窗口…
    #           如果输入框以及报表下载都打开状态，我希望发给演示用户附带报表」
    # 口径与结果邮件完全一致（同一份状态栏/条件/报表来源），只是通道换成钉钉单聊（dws chat +messages-send --as user）。
    钉钉结果 = None
    if payload.钉钉:
        if script == "co28":
            # CO28 不在 VBS_SCRIPTS_SITE 里（站点键为空）——2026-10-24 用户反馈「CO28 开了钉钉却没发内容」
            # 就是这里被「and 站点键」挡掉的；CO28 改走自己的口径（状态栏 / 勾选汇总 / co28_shots 截图）。
            钉钉结果 = _发送co28结果钉钉(正常, 被中断, 输出, 开始, 收件人=payload.钉钉收件人)
            UI信息(f"CO28 钉钉消息：{钉钉结果}")
        elif 站点键:
            钉钉结果 = _发送查询结果钉钉(站点键, 正常, 被中断, 输出, 开始, payload.params or {},
                                        收件人=payload.钉钉收件人, 附件=附件)
            UI信息(f"{站点键} 钉钉消息：{钉钉结果}")

    return {"ok": 正常, "script": script, "path": str(脚本路径),
            "interrupted": 被中断,
            "injected": 注入摘要, "自动打开": 打开信息,
            "前置回车": 前置回车说明,
            "batchFail": 批量失败,
            "mail": 邮件结果,
            "钉钉": 钉钉结果,
            "报表转换": 转换信息,
            "download": 报表下载信息,
            "cost": round((datetime.now() - 开始).total_seconds(), 1),
            "output": 输出[-8000:]}     # 只回传尾部，防止 SAP 弹窗刷屏


@app.post("/api/vbs/stop")
def vbs_stop():
    """强行中断正在运行的 SAP GUI VBS 脚本（taskkill 杀 cscript 进程树）。"""
    global VBS_STOP_REQUESTED
    with VBS_PROC_LOCK:
        proc = VBS_PROC
        if proc is None or proc.poll() is not None:
            return {"ok": False, "stopped": False, "message": "当前没有正在运行的 VBS 脚本"}
        VBS_STOP_REQUESTED = True
    try:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                       capture_output=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as exc:
        return {"ok": False, "stopped": False, "message": f"中断失败：{exc}"}
    return {"ok": True, "stopped": True, "message": "已强行中断 SAP 脚本（cscript 进程树已终止）"}


# ============================================================================
# ZPP_CPGL 集成（修改模块 ·「修改代办」页「zpp_cpgl」按钮；2026-10-24 用户需求）
# ----------------------------------------------------------------------------
# 用户口径（原话拆解）：
#   ① 修改代办页原来那排「所有修改待办抓取 / 提交无需通知 / 停止 / 日志」整排去掉，
#      改成两个按钮：zpp_cpgl、sq00预约状态（SQ00 先占位，等用户给脚本/口径）。
#   ② 点「zpp_cpgl」时的物料号来源优先级：**剪贴板 > 修改代办兜底**；
#      兜底 = 「修改代办」里「时间戳最新」那一批「管控级别 = 通知并反馈」的物料号。
#   ③ 跑 SAP **之前**，先把这批物料号贴进「修改合同状态」的「物料号」列、
#      「修改计划」的「修改的合同号」列（用户 2026-10-24 确认：修改计划只贴合同号，其它列不管）。
#      🔴 用户 2026-10-24 23:50 追加：「点击 zpp_cpgl 的一瞬间（甚至之前）就要复制过去，
#      不能等执行完再复制」→ 拆成两步：POST /prepare（毫秒级写入 + 前端立刻刷新表格）→
#      POST /run（带 已写入=True + 写入ids，只跑 SAP，不再重复插入、不重读剪贴板）。
#   ④ 跑 /nzpp_cpgl（VBS = vbs_scripts/Zpp_Cpgl.vbs，物料号按次注入）→ 下载报表
#      （Zpp_Cpgl_<时间戳>.xls，落点与其它 SAPGUI 脚本一致：页面下载路径 > 代码默认 web_xls）。
#   ⑤ 按物料号在报表里匹配行 → 「入库情况」= 从右往左第一个有值的发货类列组的组合值
#      （日期 + 抬头 + 日期后第一串，例 20260910成品出226H）→ 只回填「修改合同状态」的「入库情况」列。
#      🔴 用户 2026-11-06 明令：zpp_cpgl 这张报表只针对「入库情况」，「提货情况」必须留空；
#      2026-10-25 曾误加「入库情况单锁成品入」（20260810成品入226H），已废止。
#   ⑤b 报表只生成、不自动打开 Excel（用户 2026-11-06 明令，撤掉 2026-10-25「打开后 2 秒自动关闭」）。
#   ⑥ 「修改合同状态」新插入的行放最上面、原有数据一行不清；**「修改计划」插入前先整表清空**
#      （用户 2026-10-25 明令），表里只留本次这批号；清空前自动快照，可「恢复」。
#
# 线程模型与 modify_todo_auto 一致：run 端点只负责起 daemon 线程（SAP 跑几分钟不能占死请求），
# 前端按 after 游标轮询 status 增量取日志；同一时刻只允许一个任务（409）。
# ============================================================================
ZPPCPGL_键 = "zpp_cpgl"
ZPPCPGL_VBS = VBS_SITE_BASE / "Zpp_Cpgl.vbs"
ZPPCPGL_运行时 = VBS_SITE_BASE / "zpp_cpgl_runtime.vbs"
ZPPCPGL_超时 = 900                     # 单次 SAP 跑批上限（秒）：F8 大清单 + 导出留足
ZPPCPGL_表_物料列 = {"修改合同状态": "物料号", "修改计划": "修改的合同号"}
# 用户 2026-10-25 明令：「修改计划」在插入之前先把这张表的数据**全部清空**（只留本次这批号）。
# 「修改合同状态」不变（插最上面、历史行保留）。清空前照样走 _快照存储，点错可用工具栏「恢复」找回。
ZPPCPGL_清空后再插入的表 = {"修改计划"}
ZPPCPGL_任务 = {"job_id": None, "running": False, "lines": [], "result": None,
                "error": "", "开始": None, "结束": None}
ZPPCPGL_锁 = threading.Lock()


def _cpgl日志(文本: str) -> None:
    """写一条任务日志（前端轮询 status 时增量取走）。"""
    行 = f"[{datetime.now():%H:%M:%S}] {文本}"
    with ZPPCPGL_锁:
        ZPPCPGL_任务["lines"].append(行)
    print(f"[ZPPCPGL] {文本}", flush=True)


def _cpgl读剪贴板() -> str:
    """读 Windows 剪贴板（PowerShell Get-Clipboard -Raw）。

    为什么放后端读：网页 navigator.clipboard.readText() 在 http://localhost:8000 上
    多数浏览器直接拒绝（要求 https 或手势授权），而这就是本机服务，后端读最稳。
    """
    try:
        系统目录 = Path(os.environ.get("SystemRoot", r"C:\Windows"))
        路径 = 系统目录 / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
        解释器 = str(路径) if 路径.is_file() else "powershell"
        r = subprocess.run(
            [解释器, "-NoProfile", "-NonInteractive", "-STA", "-Command",
             "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-Clipboard -Raw"],
            capture_output=True, timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        原 = r.stdout or b""
        for 编码 in ("utf-8-sig", "utf-8", "gbk"):
            try:
                return 原.decode(编码)
            except Exception:  # noqa: BLE001
                continue
        return 原.decode("gbk", errors="replace")
    except Exception as exc:  # noqa: BLE001
        _cpgl日志(f"读剪贴板失败（不影响兜底）：{type(exc).__name__}: {exc}")
        return ""


def _cpgl兜底物料号():
    """「修改代办」兜底：时间戳最新的那批「通知并反馈」物料号（保持表内顺序、去重）。"""
    store = _modify_store_seed_if_empty()
    候选 = []
    for r in (store["sheets"].get("修改代办") or []):
        d = r.get("data") or {}
        if str(d.get("管控级别") or "").strip() != "通知并反馈":
            continue
        物料 = zpp_cpgl_parse.规范化物料号(d.get("物料号"))
        if not 物料:
            continue
        候选.append((str(d.get("时间戳") or "").strip(), 物料))
    if not 候选:
        return [], ""
    最新 = max(ts for ts, _ in 候选)
    结果 = []
    for ts, 物料 in 候选:
        if ts == 最新 and 物料 not in 结果:
            结果.append(物料)
    return 结果, 最新


def _cpgl规范物料列表(列表) -> list:
    """物料号列表 → 规范化 + 去重（只留合法形态）。"""
    out = []
    for x in (列表 or []):
        u = zpp_cpgl_parse.规范化物料号(x)
        if u and zpp_cpgl_parse.物料正则.match(u) and u not in out:
            out.append(u)
    return out


def _cpgl取物料(用剪贴板: bool, 指定=None):
    """物料号来源（用户口径：剪贴板优先，剪贴板没有才回落到「修改代办」）。

    返回 (物料号列表, 来源说明)。三步优先级：
      ① 调用方显式传入（一般不用，留给以后别的页面联动）
      ② 剪贴板（后端 PowerShell Get-Clipboard；浏览器 API 在 http 下读不了）
      ③ 「修改代办」里时间戳最新的那批「管控级别 = 通知并反馈」行的物料号
    """
    物料 = _cpgl规范物料列表(指定)
    if 物料:
        return 物料, "调用方传入"
    if 用剪贴板:
        剪 = _cpgl读剪贴板()
        物料 = zpp_cpgl_parse.提取物料号(剪)
        if 物料:
            return 物料, f"剪贴板（{len(剪)} 字符）"
    _cpgl日志("   剪贴板里没有物料号（或剪贴板读不到），改用「修改代办」兜底")
    物料, 时间戳 = _cpgl兜底物料号()
    if 物料:
        return 物料, f"修改代办·时间戳 {时间戳} 的「通知并反馈」"
    return [], ""


def _cpgl前置插入(物料号们, 来源: str):
    """把物料号写进两张表，返回 (新行 id 字典, 清空统计字典)。

    · 「修改计划」（ZPPCPGL_清空后再插入的表）：**先整表清空再插**（用户 2026-10-25 明令），
      表里只剩本次这批号，历史行不再保留（清空前有快照，可「恢复」）。
    · 「修改合同状态」：插到最上面、原有行一行不动，row_no 整体顺延。
    """
    store = _modify_store_seed_if_empty()
    # 插入 / 清空都会改表内容 → 动手前留一份快照（最近 20 份，工具栏「恢复」可一键找回）
    _快照存储(MODIFY_STORE_PATH, f"zpp_cpgl 写入 {len(物料号们)} 行（{来源}）")
    新id, 清空 = {}, {}
    for sheet, 列名 in ZPPCPGL_表_物料列.items():
        if sheet in ZPPCPGL_清空后再插入的表:
            清空[sheet] = len(store["sheets"].get(sheet) or [])
            store["sheets"][sheet] = []      # ① 先清空（快照已留底）
            行们 = store["sheets"][sheet]
        else:
            行们 = store["sheets"].setdefault(sheet, [])
        ids = []
        for i, 物料 in enumerate(物料号们):
            rid = int(store.get("next_id") or 1)
            store["next_id"] = rid + 1
            行们.insert(i, {"id": rid, "row_no": i + 1, "data": {列名: 物料}})
            ids.append(rid)
        for i, r in enumerate(行们, 1):     # ② 全表 row_no 从 1 重排
            r["row_no"] = i
        新id[sheet] = ids
    _modify_store_save(store)
    return 新id, 清空


def _cpgl回填提货情况(新id们, 匹配结果):
    """只把「入库情况」写回刚插入的那几行「修改合同状态」（按 id 精确定位，不碰别的行）。

    用户 2026-11-06 明令：zpp_cpgl 这张报表**只针对「入库情况」**，「提货情况」应该空值。
    所以这里绝不写「提货情况」列（历史行已有的值也不动）；入库情况为空仍写 "-"（2026-10-25 明令）。
    返回 (提货情况已填行数=恒 0, 入库情况命中真实值行数)，保留旧调用签名以免扩散。
    """
    store = _modify_store_seed_if_empty()
    按id = {r.get("id"): r for r in (store["sheets"].get("修改合同状态") or [])}
    提填, 入填 = 0, 0
    for rid, 结果 in zip(新id们, 匹配结果):
        r = 按id.get(rid)
        if not r:
            continue
        入 = str(结果.get("入库情况") or "").strip()
        if 入:
            r.setdefault("data", {})["入库情况"] = 入
            入填 += 1
        else:
            # 用户 2026-10-25：「入库情况匹配为空的话，默认填写 - 这个字符」→ 空也用 "-" 占位，
            # 让表里一眼能看出「已跑过、确实没有数据」，而不是留白分不清没跑还是没数据。
            r.setdefault("data", {})["入库情况"] = "-"
    _modify_store_save(store)
    return 提填, 入填


def _cpgl生成运行时(物料号们, 目录: Path, 文件名: str) -> Path:
    """按本次物料号/落点/文件名改写运行时 VBS（原脚本 vbs_scripts/Zpp_Cpgl.vbs 一行不动）。"""
    if not ZPPCPGL_VBS.is_file():
        raise RuntimeError(f"VBS 脚本不存在：{ZPPCPGL_VBS}")
    文本 = ZPPCPGL_VBS.read_bytes().decode("gbk", errors="replace")
    # 2026-11-06：源脚本曾被误插一行没带「'」的中文注释残留，cscript 直接报
    # 「Microsoft VBScript 编译器错误: 无效字符」，前端只显示「本次没有生成报表」很难定位。
    # 生成运行时前先做静态兜底，把问题直接指到行号；只作用于 zpp_cpgl 这条链。
    坏行 = []
    for i, l in enumerate(文本.splitlines(), 1):
        t = l.strip()
        if t and not t.startswith("'") and ord(t[0]) > 127:
            坏行.append((i, t[:60]))
    if 坏行:
        raise RuntimeError(
            f"Zpp_Cpgl.vbs 第 {坏行[0][0]} 行有未注释的中文/全角字符，"
            f"VBScript 编译会报「无效字符」：{坏行[0][1]}（共 {len(坏行)} 行）"
        )
    文本 = _vbs注入常量(文本, "MATNRS", '"|' + "|".join(物料号们) + '"')
    文本 = _vbs注入常量(文本, "DL_DIR", '"' + str(目录).replace("/", chr(92)) + '"')
    文本 = _vbs注入常量(文本, "DL_FILE", '"' + str(文件名) + '"')
    if '"|' + "|".join(物料号们) + '"' not in 文本:
        raise RuntimeError("运行时 VBS 注入失败：MATNRS 常量没被改写（Zpp_Cpgl.vbs 的 Const 行格式变了吗？）")
    ZPPCPGL_运行时.write_bytes(文本.encode("gbk", errors="replace"))
    return ZPPCPGL_运行时


def _cpgl跑脚本(脚本路径: Path) -> str:
    """cscript 同步跑运行时 VBS（与 /api/vbs/run 同一把 VBS_PROC 锁，避免两个 SAP 脚本打架）。"""
    global VBS_PROC, VBS_STOP_REQUESTED
    cscript = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "cscript.exe"
    解释器 = str(cscript) if cscript.is_file() else "cscript"
    环境 = dict(os.environ)
    环境["PYTHONIOENCODING"] = "utf-8"
    with VBS_PROC_LOCK:
        if VBS_PROC is not None and VBS_PROC.poll() is None:
            raise RuntimeError("已有 SAP 脚本在运行（VBS 进程忙），等它结束再点")
        VBS_STOP_REQUESTED = False
        try:
            VBS_PROC = subprocess.Popen(
                [解释器, "//nologo", "//E:vbscript", str(脚本路径)],
                cwd=str(VBS_SITE_BASE), env=环境,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception as exc:  # noqa: BLE001
            VBS_PROC = None
            raise RuntimeError(f"SAP 脚本启动失败：{exc}")
    try:
        原始, _ = VBS_PROC.communicate(timeout=ZPPCPGL_超时)
    except subprocess.TimeoutExpired:
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(VBS_PROC.pid)],
                           capture_output=True,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:  # noqa: BLE001
            pass
        raise RuntimeError(f"SAP 脚本超时（>{ZPPCPGL_超时}s）已强停：SAP GUI 是不是卡在弹窗或网络等待？")
    finally:
        with VBS_PROC_LOCK:
            被中断 = VBS_STOP_REQUESTED
            VBS_PROC = None
            VBS_STOP_REQUESTED = False
    原始 = 原始 or b""
    try:
        输出 = 原始.decode("utf-8")
    except UnicodeDecodeError:
        输出 = 原始.decode("gbk", errors="replace")
    if 被中断:
        raise RuntimeError("SAP 脚本被强行中断")
    return 输出


def _cpgl取本次报表(输出: str, 目录: Path, 文件名: str):
    """认「本次」产出的报表：只认 VBS 回显的 [dl] 已保存 或 本次预期文件名，**绝不回落到历史文件**。"""
    for 行 in (输出 or "").splitlines():
        if "[dl] 已保存：" in 行:
            候选 = 行.split("：", 1)[-1].strip()
            p = Path(候选)
            if p.is_file():
                return p
    预期 = Path(目录) / 文件名
    return 预期 if 预期.is_file() else None


class ZppCpglRunRequest(BaseModel):
    用剪贴板: bool = True           # 剪贴板里有物料号就用它（优先级最高）
    物料号: List[str] = []          # 手填/prepare 传回来的物料号（全号）
    下载目录: str = ""              # 留空 = 代码默认（与其它 SAPGUI 脚本同目录）
    # ↓ 2026-10-24：前端改成「点击瞬间先 prepare 写入两张表，再跑 SAP」，
    #   跑的时候沿用那批物料号与行 id（不再重复插入，也不再重读剪贴板——剪贴板可能已被改掉）
    已写入: bool = False            # True = 两张表已由 prepare 写好，本次直接跑 SAP
    写入ids: List[int] = []         # prepare 返回的「修改合同状态」新行 id（回填提货情况按它定位）
    来源: str = ""                  # prepare 判定的来源（只用于日志）


@app.get("/api/modify/zppcpgl/preview")
def zppcpgl_preview(用剪贴板: bool = Query(True)):
    """只预览「这次会拿哪些物料号、会拿什么条件去查 SAP」（不写库、不跑 SAP），给前端确认框用。

    用户 2026-10-24：「这些物料号需要先复制到修改合同状态/修改计划，然后来匹配」——
    写库前必须让人先看清是哪几个号，点错一次就是往两张业务表里插错行。
    另外 SAP 那个框是「合同+梯号」：*-*-*-03 要截成 *-*-* 当输入条件（表里仍写全号）。
    """
    物料, 来源 = _cpgl取物料(用剪贴板, [])
    return {"ok": True, "物料号": 物料, "来源": 来源, "数量": len(物料),
            "输入条件": zpp_cpgl_parse.批量截取(物料),
            "兜底说明": ("剪贴板里没有物料号，已回落到「修改代办」时间戳最新的「通知并反馈」行"
                       if 物料 and 来源.startswith("修改代办") else "")}


class ZppCpglPrepareRequest(BaseModel):
    """「点下去就写入两张表」用：只取物料号 + 前置插入，不跑 SAP。"""
    用剪贴板: bool = True
    物料号: List[str] = []


@app.post("/api/modify/zppcpgl/prepare")
def zppcpgl_prepare(payload: ZppCpglPrepareRequest):
    """**点 zpp_cpgl 的第一时间**就把物料号插到两张表最上面（用户 2026-10-24 明令）。

    用户原话：「点击 zpp_cpgl 的一瞬间，甚至直接之前，就把剪切板或者筛选出来的物料号
    复制到下面对应的 2 个标签那边对应的列，而不是等执行完以后再复制过去」。
    所以拆成两步：prepare（本接口，毫秒级写入 + 前端立刻刷新表格）→ run（慢慢跑 SAP）。
    返回的 写入ids 要原样带回 run，回填「提货情况」时按 id 精确定位这几行。
    「修改计划」按用户 2026-10-25 的要求**先清空再插**，清空掉的行数放在 清空 里回给前端提示。
    """
    if not ZPPCPGL_VBS.is_file():
        raise HTTPException(status_code=500, detail=f"VBS 脚本不存在：{ZPPCPGL_VBS}")
    with ZPPCPGL_锁:
        if ZPPCPGL_任务["running"]:
            raise HTTPException(status_code=409,
                                detail=f"已有 zpp_cpgl 任务在跑（job {ZPPCPGL_任务['job_id']}），等它结束再点")
    物料, 来源 = _cpgl取物料(payload.用剪贴板, payload.物料号)
    if not 物料:
        raise HTTPException(status_code=400,
                            detail="没有可用的物料号：剪贴板里没有，修改代办里也没有「通知并反馈」的行")
    写入, 清空 = _cpgl前置插入(物料, 来源)
    UI信息(f"zpp_cpgl 前置写入：{len(物料)} 个物料号 → 「修改合同状态」插最上面、"
           f"「修改计划」先清空再插（清掉 {清空.get('修改计划', 0)} 行）（{来源}）")
    return {"ok": True, "物料号": 物料, "来源": 来源, "数量": len(物料),
            "输入条件": zpp_cpgl_parse.批量截取(物料),
            "清空": 清空,
            "写入ids": {"修改合同状态": 写入["修改合同状态"], "修改计划": 写入["修改计划"]}}


@app.post("/api/modify/zppcpgl/run")
def zppcpgl_run(payload: ZppCpglRunRequest):
    """启动一次「zpp_cpgl」：取物料号（或沿用 prepare 已写入的）→ 跑 SAP → 解析报表 → 回填提货情况。"""
    if not ZPPCPGL_VBS.is_file():
        raise HTTPException(status_code=500, detail=f"VBS 脚本不存在：{ZPPCPGL_VBS}")
    with ZPPCPGL_锁:
        if ZPPCPGL_任务["running"]:
            raise HTTPException(status_code=409,
                                detail=f"已有 zpp_cpgl 任务在跑（job {ZPPCPGL_任务['job_id']}），等它结束")
        job = uuid.uuid4().hex[:12]
        ZPPCPGL_任务.update({"job_id": job, "running": True, "lines": [], "result": None,
                             "error": "", "开始": datetime.now(), "结束": None})
    threading.Thread(target=_cpgl任务线程, args=(job, payload), daemon=True).start()
    return {"ok": True, "job_id": job}


@app.get("/api/modify/zppcpgl/status")
def zppcpgl_status(job_id: str = Query(""), after: int = Query(0)):
    """任务状态 + 增量日志（after 游标）。跑完返回 result / error。"""
    with ZPPCPGL_锁:
        if job_id and ZPPCPGL_任务["job_id"] != job_id:
            return {"ok": False, "running": False, "lines": [], "next": after,
                    "error": "任务不存在（服务可能重启过）"}
        全部 = list(ZPPCPGL_任务["lines"])
        起 = max(0, min(int(after or 0), len(全部)))
        return {"ok": True, "running": bool(ZPPCPGL_任务["running"]),
                "job_id": ZPPCPGL_任务["job_id"],
                "lines": 全部[起:], "next": len(全部),
                "result": ZPPCPGL_任务["result"], "error": ZPPCPGL_任务["error"],
                "开始": ZPPCPGL_任务["开始"].strftime("%H:%M:%S") if ZPPCPGL_任务["开始"] else "",
                "结束": ZPPCPGL_任务["结束"].strftime("%H:%M:%S") if ZPPCPGL_任务["结束"] else ""}


def _cpgl任务线程(job: str, payload: ZppCpglRunRequest) -> None:
    """真正的执行体（daemon 线程）：任何异常都落成 error，绝不让任务永远 running。"""
    开始 = datetime.now()
    try:
        _cpgl日志("① 取物料号（优先级：剪贴板 > 修改代办「通知并反馈」最新时间戳）")
        if payload.已写入 and payload.写入ids:
            # 前端点下去就已经 prepare 写好了两张表 → 沿用那批（不重复插入，也不重读剪贴板）
            物料 = _cpgl规范物料列表(payload.物料号)
            if not 物料:
                raise RuntimeError("调用方式不一致：已写入=True 但没带物料号")
            来源 = payload.来源 or "点击时已写入两张表"
            写入 = {"修改合同状态": [int(x) for x in payload.写入ids], "修改计划": []}
            写入统计 = {"修改合同状态": len(payload.写入ids), "修改计划": len(payload.写入ids)}
            _cpgl日志(f"② 沿用点击时写入的 {len(物料)} 个物料号（来源：{来源}）：" + "、".join(物料))
        else:
            物料, 来源 = _cpgl取物料(payload.用剪贴板, payload.物料号)
            if not 物料:
                raise RuntimeError("没有可用的物料号：剪贴板里没有，修改代办里也没有「通知并反馈」的行")
            _cpgl日志(f"② 共 {len(物料)} 个（来源：{来源}）：" + "、".join(物料))

        # SAP 那个框是「合同+梯号」：*-*-*-03 要截成 *-*-* 当输入条件（用户 2026-10-24 明令）。
        # 表里写的、报表里匹配的仍是全号 —— 否则同合同的 -01/-02/-03 会串行。
        输入条件 = zpp_cpgl_parse.批量截取(物料)
        _cpgl日志(f"   SAP 输入条件（合同+梯号，已去掉末尾箱号）{len(输入条件)} 个：" + "、".join(输入条件))

        if payload.已写入 and payload.写入ids:
            _cpgl日志(f"③ 物料号在点「zpp_cpgl」时就已写进两张表（{len(写入['修改合同状态'])} 行），本次不重复写入")
        else:
            写入, 清空 = _cpgl前置插入(物料, 来源)
            写入统计 = {"修改合同状态": len(写入["修改合同状态"]), "修改计划": len(写入["修改计划"])}
            _cpgl日志(f"③ 已写进两张表：「修改合同状态」插最上面 {写入统计['修改合同状态']} 行（原有数据未动）、"
                      f"「修改计划」先清空 {清空.get('修改计划', 0)} 行再插 {写入统计['修改计划']} 行")

        目录 = _sapq解析下载目录(payload.下载目录)
        文件名 = _sapq报表文件名(ZPPCPGL_键)
        _cpgl日志(f"④ 生成运行时 VBS（输入 {len(输入条件)} 个）→ 执行 /nzpp_cpgl → 报表 {目录}\\{文件名}")
        运行时 = _cpgl生成运行时(输入条件, 目录, 文件名)
        输出 = _cpgl跑脚本(运行时)
        # VBS 输出只留关键行进任务日志（[dl] / [Error] / [result] / [fill] …），其余不刷屏
        关键 = [l.strip() for l in (输出 or "").splitlines()
                if l.strip().startswith(("[dl]", "[Error]", "[error]", "[result]", "[fill]", "[warn]", "[Step]"))]
        for l in 关键[:60]:
            _cpgl日志("   · " + l)

        报表 = _cpgl取本次报表(输出, 目录, 文件名)
        if 报表 is None:
            尾部 = " / ".join((输出 or "").splitlines()[-6:])[:500]
            raise RuntimeError(f"本次没有生成报表（{目录}\\{文件名}）。SAP 输出尾部：{尾部}")
        _cpgl日志(f"⑤ 报表已落盘：{报表}")

        报 = zpp_cpgl_parse.解析报表(报表)
        _cpgl日志(f"   报表解析：表头 {len(报['表头'])} 列、数据 {len(报['行'])} 行、"
                  f"列组 {len(报['列组'])} 个（表头在第 {报['表头行号'] + 1} 行）")
        # 报表自身有没有可回填的出入库数据，是「没匹配上」还是「匹配上了但这批没数据」的分水岭
        有值行 = [r for r in 报["行"] if r["入库情况"]]
        _cpgl日志(f"   报表里「从右往左有值列组」可回填的行：{len(有值行)}/{len(报['行'])}")
        匹配 = zpp_cpgl_parse.批量匹配(报, 物料)
        _提填, 入填 = _cpgl回填提货情况(写入["修改合同状态"], 匹配)
        无提货 = [m["物料号"] for m in 匹配 if m["状态"] == "无提货数据"]
        未匹配 = [m["物料号"] for m in 匹配 if m["状态"] == "未匹配"]
        _cpgl日志(f"⑥ 回填「入库情况」{入填}/{len(物料)} 行（「提货情况」按用户明令留空）")
        if 无提货:
            _cpgl日志(f"   · 命中但报表里这批没有可回填数据（中转/成品列全空）{len(无提货)} 个：{'、'.join(无提货)}")
        if 未匹配:
            _cpgl日志(f"   · 报表里根本没有这些号 {len(未匹配)} 个：{'、'.join(未匹配)}")
        for m in 匹配:
            if m["入库情况"]:
                _cpgl日志(f"   · {m['物料号']} → 入库情况 {m['入库情况']}（{m['命中方式']}命中{('，共 ' + str(m['命中行数']) + ' 行') if m['命中行数'] > 1 else ''}）")

        # ⑦ 只生成报表，不自动打开（用户 2026-11-06：「下载完以后没有人让你总自动打开，逻辑有问题」）。
        # 2026-10-25 曾按当时要求做成「打开后停留 2 秒再自动关闭」，现按最新指令整体撤掉。
        # SAPQ 查询页的「自动打开」开关是另一条链，一行不动；这里只管 zpp_cpgl。
        打开 = ""
        _cpgl日志("⑦ 报表已生成（按用户明令不自动打开）")

        结果 = {
            "物料号": 物料, "来源": 来源, "输入条件": 输入条件,
            "写入": 写入统计,
            "报表": {"文件": 报表.name, "路径": str(报表), "行数": len(报["行"]),
                     "有提货数据的行": len(有值行)},
            "匹配": 匹配, "已回填": 入填, "已入库": 入填, "未匹配": 未匹配, "无提货数据": 无提货,
            "自动打开": 打开,
            "耗时": round((datetime.now() - 开始).total_seconds(), 1),
        }
        _cpgl日志(f"✔ 完成：{入填}/{len(物料)} 个已填入库情况（提货情况留空），报表 {报表.name}")
        with ZPPCPGL_锁:
            ZPPCPGL_任务["result"] = 结果
            ZPPCPGL_任务["running"] = False
            ZPPCPGL_任务["结束"] = datetime.now()
    except Exception as exc:  # noqa: BLE001
        _cpgl日志(f"[失败] {type(exc).__name__}: {exc}")
        with ZPPCPGL_锁:
            ZPPCPGL_任务["error"] = str(exc) or type(exc).__name__
            ZPPCPGL_任务["running"] = False
            ZPPCPGL_任务["结束"] = datetime.now()


def _vbs最新日志(脚本键: str, 最大: int = 6000) -> str:
    """最近一次该脚本的运行日志尾部（供「发邮件」单独发送时拼邮件正文/状态栏）。"""
    try:
        if not VBS_LOG_DIR.is_dir():
            return ""
        候选 = sorted([p for p in VBS_LOG_DIR.glob(f"*_{脚本键}.log") if p.is_file()],
                      key=lambda x: x.stat().st_mtime, reverse=True)
        if not 候选:
            return ""
        return 候选[0].read_text(encoding="utf-8", errors="replace")[-最大:]
    except Exception:
        return ""


@app.get("/api/vbs/sapq_config")
def vbs_sapq_config():
    """查询页按钮区需要的配置（供前端显示「下载路径」框默认值）。

    默认路径以代码为准（SAPQ_下载目录），前端把它填成输入框 placeholder；
    用户在框里写了自定义路径 → 以自定义为准（后端 _sapq解析下载目录）。
    """
    _dws = _rpa导dws()
    默认收件人 = str(getattr(_dws, "默认收件人", "") or "") if _dws else ""
    if not 默认收件人:      # dws 不可用时按同一口径兜底（与 dws_mail.py 的默认值一致）
        默认收件人 = os.environ.get("APS_MAIL_TO", "demo.user67@example.com")
    return {"ok": True,
            "下载目录": str(SAPQ_下载目录),
            "默认收件人": 默认收件人,
            "报表后缀": SAPQ_报表后缀,
            "报表前缀": dict(SAPQ_报表前缀),
            "页面": {键: {"显示名": 值[1], "前缀": SAPQ_报表前缀.get(键, 键)}
                    for 键, 值 in VBS_SCRIPTS_SITE.items()}}


class VbsDownloadRequest(BaseModel):
    """「报表下载」单独点：不重跑查询，直接把当前 SAP 屏幕上的报表导出到 web_xls。"""
    script: str = "zb02"
    params: Optional[dict] = None      # 页面现值（跳过查询模式下只作记账，不写 SAP）
    报表文件: str = ""                 # 留空 = 功能名_时间戳.xls（后端生成）
    下载目录: str = ""                 # 页面「下载路径」框：留空 = 代码默认；自定义优先


class VbsSendMailRequest(BaseModel):
    """「发邮件」单独点：独立发邮件功能（不碰 SAP），把最近一次执行结果发出去。"""
    script: str = "zb02"
    收件人: str = ""                   # 多个用 , / ; / 空格分隔；留空 = dws 默认
    主题: Optional[str] = None
    附报表: bool = True                # 把下载目录里最新的报表作为普通附件一并发送
    下载目录: str = ""                 # 页面「下载路径」框：留空 = 代码默认；自定义优先


@app.post("/api/vbs/download")
def vbs_download(payload: VbsDownloadRequest):
    """「报表下载」按钮（2026-10-23 用户需求，独立于执行查询）：

    跳过查询（SKIP_QUERY=1）→ 在当前 SAP 屏幕上走「系统 → 清单 → 保存 → 本地文件」链，
    文件名 = 功能名 + 时间戳，落点 = 页面「下载路径」框（自定义优先），留空 = 代码默认目录。
    """
    script = (payload.script or "").strip()
    if script not in VBS_SCRIPTS_SITE:
        raise HTTPException(status_code=400,
                            detail=f"报表下载仅支持查询页（{'、'.join(VBS_SCRIPTS_SITE)}）")
    结果 = vbs_run(VbsRunRequest(script=script, params=payload.params or {},
                                 mail=False, download=True, skipQuery=True,
                                 下载目录=payload.下载目录))
    UI信息(f"报表下载（独立）：{结果.get('download')}")
    UI信息(f"报表下载（独立）：{'成功' if (结果.get('download') or {}).get('ok') else '未生成'}")
    return 结果


@app.post("/api/vbs/send_mail")
def vbs_send_mail(payload: VbsSendMailRequest):
    """「发邮件」按钮（2026-10-23 用户需求，独立于执行查询）：

    不碰 SAP。内容 = 该页最近一次的多屏截图（内联）+ 最近一次运行日志（状态栏）；
    若下载目录里有报表产物，则把最新的 <功能名>_*.xls 作为普通附件一并发送
    （目录 = 页面「下载路径」框，自定义优先；留空 = 代码默认目录）。
    """
    script = (payload.script or "").strip()
    站点键 = script if script in VBS_SCRIPTS_SITE else None
    if not 站点键 and script != "co28":
        raise HTTPException(status_code=400, detail=f"未知脚本 {script!r}")
    输出 = _vbs最新日志(script)
    下载目录 = _sapq解析下载目录(payload.下载目录) if 站点键 else SAPQ_下载目录
    附件 = []
    if payload.附报表 and 站点键:
        最新报表 = _sapq取最新报表(站点键, 下载目录)
        if 最新报表:
            附件 = [最新报表]
    if script == "co28":
        结果 = _co28发送结果邮件(True, False, 输出, datetime.now(), 收件人=payload.收件人)
    else:
        结果 = _发送查询结果邮件(站点键, True, False, 输出, datetime.now(), {},
                                收件人=payload.收件人, 附件=附件, 下载目录=下载目录,
                                报表回落=bool(payload.附报表))
    主题覆盖 = str(payload.主题 or "").strip()
    return {"ok": bool(结果.get("已发")), "结果": 结果, "主题覆盖": 主题覆盖,
            "附件": [p.name for p in 附件], "日志": bool(输出)}


@app.post("/api/vbs/co28_failmail")
def vbs_co28_failmail(payload: Co28FailMailRequest):
    """CO28 执行失败时的失败邮件（用户 2026-10-23 需求）：
    正文含三部分截图素材——① 执行到的页面（SAP 当前屏幕，后端抓）② 当前网站页面（前端 html2canvas 回传）
    ③ 日志记录（vbs_logs 最新 co28 日志 + 前端附带输出）。收件人固定 demo.user68@example.com。
    """
    dws = _rpa导dws()
    if dws is None:
        raise HTTPException(status_code=500,
                            detail="dws_mail 导入失败（web_auto/公共_钉钉通知 拷贝是否完整）")
    图 = []
    sap图 = _co28抓当前SAP屏()
    if sap图:
        图.append(("sap", sap图))
    web图 = _co28存网页截图(payload.网页截图 or "")
    if web图:
        图.append(("web", web图))
    日志 = (payload.日志 or '').strip() or (payload.输出 or '').strip() or _co28日志尾部()
    行 = ["CO28 下达执行失败，请人工核查。", "",
          f"时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
          "", "① 执行到的页面（SAP 当前屏幕）"]
    行.append("[inline:" + (sap图.name if sap图 else "无") + "]" if sap图 else "（未取到 SAP 截图）")
    行 += ["", "② 当前网站页面"]
    行.append("[inline:" + (web图.name if web图 else "无") + "]" if web图 else "（未取到网页截图）")
    行 += ["", "③ 日志记录", "```", (日志 or "（无日志）")[-6000:], "```"]
    正文 = "\n".join(行)
    # 2026-10-31：缩放后先暂存副本再发 —— 防止 dws 重试期间 co28_shots 被并发运行清空
    # （今天现场：截图 11:53:28 落盘，新运行 11:53:32 清目录，dws 11:53:34 重试时文件已没了）
    附件 = _co28暂存邮件图片([_co28缩放图片(p) for _, p in 图])
    收件 = _sapq解析收件人(payload.收件人)
    收件人文本 = ", ".join(收件) if 收件 else ""
    try:
        kwargs = {"内联图片": 附件}
        if 收件人文本:
            kwargs["收件人"] = 收件人文本
        状态, 说明 = dws.发送邮件(payload.主题 or CO28_MAIL_主题失败, 正文, **kwargs)
    except Exception as exc:
        状态, 说明 = "彻底失败", f"发送异常：{exc}"
    UI信息(f"CO28 失败邮件：{状态}｜{说明}｜截图 {len(附件)} 张｜收件人 {收件人文本 or '（默认）'}")
    return {"ok": 状态 == "已发送", "状态": 状态, "说明": 说明,
            "截图数": len(附件), "收件人": 收件人文本 or "demo.user69@example.com（默认）"}


# ============ 数据源：上传 / 重载 ============

@app.post("/api/tools/upload")
async def 上传文件(file: UploadFile = File(...), sheet: Optional[str] = Form(None),
                    表头行号: Optional[int] = Form(None),
                    header_row: Optional[int] = Form(None)):
    """上传 Excel 文件 → 原版 读取_文件() 解析 → 返回列信息与预览。

    字段兼容：新前端传 header_row，旧版工具页传 表头行号（二者等价）。
    """
    行号 = 表头行号 or header_row or 1
    文件名 = file.filename or "未命名.xlsx"
    后缀 = Path(文件名).suffix.lower()
    if 后缀 not in 支持的后缀:
        raise HTTPException(
            status_code=400,
            detail=f"网页版仅支持 .xlsx / .xlsm（{后缀 or '无扩展名'} 请先在 Excel 另存为 .xlsx）")
    try:
        内容 = await file.read()
        保存路径 = UPLOAD_DIR / f"{uuid.uuid4().hex[:12]}{后缀}"
        with 保存路径.open("wb") as f:
            f.write(内容)
        ds = _io.读取_文件(str(保存路径), sheet or None, 行号)
        ds.工作簿名 = 文件名        # 展示与下载用原始文件名（读取时走的是临时保存路径）
    except HTTPException:
        raise
    except Exception as 异常:
        raise HTTPException(status_code=400, detail=f"文件解析失败：{异常}")
    if ds.数据框.empty:
        raise HTTPException(status_code=400, detail="没有读到有效数据，请检查文件内容与表头行")
    ds_id = _存入(ds, str(保存路径), 文件名)
    表名列表: List[str] = []
    try:
        表名列表 = _io.列出文件表名(str(保存路径))
    except Exception:
        pass
    结果 = _摘要(ds_id, _取(ds_id))
    结果["sheet_names"] = 表名列表
    UI信息(f"数据已加载：{ds.摘要()}")
    return 结果


@app.post("/api/tools/reload")
def 重载数据(payload: dict):
    """Sheet / 表头行变更后按原路径重读（复刻原版切换 Sheet 的行为）。"""
    ds_id = payload.get("dataset_id")
    sheet = payload.get("sheet") or None
    header_row = int(payload.get("header_row") or 1)
    项 = _取(ds_id)
    if not 项.get("path"):
        raise HTTPException(status_code=400, detail="该数据集无关联文件，无法重载")
    try:
        ds = _io.读取_文件(项["path"], sheet, header_row)
    except Exception as 异常:
        raise HTTPException(status_code=400, detail=f"重新读取失败：{异常}")
    新id = _存入(ds, 项["path"], 项.get("name"))
    UI信息(f"数据已加载：{ds.摘要()}")
    return _摘要(新id, _取(新id))


# ============ 工具一：VLOOKUP（原版 021_vlookup.vlookup_合并）============

@app.post("/api/tools/vlookup")
def vlookup接口(payload: dict):
    左id = payload.get("left_id")
    右id = payload.get("right_id")
    if not 右id:
        raise HTTPException(status_code=400, detail="VLOOKUP 需要先导入右表")
    左项 = _取(左id)
    右项 = _取(右id)
    原列数 = 左项["ds"].列数
    try:
        新ds = _proc.vlookup_合并(
            左项["ds"], 右项["ds"],
            左键列=payload.get("左键列", ""),
            右键列=payload.get("右键列", ""),
            取值列列表=payload.get("取值列列表") or [],
            模式=payload.get("模式") or "全部合并",
            合并分隔符=payload.get("合并分隔符") or ",",
            去重=bool(payload.get("去重", True)),
            新列名前缀=payload.get("新列名前缀") or "")
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    except Exception as 异常:
        raise HTTPException(status_code=400, detail=f"匹配失败：{异常}")
    新id = _存入(新ds)
    UI信息("VLOOKUP 执行完成")
    return _执行结果(新id, _取(新id), 原列数, "VLOOKUP 执行完成")


# ============ 工具二：正则提取（原版 022_regex_extract / 025_param_scanner）============

@app.post("/api/tools/scan-params")
def 扫描参数接口(payload: dict):
    """扫描源列"奇数位=参数名"，返回候选参数列表（原版 025.扫描参数）。"""
    ds_id = payload.get("dataset_id")
    源列 = (payload.get("源列") or "").strip()
    if not 源列:
        raise HTTPException(status_code=400, detail="请先填写 源列（如：Y）")
    项 = _取(ds_id)
    try:
        参数 = _proc.扫描参数(项["ds"], 源列)
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    UI信息(f"扫描完成：列={源列}，候选参数 {len(参数)} 个")
    return {"params": 参数, "count": len(参数),
            "message": f"扫描完成：列={源列}，候选参数 {len(参数)} 个"}


@app.post("/api/tools/regex-extract")
def 正则提取接口(payload: dict):
    """自定义正则提取（原版 022.正则提取）。"""
    ds_id = payload.get("dataset_id")
    项 = _取(ds_id)
    原列数 = 项["ds"].列数
    源列 = (payload.get("源列") or "").strip()
    if not 源列:
        raise HTTPException(status_code=400, detail="请先填写 源列（如：Y）")
    目标列名 = (payload.get("目标列名") or "").strip() or None
    覆盖同名列 = bool(payload.get("覆盖同名列", True))
    UI信息(f"正则提取：当前模式=自定义正则 | 源列={源列} | "
           f"目标列名={目标列名 or '自动'} | 覆盖同名列={覆盖同名列}")
    try:
        新ds = _proc.正则提取(
            项["ds"], 源列=源列, 正则=payload.get("正则") or "",
            目标列名=目标列名, 捕获组=int(payload.get("捕获组") or 1),
            全部命中=bool(payload.get("全部命中", False)),
            合并分隔符=payload.get("合并分隔符") or ",",
            覆盖同名列=覆盖同名列)
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    新id = _存入(新ds)
    新列 = 目标列名 or 新ds.列名列表[-1]
    UI信息(f"正则提取执行完成 → 已写入列：{新列}")
    return _执行结果(新id, _取(新id), 原列数, f"正则提取执行完成 → 已写入列：{新列}")


@app.post("/api/tools/param-extract")
def 按参数提取接口(payload: dict):
    """按参数提取（原版 025.按参数提取，5 种模式）。"""
    ds_id = payload.get("dataset_id")
    项 = _取(ds_id)
    原列数 = 项["ds"].列数
    源列 = (payload.get("源列") or "").strip()
    if not 源列:
        raise HTTPException(status_code=400, detail="请先填写 源列（如：Y）")
    参数模式 = payload.get("参数模式") or "单个参数"
    参数列表 = payload.get("参数列表") or []
    目标列名 = (payload.get("目标列名") or "").strip() or None
    覆盖同名列 = bool(payload.get("覆盖同名列", True))
    # 复刻原版主窗口的前置校验：单个/任一/全部 模式需选择参数
    if 参数模式 in (_proc.模式_单个, _proc.模式_任一, _proc.模式_全部) and not 参数列表:
        raise HTTPException(status_code=400, detail=f"当前模式《{参数模式}》需选择参数")
    UI信息(f"正则提取：当前模式=按参数({参数模式}) | 源列={源列} | "
           f"目标列名={目标列名 or '自动'} | 覆盖同名列={覆盖同名列}")
    try:
        新ds = _proc.按参数提取(
            项["ds"], 源列=源列, 参数列表=参数列表,
            模式=参数模式, 合并分隔符=payload.get("合并分隔符") or ",",
            覆盖同名列=覆盖同名列, 目标列名=目标列名,
            输出对应值=bool(payload.get("输出对应值", False)))
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    新id = _存入(新ds)
    新列 = 目标列名 or (新ds.列名列表[-1] if 新ds.列数 else "（多列扩展）")
    UI信息(f"正则提取执行完成 → 已写入列：{新列}")
    return _执行结果(新id, _取(新id), 原列数, f"正则提取执行完成 → 已写入列：{新列}")


@app.post("/api/tools/ai-regex")
def AI生成正则(payload: dict):
    """AI 自然语言 → 正则（本地规则优先，未命中走 DeepSeek）。"""
    描述 = (payload.get("描述") or "").strip()
    模型 = (payload.get("模型") or "deepseek-v4-pro").strip()
    密钥 = (payload.get("密钥") or "").strip()
    try:
        正则, 来源 = ai_regex.生成正则(描述, 模型, 密钥)
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    UI信息(f"AI 正则生成完成（{来源}）：{正则}")
    return {"regex": 正则, "source": 来源, "message": f"已按{来源}回填"}


# ============ 工具三：替换（原版 023_replace）============

@app.post("/api/tools/replace")
def 替换接口(payload: dict):
    ds_id = payload.get("dataset_id")
    项 = _取(ds_id)
    原列数 = 项["ds"].列数
    模式 = payload.get("模式") or "单对"
    # 适配层：先把"作用列"解析为真实列名（空 = 全部文本列，兼容 StringDtype），
    # 再交给原版处理器——原版 023._应用到列列表 的 dtype == object 在新版
    # pandas 下会漏掉 StringDtype 字符串列。
    实际列列表 = _解析作用列(项["ds"].数据框, payload.get("列列表") or None)
    if not 实际列列表:
        UI错误("查找替换：没有可用列，原样返回")
        raise HTTPException(status_code=400, detail="没有可处理的文本列（请检查作用列填写）")
    try:
        if 模式 == "对照表":
            对照表 = payload.get("对照表") or {}
            if not isinstance(对照表, dict) or not 对照表:
                raise ValueError("对照表为空，请按「每行一对」填写")
            新ds = _proc.对照表替换(
                项["ds"], 列列表=实际列列表,
                对照表=对照表, 整格匹配=bool(payload.get("整格匹配", True)),
                使用正则=bool(payload.get("使用正则", False)))
        else:
            新ds = _proc.查找替换(
                项["ds"], 列列表=实际列列表,
                查找文本=payload.get("查找文本") or "",
                替换文本=payload.get("替换文本") or "",
                使用正则=bool(payload.get("使用正则", False)),
                区分大小写=bool(payload.get("区分大小写", True)),
                全词匹配=bool(payload.get("全词匹配", False)))
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    新id = _存入(新ds)
    UI信息("替换执行完成")
    return _执行结果(新id, _取(新id), 原列数, "替换执行完成")


# ============ 工具四：分列（原版 024_split_columns）============

@app.post("/api/tools/split")
def 分列接口(payload: dict):
    ds_id = payload.get("dataset_id")
    项 = _取(ds_id)
    原列数 = 项["ds"].列数
    模式 = payload.get("模式") or "按分隔符"
    try:
        if 模式 == "按正则":
            新ds = _proc.按正则分列(
                项["ds"], 源列=payload.get("源列") or "",
                正则=payload.get("正则") or "",
                新列名列表=payload.get("新列名列表") or None,
                保留原列=bool(payload.get("保留原列", True)))
        else:
            新ds = _proc.按分隔符分列(
                项["ds"], 源列=payload.get("源列") or "",
                分隔符=payload.get("分隔符") or ",",
                最大列数=int(payload.get("最大列数") or 0),
                新列名前缀=payload.get("新列名前缀") or "",
                保留原列=bool(payload.get("保留原列", True)),
                去除空白=bool(payload.get("去除空白", True)))
    except ValueError as 异常:
        raise HTTPException(status_code=400, detail=str(异常))
    新id = _存入(新ds)
    UI信息("分列执行完成")
    return _执行结果(新id, _取(新id), 原列数, "分列执行完成")


# ============ 输出：下载 xlsx（原版 012_excel_writer.写出_新建文件）============

@app.get("/api/tools/download/{ds_id}")
def 下载结果(ds_id: str, filename: Optional[str] = Query(None),
             后台任务: BackgroundTasks = None):
    项 = _取(ds_id)
    ds = 项["ds"]
    名称 = filename or 项.get("name") or "结果"
    if not 名称.lower().endswith((".xlsx", ".xlsm")):
        名称 = f"{名称}.xlsx"
    try:
        临时路径 = TMP_DIR / f"export_{uuid.uuid4().hex[:8]}.xlsx"
        _io.写出_新建文件(ds, str(临时路径))
    except Exception as 异常:
        raise HTTPException(status_code=500, detail=f"导出失败：{异常}")
    UI信息(f"已导出到：{名称}")
    if 后台任务 is not None:
        后台任务.add_task(lambda: 临时路径.unlink(missing_ok=True))
    return FileResponse(
        str(临时路径),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename*=UTF-8''{urllib.parse.quote(名称)}"})


# ============ APD 看板共用数据管线（MySQL sap_table.03_apd） ============

APD_DB = dict(host="127.0.0.1", port=3306, user="u3854",
              password="u38544253", database="sap_table", charset="utf8mb4")
_APD_CACHE = {"df": None, "ts": 0.0}
_APD_TTL = 120.0
_APD_LOCK = threading.Lock()
_APD_DATE_PAT = _re.compile(r"(20\d{2}\.\d{2}\.\d{2})")
_APD_BG_PAT = _re.compile(r"BG(\d+)")


def _apd_load_df() -> pd.DataFrame:
    """读取 03_apd 全表并派生口径列（下达日期/角标/梯种），120s 缓存。"""
    import time as _time
    now = _time.time()
    with _APD_LOCK:
        if _APD_CACHE["df"] is not None and now - _APD_CACHE["ts"] < _APD_TTL:
            return _APD_CACHE["df"]
        import pymysql
        conn = pymysql.connect(**APD_DB)
        try:
            df = pd.read_sql(
                "SELECT `合同箱号`,`箱子完工日期`,`安全钳组件物料`,"
                "`框架组件物料`,`框架属性`,`框架参数`,`来源文件夹` FROM `03_apd`", conn)
        finally:
            conn.close()
        def _date_of(s):
            m = _APD_DATE_PAT.search(str(s or ""))
            return m.group(1) if m else ""
        df["下达日期"] = df["来源文件夹"].map(_date_of)
        df = df[df["下达日期"] != ""].copy()
        def _corner(s):
            parts = str(s or "").split("-")
            return parts[-1] if parts and parts[-1] else ""
        df["角标"] = df["合同箱号"].map(_corner)
        df["梯种"] = df["合同箱号"].astype(str).str[2:5]
        df["框架属性"] = df["框架属性"].fillna("").astype(str).str.strip().str.upper()
        df["BG参数"] = df["框架参数"].fillna("").map(
            lambda s: (m.group(1) if (m := _APD_BG_PAT.search(str(s))) else ""))
        _APD_CACHE.update(df=df, ts=now)
        return df


def _apd_norm_code(v) -> str:
    """图号/物料归一：大写；物料编码去掉 DM 前缀（DMP125... → P125...）。"""
    v = str(v or "").strip().upper()
    if v.startswith("DM") and len(v) > 10:
        v = v[2:]
    return v


_APD_MAP_CACHE = {}   # {path: ((mtime_ns, size), 结果)} —— 分类 JSON 按 mtime 缓存，免重复读盘+排序


def _apd_local_store() -> dict:
    p = APP_DIR / "safety_store" / "local_store.json"
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _apd_load_frame_mapping() -> list:
    """安全钳框架配置分类（app/safety_store/local_store.json）：图号→分类。按文件 mtime 缓存。"""
    p = APP_DIR / "safety_store" / "local_store.json"
    try:
        st = p.stat()
        stamp = (st.st_mtime_ns, st.st_size)
    except OSError:
        return []
    ent = _APD_MAP_CACHE.get(str(p) + "#frame")
    if ent and ent[0] == stamp:
        return ent[1]
    try:
        store = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []
    rows = (store.get("sheets") or {}).get("安全钳框架配置分类") or []
    out = []
    for r in rows:
        d = r.get("data") or {}
        g = _apd_norm_code(d.get("图号"))
        if g:
            out.append((g, str(d.get("分类") or "").strip() or "无"))
    # 长图号优先（避免短图号抢走长图号的前缀行）
    out.sort(key=lambda x: -len(x[0]))
    _APD_MAP_CACHE[str(p) + "#frame"] = (stamp, out)
    return out


def _apd_load_mapping() -> list:
    """安全钳总图号分类（app/safety_store/local_store.json）：图号→系列→TYP型号。按文件 mtime 缓存。"""
    p = APP_DIR / "safety_store" / "local_store.json"
    try:
        st = p.stat()
        stamp = (st.st_mtime_ns, st.st_size)
    except OSError:
        return []
    ent = _APD_MAP_CACHE.get(str(p) + "#total")
    if ent and ent[0] == stamp:
        return ent[1]
    try:
        store = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []
    rows = (store.get("sheets") or {}).get("安全钳总图号分类") or []
    out = []
    for r in rows:
        d = r.get("data") or {}
        g = _apd_norm_code(d.get("图号"))
        if g:
            out.append((g, str(d.get("系列") or "").strip() or "无",
                        str(d.get("TYP型号") or "").strip() or "无"))
    # 长图号优先（P125052A000P7703 应先于 P125052A000 被尝试）
    out.sort(key=lambda x: -len(x[0]))
    _APD_MAP_CACHE[str(p) + "#total"] = (stamp, out)
    return out


def _apd_match(vals: pd.Series, mapping: list, prefix_mode: bool):
    """按映射表匹配系列/TYP/图号（2026-10-16 修正：完整图号前缀唯一匹配，长图号优先）。
    安全钳组件 → 总图号分类（三元组 图号/系列/TYP）；框架组件 → 框架配置分类（二元组 图号/分类，
    分类值同时写入系列与 TYP 两列，透视/旭日里 系列=分类、型号=分类、图号=具体图号）。
    归一后 startswith 完整图号；一行只归一个图号（先到先得），未命中→无。
    旧逻辑缺陷：安全钳用"字面包含"（YA107A18 误吞 YA107A180/181）、框架用"前8位前缀"
    （P125022A000 误吞 P125022A001/A104 的行）。
    2026-10-18 性能：mapping 已按图号长度降序 → 完整图号必然是最长可能前缀、一定排在
    其他可命中前缀之前，故先用 dict 精确命中（向量化 map），仅未命中行再走前缀扫描。"""
    norm = _apd_norm_code
    res_s, res_t, res_g = [], [], []
    # 精确命中字典（重复图号保留首行，与"先到先得"一致）
    exact = {}
    for item in mapping:
        exact.setdefault(item[0], item)
    # 第一轮：向量化精确匹配（绝大多数行在此解决）
    vns = [norm(v) for v in vals]
    hits = [exact.get(vn) if vn else None for vn in vns]
    # 第二轮：仅未命中行走前缀扫描（mapping 已按长度降序，先到先得）
    miss_idx = [i for i, h in enumerate(hits) if h is None and vns[i]]
    for i in miss_idx:
        vn = vns[i]
        for item in mapping:
            if vn.startswith(item[0]):
                hits[i] = item
                break
    for i, vn in enumerate(vns):
        item = hits[i]
        if item is not None:
            if len(item) >= 3:
                s, t, g = item[1], item[2], item[0]
            else:
                s = t = item[1]
                g = item[0]
        else:
            s = t = g = "无"
        res_s.append(s); res_t.append(t); res_g.append(g)
    return pd.Series(res_s, index=vals.index), pd.Series(res_t, index=vals.index), pd.Series(res_g, index=vals.index)


def _apd_tree(sub: pd.DataFrame, cols: list, top: int = 10):
    """构建 sunburst 树：颗粒度从左到右递减，超 top 的归「其他」。"""
    if sub.empty:
        return []
    tree = []
    if sub.empty:
        return tree
    # 2026-10-18 性能：原实现逐节点 sub[sub[col]==k] 反复全表过滤（字符串比较，160k 行×百节点=1.5s+），
    # 改为单次 groupby(cols) 后按 MultiIndex 逐层 xs/聚合取数，结果与原实现逐项一致。
    g = sub.groupby(cols, dropna=True)["合同箱号"].count()
    g1 = g.groupby(level=0).sum().sort_values(ascending=False)
    head1 = list(g1.head(top).index)
    for k1 in head1:
        g1sub = g.xs(k1, level=0)
        node = {"name": k1 or "无", "value": int(g1[k1])}
        if len(cols) > 1:
            children = []
            g2 = g1sub.groupby(level=0).sum().sort_values(ascending=False)
            head2 = list(g2.head(8).index)
            for k2 in head2:
                g2sub = g1sub.xs(k2, level=0)
                n2 = {"name": k2 or "无", "value": int(g2[k2])}
                if len(cols) > 2:
                    g3 = g2sub.sort_values(ascending=False)
                    kids = [{"name": (k3 or "无"), "value": int(v)} for k3, v in g3.head(6).items()]
                    rest3 = int(g3.iloc[6:].sum()) if len(g3) > 6 else 0
                    if rest3:
                        kids.append({"name": "其他图号", "value": rest3})
                    n2["children"] = kids
                children.append(n2)
            rest2 = int(g2.iloc[8:].sum()) if len(g2) > 8 else 0
            if rest2:
                children.append({"name": "其他型号", "value": rest2})
            node["children"] = children
        tree.append(node)
    rest1 = int(g1.iloc[top:].sum()) if len(g1) > top else 0
    if rest1:
        tree.append({"name": "其他系列", "value": rest1})
    return tree


def _apd_pivot(sub: pd.DataFrame, dates: list, prefix: str):
    """下达日期透视：日期→系列→型号→图号 数量（前端折叠展开）。
    2026-10-18 性能：单次 groupby(日期×系列×TYP×图号) + MultiIndex 逐层取数，
    替代原来"每个节点都全表布尔过滤"的写法（string 列比较是最大热点，原 1.6s → <0.2s）。
    注意：分组列（安全钳/框架 系列/TYP/图号）由 _apd_match 产出，恒为字符串（未命中="无"），无 NaN，
    故单次 groupby(dropna=True) 与原逐层 groupby(dropna 默认) 结果完全一致。"""
    scol = "安全钳系列" if prefix == "safety" else "框架系列"
    tcol = "安全钳TYP" if prefix == "safety" else "框架TYP"
    gcol = "安全钳图号" if prefix == "safety" else "框架图号"
    out = []
    if sub.empty:
        return [{"date": d, "total": 0, "children": []} for d in dates]
    g = sub.groupby(["下达日期", scol, tcol, gcol], dropna=True)["合同箱号"].count()
    date_totals = sub["下达日期"].value_counts()
    for d in dates:
        try:
            gd = g.loc[d]                       # index=(scol, tcol, gcol) 的计数 Series
        except KeyError:
            out.append({"date": d, "total": int(date_totals.get(d, 0)), "children": []})
            continue
        series_node = []
        c1s = gd.groupby(level=0).sum().sort_values(ascending=False)
        for k1, c1 in c1s.items():
            g1 = gd.loc[k1]                     # index=(tcol, gcol)
            typ_node = []
            c2s = g1.groupby(level=0).sum().sort_values(ascending=False).head(10)
            for k2, c2 in c2s.items():
                g2 = g1.loc[k2]                 # index=gcol
                figs = [{"name": (k3 or "无图号"), "value": int(v)}
                        for k3, v in g2.sort_values(ascending=False).head(12).items()]
                typ_node.append({"name": k2 or "无", "value": int(c2), "children": figs})
            # TYP 值与系列同名时跳过该层（图号直挂系列），避免"E系列下还是E系列"的重复层级（用户 2026-10-16）
            if len(typ_node) == 1 and (typ_node[0]["name"] == (k1 or "无")):
                series_node.append({"name": k1 or "无", "value": int(c1), "children": typ_node[0]["children"]})
            else:
                series_node.append({"name": k1 or "无", "value": int(c1), "children": typ_node})
        out.append({"date": d, "total": int(date_totals.get(d, 0)), "children": series_node})
    return out


def _apd_attr_pivot(sub: pd.DataFrame, date_col: str, top_dates: list):
    """框架属性×图号 透视（数量+行内占比），按下达日期或完工日期。
    2026-10-18 性能：单次 groupby(日期×属性×图号) + MultiIndex 取数（原逐节点全表过滤 1.3s → <0.1s）。
    框架属性恒为字符串（_apd_load_df 已 fillna+strip+upper，可能为""）；框架图号恒为字符串。"""
    out = []
    if sub.empty:
        return [{"date": d, "total": 0, "children": []} for d in top_dates]
    g = sub.groupby([date_col, "框架属性", "框架图号"], dropna=True)["合同箱号"].count()
    date_totals = sub[date_col].value_counts()
    for d in top_dates:
        if d == "全区间":
            gd = g.groupby(level=["框架属性", "框架图号"]).sum()
            total = int(len(sub))
        else:
            total = int(date_totals.get(d, 0))
            if total == 0:
                out.append({"date": d, "total": 0, "children": []})
                continue
            try:
                gd = g.loc[d]
            except KeyError:
                out.append({"date": d, "total": total, "children": []})
                continue
        kids = []
        for k1, c1 in gd.groupby(level=0).sum().sort_values(ascending=False).items():
            figs = [{"name": (k2 or "无图号"), "value": int(v)}
                    for k2, v in gd.loc[k1].sort_values(ascending=False).head(8).items()]
            kids.append({"name": k1 or "无", "value": int(c1),
                         "pct": round(c1 / total * 100, 1), "children": figs})
        out.append({"date": d, "total": total, "children": kids})
    return out


# ============ 3号箱APD及安全钳看板（深色指挥塔大屏版 · 同源 sap_table.03_apd） ============
# 数据管线（_apd_load_df/_apd_load_mapping/_apd_match/_apd_tree/_apd_pivot/_apd_attr_pivot）在本节上方定义，
# 独立接口 /api/apd22/*，支持多级下钻筛选：时间范围/梯种/框架属性/系列/型号/完工日期/角标/BG。

_APD22_ENRICH = {"df": None, "ts": 0.0, "ver": 0}
_APD22_ENRICH_LOCK = threading.Lock()   # 单航班：options/data/detail 并发冷启动时不重复干活


def _apd22_enriched() -> pd.DataFrame:
    """全量匹配缓存：在 _apd_load_df 之上拼接 系列/TYP/图号 两套匹配列（120s 缓存）。
    2026-10-18：加单航班锁（原无锁，页面进入时 options+data+detail 三并发会各跑一遍
    全量匹配，GIL 下串 executed 还互相拖慢）；ver 每次重建自增，供响应缓存判失效。"""
    import time as _time
    now = _time.time()
    with _APD22_ENRICH_LOCK:
        if _APD22_ENRICH["df"] is not None and now - _APD22_ENRICH["ts"] < _APD_TTL:
            return _APD22_ENRICH["df"]
        df = _apd_load_df()
        mapping = _apd_load_mapping()
        frame_mapping = _apd_load_frame_mapping()
        ss, st_, sg = _apd_match(df["安全钳组件物料"], mapping, prefix_mode=False)
        fs, ft_, fg = _apd_match(df["框架组件物料"], frame_mapping, prefix_mode=True)
        df = df.assign(安全钳系列=ss, 安全钳TYP=st_, 安全钳图号=sg,
                       框架系列=fs, 框架TYP=ft_, 框架图号=fg)
        # 3号箱APD及安全钳看板 数据清洗（仅本看板口径）：角标 32/321 的数据行整体删除后重算；03→03#箱、031→03#箱加箱
        df = df[~df["角标"].isin(["32", "321"])].copy()
        # 角标重命名：唯一值记忆化后走 dict map（C 速），替代 16 万行 Python lambda（2026-10-18 性能）
        _cmap = {k: {"03": "03#箱", "031": "03#箱加箱"}.get(k, k) for k in df["角标"].unique()}
        df["角标"] = df["角标"].map(_cmap)
        _APD22_ENRICH.update(df=df, ts=now, ver=_APD22_ENRICH["ver"] + 1)
        return df


def _apd22_tree4(sub: pd.DataFrame, cols: list, tops: list, empty_names: dict = None):
    """四级树（如 完工日期→系列→TYP→图号），逐级截断 + 「其他」兜底。
    empty_names: {列名: 空值显示名}，如 {"箱子完工日期": "未完工"}。
    2026-10-18 性能：原实现逐节点 df[df[col]==k] 反复全表扫描（dropna=False 下还有 isna 特例），
    是首屏 5s+ 的主因之一；改为单次 groupby(cols, dropna=False) 后按 MultiIndex 逐层
    xs/聚合/isin 取数，逐项结果与原实现一致（含 NaN 键进「其他」的行为）。"""
    if sub.empty:
        return []
    empty_names = empty_names or {}
    nlev = len(cols)
    g = sub.groupby(cols, dropna=False)["合同箱号"].count()

    def _name(col, k):
        s = "" if k is None or (isinstance(k, float) and pd.isna(k)) else str(k)
        if s in ("", "nan"):
            return empty_names.get(col, "无")
        return s

    def build(glev, lev):
        """glev：index 为 cols[lev:] 剩余层的计数 Series；返回该层节点列表。"""
        col = cols[lev]
        if len(glev.index.names) > 1:
            c = glev.groupby(level=0).sum()
        else:
            c = glev
        c = c.sort_values(ascending=False)
        if not c.size:
            return []
        top = tops[lev]
        head_keys = list(c.head(top).index)
        nodes = []
        for k in head_keys:
            node = {"name": _name(col, k), "value": int(c[k])}
            if lev + 1 < nlev:
                try:
                    sub_g = glev.xs(k, level=0) if len(glev.index.names) > 1 else glev
                    node["children"] = build(sub_g, lev + 1) if len(sub_g) else []
                except Exception:
                    node["children"] = []
            nodes.append(node)
        if len(c) > top:
            # 「其他」= 去掉 top 键后，把剩余层聚合出来（注意：必须 groupby 聚合，
            # 不能只 droplevel——droplevel 不合并重复元组，深层 xs 会拿到未聚合的重复行）
            rest_raw = glev[~glev.index.get_level_values(0).isin(head_keys)]
            nrem = len(rest_raw.index.names)
            rest = rest_raw.groupby(level=list(range(1, nrem))).sum() if nrem > 1 else rest_raw
            node = {"name": "其他" + col, "value": int(c.iloc[top:].sum())}
            # 注意：原实现的「其他」节点即使在末层也带 children:[]（常规末层节点不带），此处保持逐字节一致
            node["children"] = build(rest, lev + 1) if (lev + 1 < nlev and len(rest)) else []
            nodes.append(node)
        return nodes

    return build(g, 0)


class Apd22Filter(BaseModel):
    start: str = ""          # 下达时间起
    end: str = ""            # 下达时间止
    cstart: str = ""         # 角标图自定义下达起
    cend: str = ""           # 角标图自定义下达止
    tstart: str = ""         # 梯种玫瑰自定义下达起
    tend: str = ""           # 梯种玫瑰自定义下达止
    fstart: str = ""         # 完工时间起
    fend: str = ""           # 完工时间止
    tizi: List[str] = []     # 梯种
    s_series: List[str] = []   # 安全钳组件 · 系列
    s_typ: List[str] = []      # 安全钳组件 · TYP
    s_fig: List[str] = []      # 安全钳组件 · 图号
    f_series: List[str] = []   # 框架组件 · 分类（框架配置分类的「分类」列）
    f_fig: List[str] = []      # 框架组件 · 图号
    attr: List[str] = []     # 框架属性（X/F，保留供下钻联动）
    finish: List[str] = []   # 箱子完工日期下钻（chips，与 fstart/fend 并存）
    corner: List[str] = []   # 合同箱号角标下钻
    bg: List[str] = []       # BG 参数下钻
    # 2026-10-26 看板日报用：明细只导出「最近半年」+ 轻量写法（4 万行逐格描边要几十秒）
    detail_start: str = ""   # 明细起始下达日期（yyyy-mm-dd / yyyy.mm.dd / yyyymmdd 均可；空=不限）
    lite: int = 0            # 1 = 明细不逐格描边/斑马纹（只留表头样式+冻结），导出快一个数量级


_APD22_OPTS_CACHE = {"ver": -1, "body": None}   # options 响应缓存：按数据版本失效（SWR 窗口也不堵 enrich 锁）


@app.get("/api/apd22/options")
def apd22_options():
    """筛选器选项：下达日期、完工日期、梯种、安全钳（系列/TYP/图号）、框架（分类/图号）。
    2026-10-18：按 enriched 版本缓存（版本过期时先返回陈旧版，不阻塞在 enrich 锁上）。"""
    _v = _APD22_ENRICH.get("ver", -1)
    if _APD22_OPTS_CACHE["body"] is not None and _APD22_OPTS_CACHE["ver"] == _v:
        return Response(content=_APD22_OPTS_CACHE["body"], media_type="application/json")
    df = _apd22_enriched()
    dates = sorted(df["下达日期"].unique(), reverse=True)
    tizi = sorted([t for t in df["梯种"].unique() if t and t != "nan"])
    attrs = sorted([a for a in df["框架属性"].unique() if a])
    s_typs = sorted([t for t in df["安全钳TYP"].unique() if t and t != "无"])
    s_series = sorted([t for t in df["安全钳系列"].unique() if t and t != "无"])
    s_figs = sorted([t for t in df["安全钳图号"].unique() if t and t != "无"])
    f_series = sorted([t for t in df["框架系列"].unique() if t and t != "无"])
    f_figs = sorted([t for t in df["框架图号"].unique() if t and t != "无"])
    finishes = sorted([d for d in df["箱子完工日期"].fillna("").unique() if d], reverse=True)
    payload = {"dates": dates, "tizi": tizi, "attrs": attrs, "finishes": finishes,
               "s_series": s_series, "s_typs": s_typs, "s_figs": s_figs,
               "f_series": f_series, "f_figs": f_figs}
    try:
        _APD22_OPTS_CACHE.update(ver=_APD22_ENRICH.get("ver", -1),
                                 body=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
    except Exception:
        pass
    if _APD22_OPTS_CACHE["body"] is not None and _APD22_OPTS_CACHE["ver"] == _v:
        return Response(content=_APD22_OPTS_CACHE["body"], media_type="application/json")
    return payload





def _apd22_filtered(f: "Apd22Filter"):
    """按3号箱APD及安全钳看板筛选条件过滤 enriched 明细（导出与主数据共用同一口径）。"""
    df = _apd22_enriched()
    if f.tizi:
        df = df[df["梯种"].isin(f.tizi)]
    if f.attr:
        df = df[df["框架属性"].isin(f.attr)]
    if f.finish:
        df = df[df["箱子完工日期"].fillna("").isin(f.finish)]
    if f.corner:
        df = df[df["角标"].isin(f.corner)]
    if f.s_series:
        df = df[df["安全钳系列"].isin(f.s_series)]
    if f.s_typ:
        df = df[df["安全钳TYP"].isin(f.s_typ)]
    if f.s_fig:
        df = df[df["安全钳图号"].isin(f.s_fig)]
    if f.f_series:
        df = df[df["框架系列"].isin(f.f_series)]
    if f.f_fig:
        df = df[df["框架图号"].isin(f.f_fig)]
    if f.bg:
        df = df[df["BG参数"].isin(f.bg)]
    if f.start:
        df = df[df["下达日期"] >= f.start]
    if f.end:
        df = df[df["下达日期"] <= f.end]
    if f.fstart:
        df = df[df["箱子完工日期"].fillna("") >= f.fstart]
    if f.fend:
        df = df[df["箱子完工日期"].fillna("") != ""]
        df = df[df["箱子完工日期"] <= f.fend]
    return df

_APD22_DATA_CACHE = {"ver": -1, "body": None, "building": False}   # 默认筛选响应缓存：按数据版本失效
_APD22_SWR_LOCK = threading.Lock()


def _apd22_swr_rebuild():
    """stale-while-revalidate：TTL 过期后后台重建 enriched + 默认响应缓存（单航班，防并发重入）。"""
    with _APD22_SWR_LOCK:
        if _APD22_DATA_CACHE["building"]:
            return
        _APD22_DATA_CACHE["building"] = True
    try:
        _APD22_ENRICH.update(ts=0.0)          # 让 enriched 缓存过期，强制重建
        _apd22_enriched()
        apd22_data(Apd22Filter())             # 重算并刷新默认响应缓存
        print("[apd22 swr] 后台重建完成，默认响应缓存已刷新")
    except Exception as e:
        print("[apd22 swr] 后台重建失败：", e)
    finally:
        _APD22_DATA_CACHE["building"] = False


@app.post("/api/apd22/data")
def apd22_data(f: Apd22Filter):
    """3号箱APD及安全钳看板主数据：KPI 数字带 + 15日柱线 + 角标/梯种 + 安全钳透视与完工日期树 + 框架透视/旭日/BG帕累托 + 属性透视。
    2026-10-18：默认筛选（无任何条件，即页面进入）响应整段缓存为预序列化 JSON 字符串；
    缓存过期（enriched 重建过）时先返回陈旧响应，同时后台单航班重建——用户永远不等计算。"""
    # 默认筛选命中缓存：直接返回预序列化 JSON（省去 pandas 计算 + jsonable_encoder + json.dumps）
    _default = not any([f.start, f.end, f.cstart, f.cend, f.tstart, f.tend, f.fstart, f.fend,
                        f.tizi, f.s_series, f.s_typ, f.s_fig, f.f_series, f.f_fig, f.attr,
                        f.finish, f.corner, f.bg])
    if _default and _APD22_DATA_CACHE["body"]:
        if _APD22_DATA_CACHE["ver"] != _APD22_ENRICH.get("ver", -1):
            # 数据已重建过 → 响应过期：先返回陈旧版，后台刷新（SWR）
            if not _APD22_DATA_CACHE["building"]:
                threading.Thread(target=_apd22_swr_rebuild, daemon=True).start()
        return Response(content=_APD22_DATA_CACHE["body"], media_type="application/json")
    df = _apd22_enriched()
    if f.tizi:
        df = df[df["梯种"].isin(f.tizi)]
    if f.attr:
        df = df[df["框架属性"].isin(f.attr)]
    if f.finish:
        df = df[df["箱子完工日期"].fillna("").isin(f.finish)]
    if f.corner:
        df = df[df["角标"].isin(f.corner)]
    if f.s_series:
        df = df[df["安全钳系列"].isin(f.s_series)]
    if f.s_typ:
        df = df[df["安全钳TYP"].isin(f.s_typ)]
    if f.s_fig:
        df = df[df["安全钳图号"].isin(f.s_fig)]
    if f.f_series:
        df = df[df["框架系列"].isin(f.f_series)]
    if f.f_fig:
        df = df[df["框架图号"].isin(f.f_fig)]
    if f.bg:
        df = df[df["BG参数"].isin(f.bg)]
    if f.start:
        df = df[df["下达日期"] >= f.start]
    if f.end:
        df = df[df["下达日期"] <= f.end]
    if f.fstart:
        df = df[df["箱子完工日期"].fillna("") >= f.fstart]
    if f.fend:
        df = df[df["箱子完工日期"].fillna("") != ""]
        df = df[df["箱子完工日期"] <= f.fend]
    if df.empty:
        return {"empty": True}

    def _clean_nan(o):
        import math as _math
        try:
            import numpy as _np
        except Exception:
            _np = None
        if o is None or isinstance(o, (str, int, bool)):
            return o
        if _np is not None and isinstance(o, _np.floating):
            f = float(o)
            return None if (_math.isnan(f) or _math.isinf(f)) else f
        if _np is not None and isinstance(o, _np.integer):
            return int(o)
        if isinstance(o, float):
            return None if (_math.isnan(o) or _math.isinf(o)) else o
        if isinstance(o, dict):
            return {kk: _clean_nan(vv) for kk, vv in o.items()}
        if isinstance(o, list):
            return [_clean_nan(vv) for vv in o]
        return o

    dates = sorted(df["下达日期"].unique(), reverse=True)
    latest, prev = (dates[0], dates[1]) if len(dates) > 1 else (dates[0], "")
    n_latest = int((df["下达日期"] == latest).sum())
    n_prev = int((df["下达日期"] == prev).sum()) if prev else 0
    matched = int((df["安全钳图号"] != "无").sum())
    fmatched = int((df["框架图号"] != "无").sum())
    d15 = dates[:30]
    df15 = df[df["下达日期"].isin(d15)]

    daily = []
    fin_by_date = df15.groupby(["下达日期", "箱子完工日期"])["合同箱号"].count()
    for d in d15:
        try:
            ser = fin_by_date[d].sort_values(ascending=False)
        except KeyError:
            ser = pd.Series(dtype=int)
        head = list(ser.head(6).items())
        rest = int(ser.iloc[6:].sum()) if len(ser) > 6 else 0
        daily.append({"date": d, "total": int(ser.sum()),
                      "stacks": [{"date": (k if k and not pd.isna(k) else "未完工"), "value": int(v)} for k, v in head]
                                + ([{"date": "其他完工日", "value": rest}] if rest else [])})

    def _corner_range(_days):
        if _days == "all":
            _sub = df
        else:
            _n = 1 if _days == "day" else int(_days)
            _sub = df[df["下达日期"].isin(dates[:_n])]
        _g = _sub.groupby("角标")["合同箱号"].count().sort_values(ascending=False)
        return [{"name": (k or "无"), "value": int(v)} for k, v in _g.head(8).items()]
    cornerByRange = {rng: _corner_range(rng) for rng in ["day", "7", "15", "30", "all"]}
    if f.cstart or f.cend:
        _sub = df
        if f.cstart:
            _sub = _sub[_sub["下达日期"] >= f.cstart.replace("-", ".")]
        if f.cend:
            _sub = _sub[_sub["下达日期"] <= f.cend.replace("-", ".")]
        _g = _sub.groupby("角标")["合同箱号"].count().sort_values(ascending=False)
        cornerByRange["custom"] = [{"name": (k or "无"), "value": int(v)} for k, v in _g.head(8).items()]

    # 梯种玫瑰时间范围（与角标图 cornerByRange 同套路：day/7/15/30/all + custom）
    def _tizi_range(_days):
        if _days == "all":
            _sub = df
        else:
            _n = 1 if _days == "day" else int(_days)
            _sub = df[df["下达日期"].isin(dates[:_n])]
        _g = _sub.groupby("梯种")["合同箱号"].count().sort_values(ascending=False)
        return [{"name": (k or "无"), "value": int(v)} for k, v in _g.head(12).items()]
    tiziByRange = {rng: _tizi_range(rng) for rng in ["day", "7", "15", "30", "all"]}
    if f.tstart or f.tend:
        _sub = df
        if f.tstart:
            _sub = _sub[_sub["下达日期"] >= f.tstart.replace("-", ".")]
        if f.tend:
            _sub = _sub[_sub["下达日期"] <= f.tend.replace("-", ".")]
        _g = _sub.groupby("梯种")["合同箱号"].count().sort_values(ascending=False)
        tiziByRange["custom"] = [{"name": (k or "无"), "value": int(v)} for k, v in _g.head(12).items()]
    corner = df.groupby("角标")["合同箱号"].count().sort_values(ascending=False)
    tizi_g = df.groupby("梯种")["合同箱号"].count().sort_values(ascending=False)

    bg = df[df["BG参数"] != ""].groupby("BG参数")["合同箱号"].count().sort_values(ascending=False)
    bg_total = int(bg.sum())
    pareto, cum = [], 0
    for k, v in bg.head(15).items():
        cum += int(v)
        sub = df[df["BG参数"] == k]
        brk = (sub.groupby(["框架系列", "框架TYP", "框架图号"])["合同箱号"].count()
               .sort_values(ascending=False).head(10))
        pareto.append({"bg": k, "value": int(v), "cumPct": round(cum / bg_total * 100, 1),
                       "breakdown": [{"series": a or "无", "typ": b or "无", "fig": c or "无图号", "value": int(n)}
                                     for (a, b, c), n in brk.items()],
                       "extra": max(0, int(len(sub.groupby(["框架系列", "框架TYP", "框架图号"]))) - 10)})
    rest_bg = int(bg.iloc[15:].sum()) if len(bg) > 15 else 0
    if rest_bg:
        pareto.append({"bg": "其他", "value": rest_bg, "cumPct": 100.0, "breakdown": [], "extra": 0})

    fin_valid = df[df["箱子完工日期"].fillna("") != ""]
    # 悬浮明细：最近 8 个自然日（连续、不跳日）的 03#箱数量（无完工的日期补 0）
    finishDays = []
    if len(fin_valid):
        def _iso(s):
            s = str(s)[:10].strip()
            if "." in s:
                pp = s.split(".")
                s = "%s-%02d-%02d" % (pp[0], int(pp[1]), int(pp[2])) if len(pp) == 3 else s
            return s
        _tmp = fin_valid.assign(_iso=fin_valid["箱子完工日期"].map(
            {d: _iso(d) for d in fin_valid["箱子完工日期"].unique()}))   # 唯一值记忆化，替代逐行 lambda（2026-10-18 性能）
        _f03 = _tmp[_tmp["角标"] == "03#箱"]
        _by_day = _f03.groupby("_iso").size().sort_index(ascending=False)
        for _d in _by_day.head(8).index:
            _sub = _f03[_f03["_iso"] == _d].groupby("下达日期").size().sort_index(ascending=False)
            finishDays.append({"date": _d, "count": int(_by_day[_d]),
                "srcs": [{"date": _sd, "count": int(sc)} for _sd, sc in _sub.items()]})
    # 小型化每日下达（KPI 预留位3，用户 2026-10-23 要求）：安全钳系列=小型化，
    # 按下达日期统计箱数 + TYP/图号构成（全部日期降序，前端按 当天/近7/15/30/全部/自定义 切片）
    _mini = df[df["安全钳系列"] == "小型化"]
    mini_latest = int((_mini["下达日期"] == latest).sum()) if len(_mini) else 0
    mini_prev = int((_mini["下达日期"] == prev).sum()) if (prev and len(_mini)) else 0
    mini_by_date = []
    if len(_mini):
        _mg = _mini.groupby(["下达日期", "安全钳TYP", "安全钳图号"])["合同箱号"].count()
        _figs_by_dt = {}
        for (_dt, _typ, _fig), _v in _mg.items():
            _figs_by_dt.setdefault(_dt, []).append(
                {"typ": (_typ if _typ else "无"), "fig": (_fig if _fig else "无图号"), "value": int(_v)})
        for _dt, _v in _mini.groupby("下达日期")["合同箱号"].count().items():
            mini_by_date.append({"date": _dt, "total": int(_v),
                                 "figs": sorted(_figs_by_dt.get(_dt, []), key=lambda x: -x["value"])})
        mini_by_date.sort(key=lambda x: x["date"], reverse=True)
    payload = _clean_nan({
        "empty": False,
        "kpi": {"latest": latest, "latestCount": n_latest, "prev": prev, "prevCount": n_prev,
                "total": int(len(df)), "dates": len(dates),
                "avgDay": round(len(df) / max(1, len(dates)), 1),
                "matchPct": round(matched / len(df) * 100, 1),
                "frameMatchPct": round(fmatched / len(df) * 100, 1),
                "finishDates": int(df["箱子完工日期"].fillna("").ne("").sum()),
                "cornerTop": (corner.index[0] if len(corner) else "")},
        "daily15": daily, "finishDays": finishDays,
        "miniDaily": {"latest": latest, "latestCount": mini_latest, "prev": prev, "prevCount": mini_prev,
                      "dayTotal": n_latest, "byDate": mini_by_date},
        "cornerByRange": cornerByRange,
        "tiziByRange": tiziByRange,
        "corner": [{"name": (k or "无"), "value": int(v)} for k, v in corner.head(8).items()],
        "tizi": [{"name": (k or "无"), "value": int(v)} for k, v in tizi_g.head(14).items()],
        "safetyPivot": _apd_pivot(df, d15, "safety"),
        "safetyFinishTree": _apd22_tree4(fin_valid, ["箱子完工日期", "安全钳系列", "安全钳TYP", "安全钳图号"],
                                         [8, 6, 6, 5], {"箱子完工日期": "未完工"}),
        "framePivot": _apd_pivot(df, d15, "frame"),
        "frameSun": _apd_tree(df.dropna(subset=["框架组件物料"]), ["框架系列", "框架TYP", "框架图号"]),
        "bgPareto": pareto, "bgTotal": bg_total,
        "attrByIssue": _apd_attr_pivot(df, "下达日期", d15) + _apd_attr_pivot(df, "下达日期", ["全区间"]),
        "attrByFinish": _apd_attr_pivot(fin_valid, "箱子完工日期",
                                        sorted([d for d in fin_valid["箱子完工日期"].unique() if d], reverse=True)[:15])
                                        + _apd_attr_pivot(fin_valid, "箱子完工日期", ["全区间"]),
    })
    if _default:
        # 默认筛选：缓存预序列化 JSON（含时间戳会随每次请求变？——状态栏时间在前端生成，此处无时间敏感字段）
        try:
            _APD22_DATA_CACHE.update(ver=_APD22_ENRICH.get("ver", -1),
                                     body=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
        except Exception:
            pass
        return Response(content=_APD22_DATA_CACHE["body"], media_type="application/json")
    return payload


@app.post("/api/apd22/export_xlsx")
def apd22_export_xlsx(f: Apd22Filter):
    """3号箱APD及安全钳看板 导出 Excel（openpyxl 深色表头+斑马纹+列宽自适应，多工作表还原看板结构）。"""
    import io
    import urllib.parse
    import datetime as _dt
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from fastapi.responses import Response

    df = _apd22_filtered(f)
    if df.empty:
        return {"error": "无数据（当前筛选条件下无记录）"}

    HDR_FILL = PatternFill("solid", fgColor="0D1E3C")
    HDR_FONT = Font(name="微软雅黑", size=10, bold=True, color="EAF3FF")
    TITLE_FONT = Font(name="微软雅黑", size=13, bold=True, color="0D1E3C")
    BODY_FONT = Font(name="微软雅黑", size=10, color="1A2A44")
    NOTE_FONT = Font(name="微软雅黑", size=9, color="5F7A9E")
    ZEBRA_FILL = PatternFill("solid", fgColor="EDF3FB")
    KPI_FILL = PatternFill("solid", fgColor="E8F1FC")
    THIN = Side(style="thin", color="C9D6E8")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER = Alignment(horizontal="center", vertical="center")
    LEFT = Alignment(horizontal="left", vertical="center")
    RIGHT = Alignment(horizontal="right", vertical="center")

    def _wlen(s):
        return sum(2 if ord(ch) > 255 else 1 for ch in str(s))

    def style_sheet(ws, headers, rows, title=None):
        r0 = 1
        if title:
            ws.cell(1, 1, title).font = TITLE_FONT
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 2))
            r0 = 3
        for c, h in enumerate(headers, 1):
            cell = ws.cell(r0, c, h)
            cell.fill = HDR_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row, 1):
                cell = ws.cell(r0 + 1 + ri, ci, v)
                cell.font = BODY_FONT; cell.border = BORDER
                cell.alignment = LEFT if isinstance(v, str) else RIGHT
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0"
                if ri % 2 == 1:
                    cell.fill = ZEBRA_FILL
        for c, h in enumerate(headers, 1):
            w = _wlen(h) + 4
            for row in rows[:200]:
                if c <= len(row):
                    w = max(w, _wlen(row[c - 1]) + 2)
            ws.column_dimensions[get_column_letter(c)].width = min(w, 42)
        ws.freeze_panes = ws.cell(r0 + 1, 1)

    def walk_tree(ws, pv_item, rows, depth=1):
        total = pv_item["total"]
        rows.append([pv_item["date"] + " 合计", total, "100%"])
        def rec(nodes, d):
            for n in nodes:
                rows.append(["　" * d + str(n["name"]), n["value"], (str(round(n["value"] / total * 100, 1)) + "%") if total else ""])
                if n.get("children"):
                    rec(n["children"], d + 1)
        rec(pv_item["children"], depth)

    wb = Workbook()
    dates = sorted(df["下达日期"].unique(), reverse=True)
    d15 = dates[:15]

    # Sheet1 汇总
    ws = wb.active; ws.title = "汇总"
    latest = d15[0] if d15 else ""
    prev = d15[1] if len(d15) > 1 else ""
    n_latest = int(len(df[df["下达日期"] == latest])) if latest else 0
    n_prev = int(len(df[df["下达日期"] == prev])) if prev else 0
    delta = ((n_latest - n_prev) / n_prev * 100) if n_prev else 0
    ws.cell(1, 1, "3号箱APD及安全钳看板 · 导出汇总").font = Font(name="微软雅黑", size=15, bold=True, color="0D1E3C")
    ws.merge_cells("A1:D1")
    kpis = [
        ("最新下达日箱数", latest, n_latest, ("vs 前一下达日 {:+.1f}%".format(delta) if prev else "")),
        ("下达日总数（区间内）", "", int(df["下达日期"].nunique()), ""),
        ("总箱数（区间内）", "", int(len(df)), ""),
        ("已完工箱数", "", int(df["箱子完工日期"].fillna("").ne("").sum()), ""),
    ]
    r = 3
    for name, d, v, note in kpis:
        a = ws.cell(r, 1, name); a.font = Font(name="微软雅黑", size=10, bold=True, color="1A2A44"); a.fill = KPI_FILL; a.border = BORDER; a.alignment = LEFT
        b = ws.cell(r, 2, d); b.font = BODY_FONT; b.border = BORDER; b.alignment = CENTER
        c = ws.cell(r, 3, v); c.font = Font(name="微软雅黑", size=12, bold=True, color="0D1E3C"); c.border = BORDER; c.alignment = RIGHT; c.number_format = "#,##0"
        dd = ws.cell(r, 4, note); dd.font = NOTE_FONT; dd.border = BORDER; dd.alignment = LEFT
        r += 1
    no_filter = not any([f.start, f.end, f.fstart, f.fend, f.tizi, f.s_series, f.s_typ, f.s_fig, f.f_series, f.f_fig, f.attr, f.finish, f.corner, f.bg])
    ws.cell(r + 1, 1, "导出时间：" + _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " · 筛选：" + ("无（全量）" if no_filter else "自定义（数据为筛选后结果）")).font = NOTE_FONT
    ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 13; ws.column_dimensions["C"].width = 11; ws.column_dimensions["D"].width = 24

    # Sheet2 每日下达
    ws2 = wb.create_sheet("每日下达")
    g2 = df.groupby("下达日期")["合同箱号"].count().sort_values(ascending=False).head(15)
    style_sheet(ws2, ["下达日期", "下达箱数"], [[k, int(v)] for k, v in g2.items()], title="近 15 个下达日箱数（按数量降序）")

    # Sheet3 角标与梯种
    ws3 = wb.create_sheet("角标与梯种")
    cg = df.groupby("角标")["合同箱号"].count().sort_values(ascending=False).head(10)
    style_sheet(ws3, ["角标", "箱数"], [[("无" if not k else k), int(v)] for k, v in cg.items()], title="角标分布")
    tg = df.groupby("梯种")["合同箱号"].count().sort_values(ascending=False).head(12)
    r0b = 3 + len(cg) + 3
    ws3.cell(r0b - 1, 1, "梯种分布").font = TITLE_FONT
    for c, h in enumerate(["梯种", "箱数"], 1):
        cell = ws3.cell(r0b, c, h); cell.fill = HDR_FILL; cell.font = HDR_FONT; cell.alignment = CENTER; cell.border = BORDER
    for ri, (k, v) in enumerate(tg.items()):
        for ci, val in enumerate([("无" if not k else str(k)), int(v)], 1):
            cell = ws3.cell(r0b + 1 + ri, ci, val); cell.font = BODY_FONT; cell.border = BORDER
            if ci == 2: cell.number_format = "#,##0"

    # Sheet4/5 安全钳/框架透视（最新下达日）
    for sheet_name, prefix, label in (("安全钳透视", "safety", "安全钳"), ("框架透视", "frame", "框架")):
        wsn = wb.create_sheet(sheet_name)
        rows = []
        pv = _apd_pivot(df, d15[:1], prefix)
        if pv:
            walk_tree(ws4 := wsn, pv[0], rows)  # noqa
        style_sheet(wsn, ["层级（系列/TYP/图号）", "数量", "总计占比"], rows, title=label + "透视（最新下达日 " + (d15[0] if d15 else "") + "）")

    # Sheet6 框架属性透视（全区间）
    ws6 = wb.create_sheet("框架属性透视")
    rows6 = []
    ap = _apd_attr_pivot(df, "下达日期", ["全区间"])
    if ap and ap[0]["children"]:
        for a in ap[0]["children"]:
            rows6.append([a["name"], a["value"], str(a.get("pct", 0)) + "%"])
            for fg in a.get("children", [])[:8]:
                rows6.append(["　" + str(fg["name"]), fg["value"], ""])
    style_sheet(ws6, ["框架属性", "数量", "占比"], rows6, title="框架属性 × 图号（全区间）")

    # Sheet7 明细（2026-10-26 改：按下达日期倒叙；detail_start 限定起始下达日；lite=1 走轻量写法）
    ws7 = wb.create_sheet("明细")
    cols = [c for c in ["下达日期", "箱子完工日期", "合同箱号", "角标", "梯种", "安全钳组件物料", "安全钳系列", "安全钳TYP", "安全钳图号", "框架组件物料", "框架系列", "框架TYP", "框架图号", "框架属性", "BG参数"] if c in df.columns]
    d7 = df
    ds = str(getattr(f, "detail_start", "") or "").replace("-", "").replace(".", "").replace("/", "")
    if len(ds) == 8:
        d7 = d7[d7["下达日期"].map(lambda s: str(s).replace("-", "").replace(".", "").replace("/", "") >= ds)]
    d7 = d7.sort_values("下达日期", ascending=False)
    rows7 = d7[cols].fillna("").values.tolist()
    if getattr(f, "lite", 0):
        # 轻量：逐行 append（比逐格 cell() 快一个数量级），只保留表头样式 + 冻结 + 列宽
        title = ("明细（下达日期最近 " + (ds[:4] + "." + ds[4:6] + "." + ds[6:8]) + " 起 · 按下达日期倒叙 · "
                 + str(len(rows7)) + " 条）") if len(ds) == 8 else ("明细（全部 · 按下达日期倒叙 · " + str(len(rows7)) + " 条）")
        ws7.cell(1, 1, title).font = TITLE_FONT
        ws7.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(cols), 2))
        for c, h in enumerate(cols, 1):
            cell = ws7.cell(3, c, h)
            cell.fill = HDR_FILL; cell.font = HDR_FONT; cell.alignment = CENTER
        for row in rows7:
            ws7.append(list(row))
        for c, h in enumerate(cols, 1):
            w = _wlen(h) + 4
            for row in rows7[:300]:
                if c <= len(row):
                    w = max(w, _wlen(row[c - 1]) + 2)
            ws7.column_dimensions[get_column_letter(c)].width = min(w, 42)
        ws7.freeze_panes = ws7.cell(4, 1)
    else:
        style_sheet(ws7, cols, rows7, title="明细（按下达日期倒叙 " + str(len(rows7)) + " 条）")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = urllib.parse.quote("3号箱安全钳3号箱APD及安全钳看板.xlsx")
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''" + fname},
    )

_APD22_DETAIL_CACHE = {}   # {date: (ver, payload)} 明细响应缓存：按数据版本失效（SWR 窗口不堵 enrich 锁）


@app.get("/api/apd22/detail")
def apd22_detail(date: str = "", limit: int = 300):
    """最近一个下达日期（或指定日期）明细行。2026-10-18：按 (日期, 数据版本) 缓存。"""
    _v = _APD22_ENRICH.get("ver", -1)
    _key = date or "__latest__"
    _ent = _APD22_DETAIL_CACHE.get(_key)
    if _ent and _ent[0] == _v and limit == 300:
        payload = dict(_ent[1])
        if not date:
            payload["date"] = _ent[1]["date"]
        return payload
    df = _apd22_enriched()
    if not date:
        ds = sorted(df["下达日期"].unique(), reverse=True)
        date = ds[0] if ds else ""
    sub = df[df["下达日期"] == date].copy()
    sub = sub.sort_values(["箱子完工日期", "合同箱号"], na_position="last")
    sub = sub.head(min(max(limit, 1), 9999))
    cols = ["下达日期", "箱子完工日期", "合同箱号", "角标", "梯种",
            "安全钳组件物料", "安全钳系列", "安全钳TYP", "安全钳图号",
            "框架组件物料", "框架系列", "框架TYP", "框架图号", "框架属性", "BG参数"]
    rows = [[("" if pd.isna(v) else str(v)) for v in row] for row in sub[cols].itertuples(index=False, name=None)]
    payload = {"date": date, "total": int(len(rows)), "headers": cols, "rows": rows}
    if limit == 300:
        if len(_APD22_DETAIL_CACHE) > 40:
            _APD22_DETAIL_CACHE.clear()
        _APD22_DETAIL_CACHE[_key] = (_APD22_ENRICH.get("ver", -1), payload)
    return payload


# ---- 超大安全钳关注（KPI 预留位5 取消后新增，2026-10-23 用户需求） ----
# 用户关注清单（原文含 * 通配且重复，规则=「安全钳图号包含罗列编码即关注」→ 归一物料 contains 匹配）；
# 去重后 16 个编码。匹配口径用归一物料（而非映射图号）：覆盖映射未命中（图号="无"）的行。
_APD22_MEGA_PATTERNS = ["G125013A000", "P125010A001", "P125002A001", "P125011A00",
                        "P125007A001", "P125013A001", "P125068A000", "P125066A000",
                        "P125063A000", "P125062A000", "P125075A000", "P125055A000",
                        "P125064A000", "P125065A000", "P125067A001", "P125073A000"]
_APD22_MEGA_RANGES = {"default": "默认（完工 今天-7天 ~ 今天+180天）",
                      "1y": "最近1年", "3y": "最近3年", "5y": "最近5年", "all": "全部",
                      "custom": "自定义完工日期"}
_APD22_MEGA_MATCH_CACHE = {"df": None, "ver": -1}   # 命中清单的明细（全历史），按 enriched 版本缓存
_APD22_MEGA_CACHE = {}                              # {range: (ver, payload)} 响应缓存，按数据版本失效


def _apd22_mega_matched() -> pd.DataFrame:
    """命中关注清单的明细行（全历史，不受完工窗口限制）。2026-10-23。
    匹配：归一物料 contains 任一清单编码；展示图号=映射安全钳图号（非"无"），否则取命中的清单编码。"""
    _v = _APD22_ENRICH.get("ver", -1)
    ent = _APD22_MEGA_MATCH_CACHE
    if ent["df"] is not None and ent["ver"] == _v:
        return ent["df"]
    df = _apd22_enriched()
    pat = "|".join(_APD22_MEGA_PATTERNS)
    mat = df["安全钳组件物料"].map(_apd_norm_code)
    hit = mat.str.contains(pat, regex=True, na=False)
    sub = df[hit].copy()

    def _first(v):
        for p in _APD22_MEGA_PATTERNS:
            if p in v:
                return p
        return ""
    if len(sub):
        sub = sub.assign(_matched=mat[hit].map(_first))
        sub["关注图号"] = [(g if (g and g != "无") else m)
                        for g, m in zip(sub["安全钳图号"], sub["_matched"])]
    _APD22_MEGA_MATCH_CACHE.update(df=sub, ver=_v)
    return sub


def _apd22_mega_window(range_key: str, start: str = "", end: str = ""):
    """完工日期窗口（date 对象；lo=None 表示不限）。default=[今天-7天, 今天+180天]；
    1y/3y/5y=[今天-N年, 今天+180天]；all=不限（含无完工日期的行）；
    custom=用户选定的 start~end（ISO 'YYYY-MM-DD'，单边可空，两端都空则等同不限）。
    返回 (lo, hi, today, label)。"""
    import datetime as _dt
    today = _dt.date.today()
    hi = today + _dt.timedelta(days=180)
    lo = today - _dt.timedelta(days=7)
    label = _APD22_MEGA_RANGES.get(range_key, range_key)
    if range_key == "custom":
        def _p(s):
            s = (s or "").strip().replace(".", "-")
            if not s:
                return None
            try:
                return _dt.date(*[int(x) for x in s.split("-")[:3]])
            except Exception:
                return None
        lo, hi = _p(start), _p(end)
        label = "自定义完工日期"
    elif range_key in ("1y", "3y", "5y"):
        n = int(range_key[0])
        try:
            lo = today.replace(year=today.year - n)
        except ValueError:                    # 2/29 边界
            lo = today - _dt.timedelta(days=365 * n)
    elif range_key == "all":
        lo, hi = None, None
    return lo, hi, today, label


def _apd22_mega_body(range_key: str, start: str = "", end: str = "") -> dict:
    """超大安全钳关注响应体：窗口筛选 → 下达日期/箱子完工日期/关注图号 三维分类汇总计数 + 明细行。"""
    sub = _apd22_mega_matched()
    lo, hi, today, _label = _apd22_mega_window(range_key, start, end)
    if lo is not None or hi is not None:
        _fin = pd.to_datetime(sub["箱子完工日期"].fillna("").astype(str),
                              format="%Y.%m.%d", errors="coerce")
        keep = _fin.notna()
        if lo is not None:
            keep = keep & (_fin >= pd.Timestamp(lo))
        if hi is not None:
            keep = keep & (_fin <= pd.Timestamp(hi))
        sub = sub[keep]
    _tstr = "%d.%02d.%02d" % (today.year, today.month, today.day)
    if lo is None and hi is None:
        _wstr = "完工日期不限"
    elif lo is None:
        _wstr = "完工 ≤ %d.%02d.%02d" % (hi.year, hi.month, hi.day)
    elif hi is None:
        _wstr = "完工 ≥ %d.%02d.%02d" % (lo.year, lo.month, lo.day)
    else:
        _wstr = "完工 %d.%02d.%02d ~ %d.%02d.%02d" % (lo.year, lo.month, lo.day,
                                                    hi.year, hi.month, hi.day)
    payload = {"ok": True, "range": range_key, "rangeText": _label,
               "window": _wstr, "today": _tstr,
               "total": 0, "figCount": 0, "groupCount": 0, "figs": [], "groups": [], "details": []}
    if sub.empty:
        return payload
    # 卡面「分类汇总计数」：按关注图号汇总箱数（降序）；用户 2026-10-23 明确不要汇总行，
    # 故此处只给分类明细、不算合计（合计由面板统计条单独展示）。
    _fg = sub.groupby("关注图号")["合同箱号"].count().sort_values(ascending=False)
    payload["figs"] = [{"fig": str(k), "count": int(v)} for k, v in _fg.items()]
    g = (sub.groupby(["下达日期", "箱子完工日期", "关注图号"], dropna=False)["合同箱号"]
         .count().reset_index(name="count"))
    g = g.sort_values(["箱子完工日期", "count", "下达日期"], ascending=[True, False, True])
    payload["groups"] = [{"order": a, "finish": b, "fig": c, "count": int(n)}
                         for a, b, c, n in g.itertuples(index=False, name=None)]
    det = sub.sort_values(["箱子完工日期", "下达日期", "合同箱号"], na_position="last")
    # 明细字段：用户 2026-10-23 要求「展示这张明细表所有列字段」——与看板主明细表同口径 15 列
    # （关注图号 = 安全钳图号 的非"无"值，未映射到图号的行用命中的清单编码，保证每行都能看出关注原因）。
    _cols = ["下达日期", "箱子完工日期", "合同箱号", "角标", "梯种",
             "安全钳组件物料", "安全钳系列", "安全钳TYP", "关注图号",
             "框架组件物料", "框架系列", "框架TYP", "框架图号", "框架属性", "BG参数"]
    payload["headers"] = _cols
    payload["details"] = [[("" if pd.isna(v) else str(v)) for v in row]
                          for row in det[_cols].itertuples(index=False, name=None)]
    payload["total"] = int(len(det))
    payload["figCount"] = int(sub["关注图号"].nunique())
    payload["groupCount"] = int(len(g))
    return payload


@app.get("/api/apd22/mega-clamp")
def apd22_mega_clamp(range: str = "default", start: str = "", end: str = ""):
    """超大安全钳关注（2026-10-23）：命中清单图号的箱号按完工日期窗口筛选，
    以 下达日期/箱子完工日期/关注图号 三维分类汇总计数，明细行可复制粘贴。
    range=custom 时用 start/end（ISO，单边可空）作完工日期区间。"""
    key = range if range in _APD22_MEGA_RANGES else "default"
    _sk = (start or "").strip() if key == "custom" else ""
    _ek = (end or "").strip() if key == "custom" else ""
    _ck = (key, _sk, _ek)
    _v = _APD22_ENRICH.get("ver", -1)
    ent = _APD22_MEGA_CACHE.get(_ck)
    if ent and ent[0] == _v:
        return ent[1]
    payload = _apd22_mega_body(key, _sk, _ek)
    if len(_APD22_MEGA_CACHE) > 30:
        _APD22_MEGA_CACHE.clear()
    _APD22_MEGA_CACHE[_ck] = (_v, payload)
    return payload


# ---- 3号箱APD及安全钳看板 数据落地：MySQL `03#箱databoard` 目录（全量刷新重建） ----

APD22_TARGET_DB = dict(host="127.0.0.1", port=3306, user="u3854", password="u38544253",
                       database="03#箱databoard", charset="utf8mb4")

_APD22_TABLES_SQL = {
    "安全钳总图号分类": ("`图号` VARCHAR(64), `系列` VARCHAR(64), `TYP型号` VARCHAR(128)",
                    "`图号`, `系列`, `TYP型号`"),
    "安全钳框架配置分类": ("`图号` VARCHAR(64), `分类` VARCHAR(64)",
                      "`图号`, `分类`"),
    "限速器参数": ("`集团` VARCHAR(32), `图号_变量` VARCHAR(64), `型号` VARCHAR(64), `额定速度` VARCHAR(32), "
                 "`超速开关` VARCHAR(32), `上行动作速度` VARCHAR(32), `下行动作速度` VARCHAR(32), "
                 "`钢丝绳直径` VARCHAR(32), `创建日期` VARCHAR(32), `创建者` VARCHAR(32)",
                 "`集团`, `图号_变量`, `型号`, `额定速度`, `超速开关`, `上行动作速度`, `下行动作速度`, "
                 "`钢丝绳直径`, `创建日期`, `创建者`"),
    # 限速器物料号（2026-11-03 新增，与侧边栏该页同步；ZLSTMA 31列，页面列 物料(类型)→物料类型/物料类型→物料类型2 防撞）
    "限速器物料号": ("`集团` VARCHAR(16), `物料` VARCHAR(32), `物料类型` VARCHAR(16), `品目` VARCHAR(16), `部件名称` VARCHAR(64), "
                 "`图号_变量` VARCHAR(64), `品_群` VARCHAR(16), `L番` VARCHAR(32), `创建日期` VARCHAR(32), `创建者` VARCHAR(32), "
                 "`物料传输标记` VARCHAR(16), `BOM传输标志` VARCHAR(16), `生产科使用` VARCHAR(32), `变量` VARCHAR(128), `物料类型2` VARCHAR(16), "
                 "`物料号长参数表` VARCHAR(128), `标记` VARCHAR(16), `工艺更改版本编号` VARCHAR(32), `工艺生效日期` VARCHAR(32), `工序数量` VARCHAR(16), "
                 "`可选工艺数量` VARCHAR(16), `并行工艺数量` VARCHAR(16), `输入日期` VARCHAR(32), `标记2` VARCHAR(16), `印板工艺属性` VARCHAR(32), "
                 "`材料` VARCHAR(64), `Changed_at` VARCHAR(32), `RDC标记` VARCHAR(16), `海菱用物料标记` VARCHAR(16), `海菱物料传输标记` VARCHAR(16), `海菱BOM传输标记` VARCHAR(16)",
                 "`集团`, `物料`, `物料类型`, `品目`, `部件名称`, `图号_变量`, `品_群`, `L番`, `创建日期`, `创建者`, "
                 "`物料传输标记`, `BOM传输标志`, `生产科使用`, `变量`, `物料类型2`, `物料号长参数表`, `标记`, `工艺更改版本编号`, `工艺生效日期`, `工序数量`, "
                 "`可选工艺数量`, `并行工艺数量`, `输入日期`, `标记2`, `印板工艺属性`, `材料`, `Changed_at`, `RDC标记`, `海菱用物料标记`, `海菱物料传输标记`, `海菱BOM传输标记`"),
    "apd_清洗明细": ("`合同箱号` VARCHAR(64), `箱子完工日期` VARCHAR(32), `安全钳组件物料` VARCHAR(64), "
                 "`框架组件物料` VARCHAR(64), `框架属性` VARCHAR(32), `框架参数` VARCHAR(128), `来源文件夹` VARCHAR(128), "
                 "`下达日期` VARCHAR(32), `角标` VARCHAR(32), `梯种` VARCHAR(32), `BG参数` VARCHAR(32), "
                 "`安全钳系列` VARCHAR(64), `安全钳TYP` VARCHAR(128), `安全钳图号` VARCHAR(64), "
                 "`框架系列` VARCHAR(64), `框架TYP` VARCHAR(128), `框架图号` VARCHAR(64)",
                 "`合同箱号`, `箱子完工日期`, `安全钳组件物料`, `框架组件物料`, `框架属性`, `框架参数`, `来源文件夹`, "
                 "`下达日期`, `角标`, `梯种`, `BG参数`, `安全钳系列`, `安全钳TYP`, `安全钳图号`, "
                 "`框架系列`, `框架TYP`, `框架图号`"),
    "下达日汇总": ("`下达日期` VARCHAR(32), `箱数` INT", "`下达日期`, `箱数`"),
    "角标分布": ("`角标` VARCHAR(32), `箱数` INT", "`角标`, `箱数`"),
    "梯种分布": ("`梯种` VARCHAR(32), `箱数` INT", "`梯种`, `箱数`"),
    "完工日期分布": ("`完工日期` VARCHAR(32), `箱数` INT", "`完工日期`, `箱数`"),
    "BG参数分布": ("`BG参数` VARCHAR(32), `箱数` INT", "`BG参数`, `箱数`"),
}


def _apd22_rebuild_tables() -> dict:
    """全量重建 03#箱databoard 下的数据处理表（DROP→CREATE→批量 INSERT），返回各表行数。"""
    import pymysql
    df = _apd22_enriched()          # 清洗+匹配后的全量明细（含 32/321 删除与角标重命名）
    mapping = _apd_load_mapping()   # 侧边栏「安全钳总图号分类」最新数据
    frame_mapping = _apd_load_frame_mapping()   # 侧边栏「安全钳框架配置分类」
    store_load = _safety_store_seed_if_empty   # 侧边栏两份分类数据（含框架配置分类）
    conn = pymysql.connect(**APD22_TARGET_DB)
    counts = {}
    try:
        cur = conn.cursor()
        for tbl, (cols_def, _cols_ins) in _APD22_TABLES_SQL.items():
            cur.execute(f"DROP TABLE IF EXISTS `{tbl}`")
            cur.execute(f"CREATE TABLE `{tbl}` ({cols_def}) DEFAULT CHARACTER SET utf8mb4")
        conn.commit()

        def _chunks(rows, n=5000):
            for i in range(0, len(rows), n):
                yield rows[i:i + n]

        # 1) 安全钳总图号分类（映射表，与侧边栏同步）
        rows = [(g, s, t) for g, s, t in mapping]
        for ch in _chunks(rows):
            cur.executemany("INSERT INTO `安全钳总图号分类` VALUES (%s, %s, %s)", ch)
        conn.commit(); counts["安全钳总图号分类"] = len(rows)

        # 1b) 安全钳框架配置分类（与侧边栏该页同步：图号+分类 两列）
        frame_rows = []
        fr = (store_load().get("sheets") or {}).get("安全钳框架配置分类") or []
        for r in fr:
            d = r.get("data") or {}
            g = str(d.get("图号") or "").strip()
            if g:
                frame_rows.append((g, str(d.get("分类") or "").strip()))
        for ch in _chunks(frame_rows):
            cur.executemany("INSERT INTO `安全钳框架配置分类` VALUES (%s, %s)", ch)
        conn.commit(); counts["安全钳框架配置分类"] = len(frame_rows)

        # 1c) 限速器参数（与侧边栏该页同步：集团/图号变量/型号/额定速度/超速开关/上行动作速度/下行动作速度/钢丝绳直径/创建日期/创建者）
        xs_rows = []
        xr = (store_load().get("sheets") or {}).get("限速器参数") or []
        for r in xr:
            d = r.get("data") or {}
            xs_rows.append((
                str(d.get("集团") or "").strip(),
                str(d.get("图号/变量") or "").strip(),
                str(d.get("型号") or "").strip(),
                str(d.get("额定速度") or "").strip(),
                str(d.get("超速开关") or "").strip(),
                str(d.get("上行动作速度") or "").strip(),
                str(d.get("下行动作速度") or "").strip(),
                str(d.get("钢丝绳直径") or "").strip(),
                str(d.get("创建日期") or "").strip(),
                str(d.get("创建者") or "").strip(),
            ))
        for ch in _chunks(xs_rows):
            cur.executemany(
                "INSERT INTO `限速器参数` VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", ch)
        conn.commit(); counts["限速器参数"] = len(xs_rows)

        # 1d) 限速器物料号（2026-11-03 新增，与侧边栏该页同步：ZLSTMA 31 列；
        #     页面列名 物料(类型)→物料类型、物料类型→物料类型2 防重名，L 番→L番、BOM 传输标志→BOM传输标志、Changed at→Changed_at）
        lm_rows = []
        lm = (store_load().get("sheets") or {}).get("限速器物料号") or []
        for r in lm:
            d = r.get("data") or {}
            lm_rows.append((
                str(d.get("集团") or "").strip(),
                str(d.get("物料") or "").strip(),
                str(d.get("物料(类型)") or "").strip(),
                str(d.get("品目") or "").strip(),
                str(d.get("部件名称") or "").strip(),
                str(d.get("图号/变量") or "").strip(),
                str(d.get("品/群") or "").strip(),
                str(d.get("L 番") or "").strip(),
                str(d.get("创建日期") or "").strip(),
                str(d.get("创建者") or "").strip(),
                str(d.get("物料传输标记") or "").strip(),
                str(d.get("BOM 传输标志") or "").strip(),
                str(d.get("生产科使用") or "").strip(),
                str(d.get("变量") or "").strip(),
                str(d.get("物料类型") or "").strip(),
                str(d.get("物料号长参数表") or "").strip(),
                str(d.get("标记") or "").strip(),
                str(d.get("工艺更改版本编号") or "").strip(),
                str(d.get("工艺生效日期") or "").strip(),
                str(d.get("工序数量") or "").strip(),
                str(d.get("可选工艺数量") or "").strip(),
                str(d.get("并行工艺数量") or "").strip(),
                str(d.get("输入日期") or "").strip(),
                str(d.get("标记2") or "").strip(),
                str(d.get("印板工艺属性") or "").strip(),
                str(d.get("材料") or "").strip(),
                str(d.get("Changed at") or "").strip(),
                str(d.get("RDC标记") or "").strip(),
                str(d.get("海菱用物料标记") or "").strip(),
                str(d.get("海菱物料传输标记") or "").strip(),
                str(d.get("海菱BOM传输标记") or "").strip(),
            ))
        _lm_ph = ",".join(["%s"] * 31)
        for ch in _chunks(lm_rows):
            cur.executemany(
                "INSERT INTO `限速器物料号` VALUES (" + _lm_ph + ")", ch)
        conn.commit(); counts["限速器物料号"] = len(lm_rows)

        # 2) 清洗明细全量
        det_cols = ["合同箱号", "箱子完工日期", "安全钳组件物料", "框架组件物料", "框架属性", "框架参数",
                    "来源文件夹", "下达日期", "角标", "梯种", "BG参数",
                    "安全钳系列", "安全钳TYP", "安全钳图号", "框架系列", "框架TYP", "框架图号"]
        sub = df[det_cols].fillna("")
        rows = [tuple(str(v) for v in rec) for rec in sub.itertuples(index=False, name=None)]
        for ch in _chunks(rows):
            cur.executemany("INSERT INTO `apd_清洗明细` VALUES (" + ",".join(["%s"] * len(det_cols)) + ")", ch)
        conn.commit(); counts["apd_清洗明细"] = len(rows)

        # 3) 聚合表
        aggs = [
            ("下达日汇总", "`下达日期`", "下达日期"),
            ("角标分布", "`角标`", "角标"),
            ("梯种分布", "`梯种`", "梯种"),
            ("完工日期分布", "`完工日期`", "箱子完工日期"),
            ("BG参数分布", "`BG参数`", "BG参数"),
        ]
        for tbl, cols_sql, col in aggs:
            g = df[df[col].fillna("") != ""].groupby(col)["合同箱号"].count().sort_values(ascending=False)
            rows = [(str(k), int(v)) for k, v in g.items()]
            for ch in _chunks(rows):
                cur.executemany(f"INSERT INTO `{tbl}` VALUES (%s, %s)", ch)
            conn.commit(); counts[tbl] = len(rows)
    finally:
        conn.close()
    return counts


@app.post("/api/apd22/refresh")
def apd22_refresh():
    """全量刷新：清缓存 → 重读 sap_table.03_apd 与侧边栏映射 → 重建 03#箱databoard 全部表。"""
    with _APD_LOCK:
        _APD_CACHE.update(df=None, ts=0.0)
    _APD22_ENRICH.update(df=None, ts=0.0)
    try:
        counts = _apd22_rebuild_tables()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"全量刷新失败：{e}")
    return {"ok": True, "tables": counts, "ts": datetime.now().strftime("%H:%M:%S")}


@app.get("/api/apd22/dbstats")
def apd22_dbstats():
    """03#箱databoard 各表行数（供看板显示数据落地状态）。"""
    import pymysql
    try:
        conn = pymysql.connect(**APD22_TARGET_DB)
        cur = conn.cursor()
        out = {}
        for tbl in _APD22_TABLES_SQL:
            try:
                cur.execute(f"SELECT COUNT(*) FROM `{tbl}`")
                out[tbl] = int(cur.fetchone()[0])
            except Exception:
                out[tbl] = None
        conn.close()
        return {"ok": True, "tables": out}
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ============ MESE限速器汇总（车间呼报 · MESE呼报 · MESE限速器呼报，2026-11-03 接入）============
# 数据源：桌面「限速器」目录下**文件名带减号后缀**的批次导出文件
#         （--.xlsx / --最新.xlsx / -- (1).xlsx / -------.xlsx …，减号数量与尾巴都不限）
# 产出：  限速器\MESE限速器呼报制作\MESE限速器汇总.xlsx（末列「下达批次」= 源文件名 + 相对超链接）
# 逻辑全部在 mese_limiter.py；这里只做 HTTP 壳。

@app.get("/api/mese/limiter/summary")
def mese_limiter_summary(refresh: int = 0):
    """返回汇总表数据（meta + rows）。refresh=1 时重新扫盘。"""
    try:
        import mese_limiter as ml
        data = ml.summary(force=bool(refresh))
        return {"ok": True, "meta": data["meta"], "rows": data["rows"]}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}


@app.post("/api/mese/limiter/rebuild")
def mese_limiter_rebuild():
    """重新扫描源目录、重写 MESE限速器汇总.xlsx，并返回最新数据。"""
    try:
        import mese_limiter as ml
        data = ml.rebuild()
        return {"ok": True, "meta": data["meta"], "written": data["written"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"重新汇总失败：{type(e).__name__}: {e}")


@app.get("/api/mese/limiter/file")
def mese_limiter_file(name: str):
    """按文件名打开某个批次源表（页面上「下达批次」列的超链接目标）。"""
    import mese_limiter as ml
    p = ml.source_file_path(name)
    if not p:
        raise HTTPException(status_code=404, detail="源文件不存在，或文件名不在允许清单内")
    return FileResponse(
        str(p), filename=p.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache"},
    )


@app.get("/api/mese/limiter/export")
def mese_limiter_export():
    """下载汇总表 xlsx（不存在则按当前源目录即时生成）。"""
    import mese_limiter as ml
    try:
        p = ml.summary_file_path(generate=True)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成汇总表失败：{type(e).__name__}: {e}")
    return FileResponse(
        str(p), filename=p.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache"},
    )


@app.post("/api/mese/limiter/open")
def mese_limiter_open(name: str):
    """Ctrl+单击「下达批次」：在本机用 Excel 打开对应源表（等同 Excel 里的 Ctrl+单击超链接）。

    只认 mese_limiter 白名单里的文件名（必须命中减号命名规则且真实存在于源目录）；
    只读打开，不写任何东西；普通单击前端已拦掉，不会走到这里。
    """
    import mese_limiter as ml
    p = ml.source_file_path(name)
    if not p:
        raise HTTPException(status_code=404, detail="源文件不存在，或文件名不在允许清单内")
    try:
        os.startfile(str(p))          # 本机专用：交给 Windows 默认程序（Excel）打开
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"打开失败：{type(e).__name__}: {e}")
    return {"ok": True, "name": p.name, "path": str(p)}


# ---- MESE限速器汇总：可编辑本地存储（结构完全同「安全钳基础数据」store）----
# 页面「MESE限速器汇总」用它承载汇总结果，走通用工厂 makeStorePage，因此自动获得全站通用表格能力：
#   行号列 / 行首复选框 + Shift 范围勾选 / 点表头排序 / 列筛选漏斗 / 关键字查询（命中高亮）/
#   自定义查询抽屉 / 分页（右下角 NN 条/页）/ Ctrl+C 复制带表头 / 右键菜单与导出 /
#   双击右键切整行⇄单元格模式 / 列宽行高拖拽 / Esc 一键重置 / Ctrl+Z Ctrl+Y 撤销重做。
#   「重新汇总」= POST /api/mese/limiter/resync：重扫源目录 → 覆盖 store + 重写汇总表 xlsx。
MESE_LIMITER_STORE_PATH = os.path.join(APP_DIR, "mese_limiter_store.json")
MESE_LIMITER_SHEET = "MESE限速器汇总"


def _mese_limiter_rows_payload():
    """按 mese_limiter 汇总结果生成 store 行（id/row_no/data，data 的键=表头）。"""
    import mese_limiter as ml
    rows, meta = ml.read_rows()
    cols = meta["columns"]
    out = []
    for i, r in enumerate(rows, start=1):
        out.append({"id": i, "row_no": i,
                    "data": {c: ("" if v is None else str(v)) for c, v in zip(cols, r)}})
    return out, meta


def _mese_limiter_store_seed():
    """首次为空时用源目录即时汇总结果当种子（之后以 store 为准，重新汇总走 resync）。"""
    try:
        out, _ = _mese_limiter_rows_payload()
        return {"next_id": len(out) + 1, "sheets": {MESE_LIMITER_SHEET: out}}
    except Exception:
        return {"next_id": 1, "sheets": {MESE_LIMITER_SHEET: []}}


_mel_load, _mel_upsert, _mel_delete, _mel_replace = _store_routes(
    "mese_limiter", MESE_LIMITER_STORE_PATH, _mese_limiter_store_seed, "MESE限速器汇总")


@app.get("/api/mese/limiter/store")
def mese_limiter_store_get():
    """MESE限速器汇总本地存储全量（sheet → [{id,row_no,data}]）。"""
    return _mel_load()


@app.post("/api/mese/limiter/store/upsert")
def mese_limiter_store_upsert(payload: PermUpsert):
    return _mel_upsert(payload)


@app.post("/api/mese/limiter/store/delete")
def mese_limiter_store_delete(payload: dict):
    return _mel_delete(payload)


@app.post("/api/mese/limiter/store/bulk")
def mese_limiter_store_bulk(payload: dict):
    """批量更新多行 data（2026-11-04）：拖动填充 / 区域粘贴走这里，一次往返，
    避免「每格一次 upsert + 末尾整表重载」造成的页面闪烁。body: {sheet, updates:[{id,data}]}"""
    sheet = str((payload or {}).get("sheet") or "")
    updates = (payload or {}).get("updates") or []
    if not sheet or not isinstance(updates, list) or not updates:
        return {"ok": True, "count": 0}
    s = _generic_store_load(MESE_LIMITER_STORE_PATH) or {}
    rows = (s.setdefault("sheets", {})).setdefault(sheet, [])
    by_id = {r.get("id"): r for r in rows}
    n = 0
    for u in updates[:5000]:
        r = by_id.get((u or {}).get("id"))
        data = (u or {}).get("data")
        if r is None or not isinstance(data, dict):
            continue
        r["data"] = data
        n += 1
    if n:
        _generic_store_save(MESE_LIMITER_STORE_PATH, s)
        UI信息(f"限速器批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/mese/limiter/store/replace")
def mese_limiter_store_replace(payload: PermReplace):
    return _mel_replace(payload)


@app.post("/api/mese/limiter/resync")
def mese_limiter_resync():
    """重新扫描限速器源目录 → 覆盖 store（页面「重新汇总」）→ 同时重写 MESE限速器汇总.xlsx。"""
    try:
        import mese_limiter as ml
        out, meta = _mese_limiter_rows_payload()
        # 只覆盖「MESE限速器汇总」这一张表：同一 store 里的其它页签
        # （MESE临时（不齐-发群）/ MESE下载保留）必须原样保留，不能被重新汇总冲掉。
        store = _generic_store_load(MESE_LIMITER_STORE_PATH) or {}
        sheets = dict(store.get("sheets") or {})
        sheets[MESE_LIMITER_SHEET] = out
        _max_id = max([int(r.get("id") or 0) for s in sheets.values() for r in s] or [0])
        next_id = max(int(store.get("next_id") or 1), len(out) + 1, _max_id + 1)
        _generic_store_save(MESE_LIMITER_STORE_PATH, {"next_id": next_id, "sheets": sheets})
        rows, _meta2 = ml.read_rows()
        written = ml.write_workbook(rows)
        return {"ok": True, "file_count": meta.get("file_count"), "row_count": len(out),
                "written": str(written), "generated_at": meta.get("generated_at")}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"重新汇总失败：{type(e).__name__}: {e}")


# ============ MESE限速器「呼报制作」一键（2026-11-04 用户：把 VBA moudle108_组合main 搬到网页）============
# 原 VBA 位置：桌面\安全钳限速器铭牌数据\MESE散件_限速器.xlsm → moudle108_组合main.主组合流程
#   moudle103 下载3张原表（SE16 ZPP_PO ×2 + SQ00 ZB001）+ moudle104 导入 + moudle105 呼报制作
#   + moudle107 typ/pul/spd + moudle106 调试参数匹配 + 检查空值
# 网站版分工：
#   步骤1（SAP 下载）→ 复用站内既有 VBS（vbs_scripts/Zpp_Po.vbs、Sq00.vbs，字段 ID 与原 VBA 的同一张 SE16/SQ00 屏幕）
#   步骤2~6（数据加工）→ app/mese_limiter_build.py（逐段移植，已用 2026-11-02 真实产物 54 行逐列对账一致）
# 结果写进「MESE临时（不齐-发群）」页签（＝不齐先发群的暂存区），再由页面的「发送钉钉群」发群。
MESE_COMBINE_SHEET = "MESE临时（不齐-发群）"

# 🔴🔴 2026-11-06 用户明令：「MESE限速器以及MESE安全钳 我不需要你发邮件给我，改下代码」
#   总闸：MESE 两条线一律不往演示用户本人邮箱（demo.user70@example.com）发任何邮件。
#   实证：2026-11-06 10:45~10:46 连发 5 封「Zpp_Po（ZPP_PO） 查询结果【成功】」
#         /「SQ00配方查询（ZB001） 查询结果【成功】」，日志对应
#         「zpp_po 报表邮件（下载开关触发，发 songjian）」—— 是「呼报制作」内部跑 SAP
#         三张原表时，VbsRunRequest.报表邮件 默认 True 触发的自发自收结果邮件。
#   🔴 只掐「发给本人」的邮件：发给车间同事的业务邮件（mese_pipeline.py 的 MAIL_TO_*）
#      照旧保留（用户 11-06 明令「发邮件的代码不要改」），钉钉单聊也照旧。
MESE_不发本人邮件 = True
MESE_本人邮箱 = "demo.user71@example.com"


def _mese_去本人邮箱(地址串: str) -> str:
    """从收件人/抄送里剔除演示用户本人邮箱；分隔符口径与 _拆分收件人 一致（逗号/分号/空格）。

    总闸 MESE_不发本人邮件=False 时原样返回（方便用户本人临时测试收件时一处切回）。
    """
    if not MESE_不发本人邮件:
        return str(地址串 or "")
    import re as _re
    段 = [x.strip() for x in _re.split(r"[,;\s]+", str(地址串 or "")) if x.strip()]
    return ", ".join(x for x in 段 if x.lower() != MESE_本人邮箱)

# 「发完群 @ 的人」：**必须用 openDingTalkId**（不是 userId），否则钉钉只当普通文字。
# 来源：`dws contact user search --query 朱诚 / 胡旺来`（2026-11-04 实测查出）：
#   朱诚（科员，userId 07001）      → DDqjQOLradAAhMvCiPHiSNw1k03AOjgiSFcE
#   胡旺来（200785三菱电机上海机电电梯有限公司）→ demoUserIdPlaceholder
MESE_LIMITER_AT_IDS = [
    "DDqjQOLradAAhMvCiPHiSNw1k03AOjgiSFcE",
    "demoUserIdPlaceholder",
]
MESE_LIMITER_AT_TEXT = "今日 MESE限速器呼报（{n} 行）"


class MeseCombineRequest(BaseModel):
    download: bool = True        # True＝先跑 SAP 下载三张原表；False＝直接用「限速器原表」目录里最新的原表重算
    append: bool = False         # False＝覆盖本页签（与 VBA 清表重算同语义）；True＝追加到后面


# ZPP_PO 选择屏（SE16）要清空的注入键 —— 对应 moudle100 aaa() 里「只有 实际下达=今天」的查询范围。
# 不清空的话会带上页面脚本自带默认值（生产调度员 24*、物料组 002、MRP 303、计划日期区间…），查出来的就不是今天全部下达。
_ZPP_PO_清空键 = [
    "delflagALow", "delflagAHigh", "delflagBLow", "delflagBHigh",
    "fevorLow", "fevorHigh", "aufnrLow", "aufnrHigh", "plnumLow", "plnumHigh",
    "posnrLow", "posnrHigh", "matnrLow", "matnrHigh", "matklLow", "matklHigh",
    "dispoLow", "dispoHigh", "stlanLow", "stlanHigh", "pstartLow", "pstartHigh",
    "pdelvLow", "pdelvHigh", "gstrpLow", "gstrpHigh", "gltrpLow", "gltrpHigh",
    "fstrtLow", "fstrtHigh", "fendenLow", "fendenHigh", "ladatLow", "ladatHigh",
    "wzeitLow", "wzeitHigh", "lgortLow", "lgortHigh", "teckoLow", "teckoHigh",
    "lifskLow", "lifskHigh", "psmngLow", "psmngHigh", "wemngLow", "wemngHigh",
    "meinsLow", "meinsHigh", "quanLow", "quanHigh", "ueberLow", "ueberHigh",
    "maktxLow", "maktxHigh", "sprasLow", "sprasHigh", "plnnumLow", "plnnumHigh",
    "plnnrLow", "plnnrHigh", "ltxspLow", "ltxspHigh", "erdatLow", "erdatHigh",
]


def _mese_combine_vbs(script: str, params: dict, raw_dir: "Path", 前缀: str, 时间戳: str) -> "Path":
    """跑一次站内 VBS 导出，拿到报表后改名成原 VBA 的命名（zpp_poDM_/sq00配方查询_<时间戳>.xls）。"""
    try:
        # 🔴🔴 2026-11-06 用户：「MESE限速器以及MESE安全钳 我不需要你发邮件给我」
        #   这里只是「呼报制作」内部下载三张原表，绝不能顺手发结果邮件给演示用户本人。
        #   之前没传 报表邮件 → 默认 True，导致每跑一张表就自发自收一封「查询结果【成功】」
        #   （2026-11-06 10:45~10:46 实测 5 封）。显式关掉，并加一道总闸兜底。
        r = vbs_run(VbsRunRequest(script=script, params=params, download=True,
                                  下载目录=str(raw_dir),
                                  mail=False, 钉钉=False,
                                  报表邮件=False,
                                  不发本人邮件=MESE_不发本人邮件))
    except HTTPException as e:
        raise RuntimeError(f"{script} 启动失败：{e.detail}") from e
    except Exception as e:                                     # noqa: BLE001
        raise RuntimeError(f"{script} 执行异常：{type(e).__name__}: {e}") from e
    if not r.get("ok"):
        尾巴 = "\n".join((r.get("output") or "").splitlines()[-6:])
        raise RuntimeError(f"{script} 运行失败：{尾巴 or '（无输出）'}")
    dl = r.get("download") or {}
    p = Path(str(dl.get("路径") or ""))
    if not dl.get("ok") or not p.is_file():
        raise RuntimeError(f"{script} 没拿到导出文件：{dl.get('vbs') or dl}")
    新路径 = raw_dir / f"{前缀}{时间戳}.xls"
    try:
        if p.resolve() != 新路径.resolve():
            p.replace(新路径)
    except Exception:                                          # noqa: BLE001
        return p                                               # 改名失败就用原名，后续解析不依赖文件名
    return 新路径


def _mese_clipboard_materials() -> List[str]:
    """读「物料编号」条件 —— 与 VBA moudle100 一致：取**剪贴板内容**（用户在 Excel 里复制的那份物料清单）。

    VBA 里的动作是：打开 SE16 ZPP_PO → 点物料编号的多重选择 → 「从剪贴板导入」→ 执行。
    网站版等价做法：读 Windows 剪贴板文本 → 按 换行/制表符/逗号/分号/空格 切成值 → 去重保序 →
    交给 BatchWrite 写进同一个多重选择框（逗号分隔多值，SAP 侧语义相同）。
    """
    raw = _读剪贴板文本() or ""
    值们: List[str] = []
    见过 = set()
    for 段 in _re.split(r"[\s,;，；、\t]+", raw):
        v = (段 or "").strip().strip('"').strip("'")
        if v and v not in 见过:
            见过.add(v)
            值们.append(v)
    return 值们


@app.post("/api/mese/limiter/combine")
def mese_limiter_combine(payload: MeseCombineRequest):
    """一键呼报制作（VBA moudle108_组合main 的网站版）→ 结果写进「MESE临时（不齐-发群）」页签。

    逐步日志与提醒（空值字段/未匹配/SPD 不一致）原样回传页面，和 VBA 里那些 MsgBox 一一对应。
    """
    import mese_limiter_build as mb

    log: List[str] = []
    warn: List[str] = []
    files: Dict[str, str] = {}
    try:
        raw_dir = mb.RAW_DIR
        raw_dir.mkdir(parents=True, exist_ok=True)
        现在 = datetime.now()
        今天 = 现在.strftime("%Y.%m.%d")
        时间戳 = 现在.strftime("%Y%m%d_%H%M%S")

        if payload.download:
            # ---------- 步骤1-a：SE16 ZPP_PO（实际下达=今天 + 物料编号=剪贴板内容）→ zpp_poDM_*.xls ----------
            # 物料来源与 VBA moudle100 一致：用户先把物料清单复制到剪贴板，脚本把它粘进多重选择框。
            剪贴板物料 = _mese_clipboard_materials()
            if not 剪贴板物料:
                raise RuntimeError(
                    "物料编号要用剪贴板内容，但剪贴板是空的：请先复制物料清单（一行一个）再点按钮"
                    "（对应 VBA 里「从剪贴板导入」那一步）")
            dm_物料 = [m for m in 剪贴板物料 if mb.is_dm_material(m)]
            log.append(f"① 物料编号条件：剪贴板 {len(剪贴板物料)} 个值（样例 {', '.join(剪贴板物料[:3])}…），"
                       f"其中 DM 级物料 {len(dm_物料)} 个")
            if not dm_物料:
                raise RuntimeError(f"剪贴板里的 {len(剪贴板物料)} 个值都不是 DM 级物料（DM 开头），"
                                   "请确认复制的是限速器总成的物料清单")

            zpp_list: List[List[str]] = []
            zpp_hdr: List[str] = []
            for i, 块 in enumerate([剪贴板物料[j:j + 180] for j in range(0, len(剪贴板物料), 180)], 1):
                参数 = {k: "" for k in _ZPP_PO_清空键}
                参数.update({"matnrLow": ",".join(块), "freizLow": 今天, "freizHigh": 今天})
                p_zpp = _mese_combine_vbs("zpp_po", 参数, raw_dir, f"zpp_poDM_{i}_" if len(剪贴板物料) > 180 else "zpp_poDM_",
                                           f"{时间戳}" if len(剪贴板物料) <= 180 else f"{时间戳}_{i}")
                h, rs = mb.parse_se16_tsv(p_zpp)
                zpp_hdr = zpp_hdr or h
                zpp_list.extend(rs)
                log.append(f"   SE16 ZPP_PO 第 {i} 批（物料 {len(块)} 个，实际下达 {今天}）→ {p_zpp.name}：{len(rs)} 条")
                files["zpp_po"] = str(p_zpp)
            # 多批合并时按「订单 + 订单项目编号 + 物料编号」去重（防同一订单在两批里重复展开）
            _k = {n: mb._col_index(zpp_hdr, (n,), fb) for n, fb in
                  (("订单", mb.DM_COL_K), ("订单项目编号", 13), ("物料编号", mb.DM_COL_P))}
            _seen = set()
            zpp_rows = []
            for cells in zpp_list:
                key = tuple(mb._cell(cells[idx]) if len(cells) > idx else "" for idx in _k.values())
                if key in _seen:
                    continue
                _seen.add(key)
                zpp_rows.append(cells)
            if not zpp_rows:
                raise RuntimeError(f"ZPP_PO 查询 0 条（物料 {len(剪贴板物料)} 个 + 实际下达 {今天}）："
                                   "今天这批物料没有下达记录")
            log.append(f"   ZPP_PO 合计 {len(zpp_rows)} 条（去重后）")

            # ---------- 步骤1-b：SQ00 ZB001（物料=同一份剪贴板物料）→ sq00配方查询_*.xls（moudle101）----------
            sq00_rows: List[List[str]] = []
            分块 = [剪贴板物料[i:i + 180] for i in range(0, len(剪贴板物料), 180)]
            for i, 块 in enumerate(分块, 1):
                参数 = {
                    "sq00MatnrLow": ",".join(块),
                    "sq00MatnrHigh": "", "sq00BomCompLow": "*", "sq00BomCompHigh": "",
                    "sq00ItemLow": "115", "sq00ItemHigh": "",          # 115 = 限速器品目（不是 125 安全钳）
                    "sq00ItemGrpLow": "G*", "sq00ItemGrpHigh": "",
                    "sq00LNoLow": "*", "sq00LNoHigh": "",
                    "sq00LongParamLow": "*", "sq00LongParamHigh": "",
                    "sq00DrawNoLow": "*", "sq00DrawNoHigh": "",
                    "sq00PartNameLow": "", "sq00PartNameHigh": "",     # 脚本默认 *安全钳*，限速器必须清空
                    "sq00BomFlagLow": "", "sq00BomFlagHigh": "",
                    "sq00CreatedLow": "", "sq00CreatedHigh": "",
                    "sq00CreatorLow": "", "sq00CreatorHigh": "",
                }
                p_sq = _mese_combine_vbs("sq00", 参数, raw_dir,
                                         "sq00配方查询_" if len(分块) == 1 else f"sq00配方查询_{i}_",
                                         f"{时间戳}" if len(分块) == 1 else f"{时间戳}_{i}")
                sq00_rows.extend(mb.parse_sq00_tsv(p_sq))
                log.append(f"② SQ00 ZB001（物料 {len(块)} 个"
                           + (f"，第 {i}/{len(分块)} 批" if len(分块) > 1 else "")
                           + f"）→ {p_sq.name}：{len(sq00_rows)} 条")
            if len(分块) > 1:
                warn.append(f"剪贴板物料超过 180 个，SQ00 已分 {len(分块)} 批下载（合并解析）")

            # ---------- 步骤1-c：ZPP_PO 第二次下载（物料 = sq00 的组件物料）→ zpp_po通用件级别_*.xls（moudle102）----------
            # 必须单独查一次：按物料过滤后的下载里不会包含另一类物料的行，
            # 所以「通用件级」不能靠第一次下载筛出来（2026-11-04 现场：KHT 列全空就是这么来的）。
            组件物料 = mb.component_materials(sq00_rows)
            ty_list: List[List[str]] = []
            ty_hdr: List[str] = list(zpp_hdr)
            if 组件物料:
                for i, 块 in enumerate([组件物料[j:j + 180] for j in range(0, len(组件物料), 180)], 1):
                    参数 = {k: "" for k in _ZPP_PO_清空键}
                    参数.update({"matnrLow": ",".join(块), "freizLow": 今天, "freizHigh": 今天})
                    p_ty = _mese_combine_vbs("zpp_po", 参数, raw_dir,
                                             "zpp_po通用件级别_" if len(组件物料) <= 180 else f"zpp_po通用件级别_{i}_",
                                             f"{时间戳}" if len(组件物料) <= 180 else f"{时间戳}_{i}")
                    h2, rs2 = mb.parse_se16_tsv(p_ty)
                    ty_hdr = h2 or ty_hdr
                    ty_list.extend(rs2)
                    log.append(f"③ SE16 ZPP_PO 通用件级别（物料 {len(块)} 个"
                               + (f"，第 {i}/{len(组件物料)//180 + 1} 批" if len(组件物料) > 180 else "")
                               + f"）→ {p_ty.name}：{len(rs2)} 条")
                    files["zpp_po_ty"] = str(p_ty)
            else:
                warn.append("sq00 配方查询没有组件物料，通用件级为空（KHT/WTH 列会空）")

            _kt = {n: mb._col_index(ty_hdr, (n,), fb) for n, fb in
                   (("订单", mb.DM_COL_K), ("订单项目编号", 13), ("物料编号", mb.DM_COL_P))}
            _seen2 = set()
            ty_rows = []
            for cells in ty_list:
                key = tuple(mb._cell(cells[idx]) if len(cells) > idx else "" for idx in _kt.values())
                if key in _seen2:
                    continue
                _seen2.add(key)
                ty_rows.append(cells)

            # DM 级原表 = 第一次下载里属于「剪贴板物料 ∩ DM 级」的行（moudle100 口径）
            dm_rows = mb.pick_dm_rows(zpp_hdr, zpp_rows, dm_set=set(剪贴板物料))
            log.append(f"④ 拆分原表：DM 级 {len(dm_rows)} 条 / 通用件级 {len(ty_rows)} 条")
            if not ty_rows:
                warn.append(f"通用件级 0 条：sq00 的组件物料（{len(组件物料)} 个）在今天没有下达记录"
                            "（此时「通用组件订单 / 唯一码」两列会空）")

            menu, menu_from = mb.load_menu()
            rows, steps, warns = mb.build_tables(zpp_hdr, dm_rows, sq00_rows, ty_hdr, ty_rows, menu, now=现在)
            log.extend(steps)
            warn.extend(warns)
            log.insert(0, f"参数表来源：{menu_from}")
        else:
            r = mb.combine(now=现在)
            if not r.get("ok"):
                raise RuntimeError(r.get("error") or "原表缺失")
            log.extend(r["log"])
            warn.extend(r["warn"])
            rows = r["rows"]
            files.update(r.get("files") or {})

        if not rows:
            raise RuntimeError("呼报制作结果 0 行（原表里没有可展开的 DM 级订单）")

        # ---------- 写进「MESE临时（不齐-发群）」页签（data 键 = 页面 27 列）----------
        store = _generic_store_load(MESE_LIMITER_STORE_PATH) or {}
        sheets = dict(store.get("sheets") or {})
        旧行 = sheets.get(MESE_COMBINE_SHEET) or []
        max_id = max([int(r.get("id") or 0) for s in sheets.values() for r in s] or [0])
        next_id = max(int(store.get("next_id") or 1), max_id + 1)
        批次 = rows[0].get("通知批次", "")

        def _封装(row: dict, rid: int, rno: int) -> dict:
            data = {h: row.get(h, "") for h in mb.HEADERS}
            # 「下达批次」= 网站给汇总表用的源文件溯源列；组合main 不产生它（VBA 的目标表只有 A..Z），
            # 发群附件也不带这一列（2026-11-04 用户口径），所以这里留空，别把批次值塞进去。
            data["下达批次"] = ""
            return {"id": rid, "row_no": rno, "data": data}

        if payload.append:
            base_no = max([int(r.get("row_no") or 0) for r in 旧行], default=0)
            新行 = [_封装(row, next_id + i, base_no + i + 1) for i, row in enumerate(rows)]
            sheets[MESE_COMBINE_SHEET] = 旧行 + 新行
            next_id += len(rows)
        else:
            新行 = [_封装(row, next_id + i, i + 1) for i, row in enumerate(rows)]
            sheets[MESE_COMBINE_SHEET] = 新行
            next_id += len(rows)
        _generic_store_save(MESE_LIMITER_STORE_PATH, {"next_id": next_id, "sheets": sheets})
        log.append(f"④ 已写入「{MESE_COMBINE_SHEET}」{len(rows)} 行"
                   + ("（追加）" if payload.append else "（覆盖本页签）"))

        return {"ok": True, "sheet": MESE_COMBINE_SHEET, "rows": len(rows), "batch": 批次,
                "log": log, "warn": warn, "files": files,
                "replace_old": 0 if payload.append else len(旧行)}
    except Exception as e:                                     # noqa: BLE001
        log.append(f"✗ {type(e).__name__}: {e}")
        return {"ok": False, "detail": f"{type(e).__name__}: {e}", "log": log, "warn": warn, "files": files}


@app.post("/api/mese/limiter/precheck")
def mese_limiter_precheck(payload: dict | None = None):
    """「发送钉钉群」前的检查（2026-11-04 用户口径：备注 / WTH 要么都有值、要么都没有）。

    只读，不改数据。有问题时页面**不允许发送**（硬拦）。
    """
    import mese_limiter_build as mb
    store = _generic_store_load(MESE_LIMITER_STORE_PATH) or {}
    行们 = (store.get("sheets") or {}).get(MESE_COMBINE_SHEET) or []
    数据行 = []
    for r in sorted(行们, key=lambda x: int(x.get("row_no") or 0)):
        d = r.get("data") or {}
        数据行.append({h: str(d.get(h) or "") for h in mb.HEADERS})
    problems, summary = mb.precheck(数据行)
    return {"ok": True, "passed": not problems, "problems": problems[:500],
            "total": len(problems), "summary": summary}


@app.post("/api/mese/limiter/send-file")
def mese_limiter_send_file(payload: dict):
    """「发送钉钉群」= 发**附件**（与 VBA moudle109 + ab1.发送文件消息 同一口径）。

    1) 把「MESE临时（不齐-发群）」页签内容写成 xlsx：
       文件名 MESE-115计划__N个_YYYYMMDD_HHMMSS.xlsx，落 桌面\\安全钳限速器铭牌数据\\限速器\\
       **只写 26 列**（通知批次…zl）——「下达批次」是网站溯源列，moudle109 的表里没有、发群附件也不带。
    2) dws chat +messages-send --file 发到群（站内 dws_msg.发送文件消息）。
    body: {group: 群名, dry_run?: 只生成附件不发送}
    """
    import mese_limiter_build as mb

    group = str((payload or {}).get("group") or "").strip()
    dry = bool((payload or {}).get("dry_run"))
    if not group:
        raise HTTPException(status_code=400, detail="请选择要发送的钉钉群")
    if _re.search(r"[;；,，\n]", group):
        raise HTTPException(status_code=400, detail="群名里不能带分隔符")

    store = _generic_store_load(MESE_LIMITER_STORE_PATH) or {}
    行们 = (store.get("sheets") or {}).get(MESE_COMBINE_SHEET) or []
    if not 行们:
        raise HTTPException(status_code=400, detail=f"「{MESE_COMBINE_SHEET}」没有数据，先「组合main」生成再发")
    数据行 = []
    for r in sorted(行们, key=lambda x: int(x.get("row_no") or 0)):
        d = r.get("data") or {}
        数据行.append({h: str(d.get(h) or "") for h in mb.HEADERS})

    try:
        附件 = mb.write_batch_xlsx(数据行, now=datetime.now())
    except Exception as e:                                     # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"生成附件失败：{type(e).__name__}: {e}")

    if dry:
        return {"ok": True, "dry_run": True, "file": str(附件), "name": 附件.name,
                "rows": len(数据行), "columns": len(mb.HEADERS), "group": group}

    mod = _rpa导dws_msg()
    if not mod:
        raise HTTPException(status_code=503, detail="钉钉消息不可用（web_auto/公共_钉钉通知/dws_msg.py 是否完整）")
    try:
        状态, 说明 = mod.发送文件消息(str(附件), f"{MESE_COMBINE_SHEET} · {len(数据行)} 行", "群:" + group)
    except Exception as e:                                     # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"发送失败：{type(e).__name__}: {e}")
    ok = str(状态).startswith("已发送")

    # 2026-11-04 用户：贴附件之后，群里再发一条 @ 指定的人（与安全钳页「末尾消息」同一做法：
    # dws 文本里写 @昵称 即真正的 @）。附件没发成功就不 @，免得只 @ 不带文件。
    末尾消息 = None
    if ok:
        try:
            文 = MESE_LIMITER_AT_TEXT.format(n=len(数据行))
            发送并at = getattr(mod, "发送群消息并at", None)
            if callable(发送并at):
                状态2, 说明2 = 发送并at(文, group, MESE_LIMITER_AT_IDS)     # 真 @（--at-open-dingtalk-ids）
            else:
                状态2, 说明2 = mod.发送文本消息(文, "群:" + group)          # 老版本退回：至少发得出文本
            末尾消息 = {"ok": str(状态2).startswith("已发送"), "status": str(状态2), "text": 文,
                        "msg": str(说明2)[:200]}
        except Exception as e:                                 # noqa: BLE001
            末尾消息 = {"ok": False, "status": "失败", "msg": f"{type(e).__name__}: {e}"}

    return {"ok": ok, "status": str(状态), "msg": str(说明)[:500], "group": group,
            "file": str(附件), "name": 附件.name, "rows": len(数据行), "columns": len(mb.HEADERS),
            "at": 末尾消息}


@app.post("/api/mese/safety/open")
def mese_safety_open(name: str):
    """Ctrl+单击「下达批次」：在本机用 Excel 打开对应安全钳源表。"""
    import mese_safety as ms
    p = ms.source_file_path(name)
    if not p:
        raise HTTPException(status_code=404, detail="源文件不存在，或文件名不在允许清单内")
    try:
        os.startfile(str(p))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"打开失败：{type(e).__name__}: {e}")
    return {"ok": True, "name": p.name, "path": str(p)}


# ============ MESE安全钳汇总计划（车间呼报 · MESE呼报 · MESE安全钳呼报 → MESE安全钳汇总计划，2026-11-03 接入）============
# 数据源：桌面「安全钳」目录下**文件名带减号后缀**的批次打印文件
#         （-.xlsx / --.xlsx / ---.xlsx / -- - 副本.xlsx / -2--.xlsx …，减号数量与尾巴都不限）
#         另按用户指定纳入 5 个无减号副本（见 mese_safety.EXTRA_NAMES）
# 产出：  安全钳\MESE安全钳呼报\MESE安全钳汇总计划.xlsx（末列「下达批次」= 源文件名 + 相对超链接）
# 逻辑全部在 mese_safety.py；这里只做 HTTP 壳。

@app.get("/api/mese/safety/summary")
def mese_safety_summary(refresh: int = 0):
    """返回 MESE安全钳汇总计划数据（meta + rows）。refresh=1 时重新扫盘。"""
    try:
        import mese_safety as ms
        data = ms.summary(force=bool(refresh))
        return {"ok": True, "meta": data["meta"], "rows": data["rows"]}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}


@app.post("/api/mese/safety/rebuild")
def mese_safety_rebuild():
    """重新扫描安全钳源目录、重写 MESE安全钳汇总计划.xlsx，并返回最新数据。"""
    try:
        import mese_safety as ms
        data = ms.rebuild()
        return {"ok": True, "meta": data["meta"], "written": data["written"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"重新汇总失败：{type(e).__name__}: {e}")


@app.get("/api/mese/safety/file")
def mese_safety_file(name: str):
    """按文件名打开某个批次源表（页面上「下达批次」列的超链接目标）。"""
    import mese_safety as ms
    p = ms.source_file_path(name)
    if not p:
        raise HTTPException(status_code=404, detail="源文件不存在，或文件名不在允许清单内")
    return FileResponse(
        str(p), filename=p.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache"},
    )


@app.get("/api/mese/safety/export")
def mese_safety_export():
    """下载汇总表 xlsx（不存在则按当前源目录即时生成）。"""
    import mese_safety as ms
    try:
        p = ms.summary_file_path(generate=True)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成汇总表失败：{type(e).__name__}: {e}")
    return FileResponse(
        str(p), filename=p.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache"},
    )


# ---- MESE安全钳汇总：可编辑本地存储（结构完全同「安全钳基础数据」store）----
# 页面「MESE安全钳汇总」（车间呼报 · MESE呼报 · MESE安全钳呼报 下一级）用它承载汇总结果：
#   双击改单元格 / 新增行 / 删除 / 删除所选 / 撤销重做（Ctrl+Z Ctrl+Y）全部按全站通用规则走；
#   「重新汇总」按钮 = POST /api/mese/safety/resync：重扫源目录 → 覆盖 store + 重写汇总表 xlsx。
MESE_SAFETY_STORE_PATH = os.path.join(APP_DIR, "mese_safety_store.json")
MESE_SAFETY_SHEET = "MESE安全钳汇总"


def _mese_safety_rows_payload():
    """按 mese_safety 汇总结果生成 store 行（id/row_no/data，data 的键=表头）。"""
    import mese_safety as ms
    rows, meta = ms.read_rows()
    cols = meta["columns"]
    out = []
    for i, r in enumerate(rows, start=1):
        out.append({"id": i, "row_no": i,
                    "data": {c: ("" if v is None else str(v)) for c, v in zip(cols, r)}})
    return out, meta


def _mese_safety_store_seed():
    """首次为空时用源目录即时汇总结果当种子（之后以 store 为准，重新汇总走 resync）。"""
    try:
        out, _ = _mese_safety_rows_payload()
        return {"next_id": len(out) + 1, "sheets": {MESE_SAFETY_SHEET: out}}
    except Exception:
        return {"next_id": 1, "sheets": {MESE_SAFETY_SHEET: []}}


_mesms_load, _mesms_upsert, _mesms_delete, _mesms_replace = _store_routes(
    "mese_safety", MESE_SAFETY_STORE_PATH, _mese_safety_store_seed, "MESE安全钳汇总")


@app.get("/api/mese/safety/store")
def mese_safety_store_get():
    """MESE安全钳汇总本地存储全量（sheet → [{id,row_no,data}]）。"""
    return _mesms_load()


@app.post("/api/mese/safety/store/upsert")
def mese_safety_store_upsert(payload: PermUpsert):
    return _mesms_upsert(payload)


@app.post("/api/mese/safety/store/delete")
def mese_safety_store_delete(payload: dict):
    return _mesms_delete(payload)


@app.post("/api/mese/safety/store/replace")
def mese_safety_store_replace(payload: PermReplace):
    return _mesms_replace(payload)


@app.post("/api/mese/safety/store/reorder")
def mese_safety_store_reorder(payload: dict):
    """按给定 id 顺序重排「MESE临时（不齐-发群）」等页签的行序，并重写行号 1..N。
    保留原 id、不动任何单元格内容 —— 用于「按安全钳物料号排序、相同物料号归拢」。
    """
    sheet = payload.get("sheet") or "MESE临时（不齐-发群）"
    order = payload.get("order") or []
    r = _generic_store_reorder(MESE_SAFETY_STORE_PATH, sheet, order)
    UI信息(f"MESE安全钳重排行序 {sheet}（{r.get('count')} 行）")
    return r


@app.post("/api/mese/safety/resync")
def mese_safety_resync():
    """重新扫描安全钳源目录 → 覆盖 store（页面「重新汇总」）→ 同时重写 MESE安全钳汇总计划.xlsx。"""
    try:
        import mese_safety as ms
        out, meta = _mese_safety_rows_payload()
        # 只覆盖「MESE安全钳汇总」这一张表：同一 store 里的其它页签
        # （MESE临时（不齐-发群）/ MESE下载保留）必须原样保留，不能被重新汇总冲掉。
        store = _generic_store_load(MESE_SAFETY_STORE_PATH) or {}
        sheets = dict(store.get("sheets") or {})
        sheets[MESE_SAFETY_SHEET] = out
        _max_id = max([int(r.get("id") or 0) for s in sheets.values() for r in s] or [0])
        next_id = max(int(store.get("next_id") or 1), len(out) + 1, _max_id + 1)
        _generic_store_save(MESE_SAFETY_STORE_PATH, {"next_id": next_id, "sheets": sheets})
        rows, _meta2 = ms.read_rows()
        written = ms.write_workbook(rows)
        return {"ok": True, "file_count": meta.get("file_count"), "row_count": len(out),
                "written": str(written), "generated_at": meta.get("generated_at")}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"重新汇总失败：{type(e).__name__}: {e}")


# ---- MESE 汇总页「发送钉钉群」（车间呼报：临时不齐 → 发群）----
# 前端把当前页签内容拼成纯文本发过来，这里只负责投递到指定群（走 web_auto/公共_钉钉通知/dws_msg）。
@app.post("/api/mese/notify")
def mese_notify(payload: dict):
    """把当前页签内容发到钉钉群。payload: {group: 群名, text: 消息正文}"""
    import re as _re
    text = str(payload.get("text") or "").strip()
    group = str(payload.get("group") or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail="没有可发送的内容（当前页签是空的）")
    if not group:
        raise HTTPException(status_code=400, detail="请选择要发送的钉钉群")
    if _re.search(r"[;；,，\n]", group):
        raise HTTPException(status_code=400, detail="群名里不能带分隔符")
    mod = _rpa导dws_msg()
    if not mod:
        raise HTTPException(status_code=503, detail="钉钉消息不可用（web_auto/公共_钉钉通知/dws_msg.py 是否完整）")
    try:
        状态, 说明 = mod.发送文本消息(text, "群:" + group)
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"发送失败：{type(e).__name__}: {e}")
    ok = str(状态).startswith("已发送")
    return {"ok": ok, "status": str(状态), "msg": str(说明)[:500], "group": group, "bytes": len(text)}


# ---- MESE安全钳呼报「发送钉钉群」一键：校验 → 三页签截图 → 发群 → 归档 → @消息 ----
# 逻辑在 mese_report.py（口径全在 mese_report_config.json）；这里只做 HTTP 壳。
# 校验不通过时返回 409 + 问题行清单，前端抬头弹窗展示，**不发送**。
@app.post("/api/mese/safety/report")
def mese_safety_report(payload: dict | None = None):
    """body: {only_check?: bool}。发前校验「假想合同号 / 安全钳物料号」配对，再截图发群。"""
    only_check = bool((payload or {}).get("only_check"))
    try:
        import mese_report as mr
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"导入 mese_report 失败：{type(e).__name__}: {e}")
    store = _generic_store_load(MESE_SAFETY_STORE_PATH) or {}
    sheets = store.get("sheets") or {}
    try:
        return mr.run(sheets, 只校验=only_check)
    except mr.校验未通过 as exc:
        raise HTTPException(status_code=409, detail={
            "kind": "validate",
            "summary": exc.summary,
            "problems": exc.problems,
        })
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"发送失败：{type(e).__name__}: {e}")


# ---- MESE安全钳「MESE临时（不齐-发群）」→ 发给演示用户（2026-11-06 用户改口径）----
# 旧流程 mese_report.py（发前校验 + 三页签截图 + 文本/截图发群）不再由本按钮调用。
# 现在＝按参考格式生成 Excel 附件（固定带「数据库」sheet）→ **发给演示用户钉钉单聊** → 再发一条普通单聊文本。
# 🔴 2026-11-06 用户：「先别发到钉钉群，先改到发给演示用户，@不成功先不用管」→ 不发群、不 @。
#    投递目标常量在 mese_safety_send.py（SEND_TO / SEND_TO_LABEL / 发群）；本路由只做校验 + HTTP 壳。
# 🔴 邮件代码（mese_pipeline.py 的 MAIL_TO_SAFETY / MAIL_CC_SAFETY）**不要改**，仍是原收件人。
MESE_SAFETY_TMP_SHEET = "MESE临时（不齐-发群）"


# 🔴🔴 2026-11-06 用户明令：「邮件、钉钉群消息我自己测试发送，你不要随便瞎发」
#   所有自动化验收脚本都用 auth_config.json 的 test_token 作 X-Auth-Token；
#   凡是这个 token 发起的「发送类」请求，**无论 body 写没写 dry_run，一律强制预演**。
#   用户本人从浏览器点按钮走登录会话、不带这个头 → 真实发送完全不受影响。
#   这是系统级保险：即使某个脚本忘了填 dry_run、或真机点了发送按钮，也发不出去。
def _自动化请求(request: Request) -> bool:
    try:
        import json as _json
        from pathlib import Path as _Path
        cfg = _json.loads((_Path(__file__).parent / "auth_config.json").read_text(encoding="utf-8"))
        tok = str(cfg.get("test_token") or "")
        return bool(tok) and request.headers.get("x-auth-token", "") == tok
    except Exception:                                              # noqa: BLE001
        return False


@app.post("/api/mese/safety/send-file")
def mese_safety_send_file(request: Request, payload: dict | None = None):
    """body: {dry_run?: bool}。收件目标固定＝演示用户钉钉单聊（当前 dws 登录用户本人），不弹窗选择、不发群。

    dry_run=True 只生成预览附件并回报将要发送的内容，绝不调用 dws（测试用，不发消息打扰别人）。
    当前页签没有数据 / 附件没发成功 → 不发任何消息。

    🔴 2026-11-05 用户明令：发之前先校验「假想合同号」和「安全钳物料号」两列
       **要么同时都有值、要么同时都没有值**；不满足就**不发**（连附件都不生成），
       返回 409 + 问题行清单，由前端弹窗告诉用户。
    """
    import mese_safety_send as mss
    # 🔴🔴 自动化 token 强制预演（用户 2026-11-06：测试不许真发钉钉群打扰别人）
    dry = bool((payload or {}).get("dry_run")) or _自动化请求(request)
    store = _generic_store_load(MESE_SAFETY_STORE_PATH) or {}
    行们 = (store.get("sheets") or {}).get(MESE_SAFETY_TMP_SHEET) or []
    # 🔴 2026-11-05 用户驳回：「附件里面没排序」→ 送进附件生成器之前**先按「安全钳物料号」排一遍**
    #   （自然序；相同物料号相邻）。绝对不能只按 row_no 排 —— store 里的行序曾经被搅乱过
    #   （老批次残留 + 新行追加到表尾），那样打出来的附件就是乱的。
    #   比较口径与附件生成器共用 mese_safety_send 的同一套函数，避免两处口径走偏。
    行们 = sorted(行们, key=lambda x: (mss.原始行物料号排序键(x), int(x.get("row_no") or 0)))
    # 🔴 发前闸门：两列必须成对（同时有值 / 同时为空）。不通过 → 409，不生成附件、不发任何消息。
    行数据 = mss.非空行数据(行们)
    问题 = mss.配对问题(行数据, MESE_SAFETY_TMP_SHEET)
    if 问题:
        raise HTTPException(status_code=409, detail={
            "kind": "validate",
            "summary": {"页签": {MESE_SAFETY_TMP_SHEET: {"行数": len(行数据)}},
                        "检查行数": len(行数据)},
            "problems": 问题,
        })
    return mss.run(行们, dry_run=dry)


# ---- 03#箱 邮件发送（2026-11-05 用户口径）----
# 逻辑在 box03_mail.py；这里只做 HTTP 壳。
# 两次点击：第一次 confirm 缺省 False → 只返回「预览计划」（绝不调用 dws）；第二次 confirm=True → 真发。
@app.post("/api/box03/send-mail")
def box03_send_mail(request: Request, payload: dict | None = None):
    """body: {confirm?: bool}。

    返回口径：
      skipped=True            → 今天没有对应文件夹 / 没有 A2、A钳 开头 Excel → 不发；
      missing_addr 非空      → 有收件人地址没解析出来 → 禁止发送（要用户补地址）；
      need_confirm=True      → 预览通过，等用户二次确认；
      ok=True                 → 预览通过，或（confirm=True 时）已真的发送成功。
    """
    # 🔴🔴 自动化 token 只许预览、不许真发邮件（用户 2026-11-06 明令）
    confirm = bool((payload or {}).get("confirm")) and not _自动化请求(request)
    try:
        import box03_mail as b3
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"导入 box03_mail 失败：{type(e).__name__}: {e}")
    try:
        return b3.run(confirm=confirm)
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"发送失败：{type(e).__name__}: {e}")


# ---- MESE「下载保留」一键流水线：钉钉群附件下载 → 登记 → 邮件发送 → 汇总备份 ----
# 逻辑在 mese_pipeline.py；这里只做 HTTP 壳（页面「MESE下载保留」页签上的按钮调它）。
@app.post("/api/mese/pipeline")
def mese_pipeline_run(request: Request, payload: dict):
    """body: {kind: 'safety'|'limiter', group?: 群名, to?: 收件人, cc?: 抄送, limit?: 条数,
              force?: 同一节点未再更新也强制重跑, dry_run?: 只预演不发邮件/不写库}

    限速器线不传 to/cc 时＝截图口径（贾安琪、张三；抄送寿辉）；安全钳线保持原默认收件人。
    """
    kind = str(payload.get("kind") or "").strip().lower()
    if kind not in ("safety", "limiter"):
        raise HTTPException(status_code=400, detail="kind 只能是 safety 或 limiter")
    try:
        import mese_pipeline as mp
        # 🔴🔴 2026-11-06 用户：「MESE限速器以及MESE安全钳 我不需要你发邮件给我」
        #   → 两条线的收件人/抄送里一律剔除演示用户本人邮箱，绝不把邮件发回给自己。
        #   🔴 发给车间同事的业务邮件（MAIL_TO_* / MAIL_CC_*）原样保留，一个字不改
        #      （用户 11-06 明令「发邮件的代码不要改」）。
        report = mp.run(kind=kind,
                        group=str(payload.get("group") or ""),
                        收件人=_mese_去本人邮箱(payload.get("to")),
                        抄送=_mese_去本人邮箱(payload.get("cc")),
                        limit=int(payload.get("limit") or 100),
                        force=bool(payload.get("force")),
                        # 🔴🔴 自动化 token 强制预演（用户 2026-11-06：测试不许真发邮件）
                        dry_run=bool(payload.get("dry_run")) or _自动化请求(request))
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"流水线执行失败：{type(e).__name__}: {e}")
    # 🔴 2026-11-05 用户明令：发前校验（假想合同号 ↔ 安全钳物料号 必须成对）不通过 → 不发送，
    #   统一用 409 + 问题行清单回给前端弹窗（与 send-file 同一套形状，前端一个分支处理）。
    if report.get("kind") == "validate":
        raise HTTPException(status_code=409, detail={
            "kind": "validate",
            "summary": report.get("summary") or {},
            "problems": report.get("problems") or [],
        })
    return report


# ---- MESE 汇总页「手工备份到 MySQL」----
# 2026-11-03 建按钮；2026-11-05 用户改口径（🔴 明令）：
#   ① 点一下就自动全量替换，不再弹确认框（confirm 已在 index.html 两处去掉）；
#   ② 「MESE限速器」侧栏的按钮只替换 sap_table.MESE_115，「MESE安全钳」侧栏只替换 sap_table.MESE_125，
#      **一个按钮绝不把两张表都换掉**；03#箱databoard 的页签名表不再由本按钮维护；
#   ③ 与每日触发器 mese_summary_mysql_sync 同口径：行号升序做主键、全字段（以 store 键顺序为准）、
#      空串不转 NULL、先灌 __tmp 并逐格回读比对通过后才替换旧表；护栏防「汇总被洗成几十行」。
@app.post("/api/mese/mysql-backup")
def mese_mysql_backup(payload: dict):
    """body: {kind: 'safety'|'limiter'} → 只全量替换 sap_table 里该条线对应的那一张表。"""
    kind = str(payload.get("kind") or "").strip().lower()
    if kind not in ("safety", "limiter"):
        raise HTTPException(status_code=400, detail="kind 只能是 safety 或 limiter")
    if kind == "safety":
        store_path, sheet, table, label = MESE_SAFETY_STORE_PATH, MESE_SAFETY_SHEET, "MESE_125", "MESE安全钳汇总"
    else:
        store_path, sheet, table, label = MESE_LIMITER_STORE_PATH, MESE_LIMITER_SHEET, "MESE_115", "MESE限速器汇总"

    import pymysql
    try:
        # 定义在本文件下方（模块加载完才处理请求，运行时可解析）；返回 行号+全字段、按行号升序
        cols, vals = _mese_summary_mysql_read(store_path, sheet, table)
    except Exception as e:                                        # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"「{sheet}」数据校验未通过：{e}")

    tmp = table + "__tmp"
    conn = None
    try:
        conn = pymysql.connect(**MESE_SUMMARY_MYSQL_DB)
        with conn.cursor() as cur:
            # 护栏：库里已有 ≥100 行而本次不足一半 → 拒绝覆盖（页面停在旧快照 / 汇总被洗）
            cur.execute("SELECT COUNT(*) FROM information_schema.tables "
                        "WHERE table_schema='sap_table' AND LOWER(table_name)=LOWER(%s)", (table,))
            old_n = 0
            if cur.fetchone()[0]:
                cur.execute("SELECT COUNT(*) FROM `%s`" % table)
                old_n = int(cur.fetchone()[0])
            if old_n >= 100 and len(vals) < old_n / 2:
                raise RuntimeError(f"护栏拦截：{table} 现有 {old_n} 行，本次只读到 {len(vals)} 行，拒绝覆盖")

            maxlens = [max(len(str(v[i])) for v in vals) for i in range(len(cols))]
            ddl = ["`行号` INT NOT NULL"]
            for i, c in enumerate(cols[1:], start=1):
                ml = maxlens[i]
                ddl.append("`%s` %s" % (c, "TEXT" if ml > 255 else "VARCHAR(%d)" % max(32, min(255, ml + 32))))
            ddl.append("PRIMARY KEY (`行号`)")
            cur.execute("DROP TABLE IF EXISTS `%s`" % tmp)
            cur.execute("CREATE TABLE `%s` (%s) DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci "
                        "COMMENT='%s（行号顺序全字段，手工按钮自动覆盖）'" % (tmp, ", ".join(ddl), label))
            sql = "INSERT INTO `%s` (%s) VALUES (%s)" % (
                tmp, ", ".join("`%s`" % c for c in cols), ", ".join(["%s"] * len(cols)))
            for i in range(0, len(vals), 500):
                cur.executemany(sql, vals[i:i + 500])
            cur.execute("SELECT %s FROM `%s` ORDER BY `行号`" % (
                ", ".join("`%s`" % c for c in cols), tmp))
            got = [list(r) for r in cur.fetchall()]
            diff = sum(1 for a, b in zip(got, vals) if a != b)
            if len(got) != len(vals) or diff:
                raise RuntimeError(f"临时表回读不一致（{len(got)}/{len(vals)} 行，差异 {diff} 行），拒绝替换旧表")
            cur.execute("DROP TABLE IF EXISTS `%s`" % table)
            cur.execute("RENAME TABLE `%s` TO `%s`" % (tmp, table))
            cur.execute("SELECT COUNT(*) FROM `%s`" % table)
            n = int(cur.fetchone()[0])
        conn.commit()
        UI信息(f"MESE手工备份到 MySQL：{table}（{n} 行，{kind} 单线覆盖）")
        return {"ok": True, "table": table, "rows": n, "columns": len(cols) - 1, "kind": kind,
                "db": MESE_SUMMARY_MYSQL_DB.get("database"),
                "at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    except Exception as e:                                        # noqa: BLE001
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        raise HTTPException(status_code=500, detail=f"写 MySQL 失败：{type(e).__name__}: {e}（旧表保持原样）")
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


# ============ 静态托管 ============

@app.get("/")
def index(request: Request):
    """首页：把当前登录账号注入进页面（window.__authUser），前端据此过滤菜单/页面。

    2026-10-25 加登录门禁后，未登录的请求在中间件阶段就被 302 到 /login 了，
    能走到这里的都是已登录用户（或本机带 X-Auth-Token 的自动化脚本）。
    """
    路径 = os.path.join(APP_DIR, "index.html")
    注入 = {}
    try:
        import auth_gate as _ag
        注入 = _ag.me_for_request(request) or {}
    except Exception:
        pass
    try:
        html = Path(路径).read_text(encoding="utf-8")
        boot = "<script>window.__authUser=" + json.dumps(注入, ensure_ascii=False) + ";</script>"
        html = html.replace("<!--AUTH_BOOT-->", boot, 1)
        return Response(content=html, media_type="text/html",
                        headers={"Cache-Control": "no-cache, no-store, must-revalidate"})
    except Exception:
        return FileResponse(路径, headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


# ============ RAG 知识库子应用（合并到本站 8000 端口，2026-10-21）============
# 必须在 /{name}.js 等通配路由之前注册，否则 /rag/assets/xxx.js 会被通配吃掉。
try:
    from rag_mount import mount_rag as _mount_rag
    _rag_ok = _mount_rag(app)
except Exception as _e:      # 挂载失败不影响主站任何原有功能
    print(f"[RAG] 挂载异常，已跳过：{type(_e).__name__}: {_e}")
    _rag_ok = False


@app.get("/cron_help")
def cron_help():
    """时间表达式参考文档页（定时触发器 · 高级频率「了解时间表达式」链接目标）。"""
    return FileResponse(os.path.join(APP_DIR, "cron_help.html"), headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


@app.get("/{name}.js")
def static_js(name: str):
    """托管 app 目录下的 JS 模板/脚本文件（如 auto_tasks.js）。"""
    路径 = (APP_DIR / f"{name}.js").resolve()
    if 路径.parent != APP_DIR.resolve() or not 路径.is_file():
        raise HTTPException(status_code=404, detail="not found")
    return FileResponse(str(路径), media_type="application/javascript", headers={"Cache-Control": "no-cache, no-store, must-revalidate"})


# ============ 登录页素材（2026-10-25：背景图 / 右侧插画，登录前就要能取） ============
LOGIN_ASSET_DIR = APP_DIR / "login_assets"
_LOGIN_ASSET_MIME = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp", ".svg": "image/svg+xml"}


@app.get("/login-assets/{name}")
def login_asset(name: str):
    根 = LOGIN_ASSET_DIR.resolve()
    目标 = (根 / name).resolve()
    if not str(目标).startswith(str(根)) or not 目标.is_file():
        raise HTTPException(status_code=404, detail="素材不存在")
    类型 = _LOGIN_ASSET_MIME.get(目标.suffix.lower())
    if not 类型:
        raise HTTPException(status_code=404, detail="不支持的素材类型")
    return FileResponse(str(目标), media_type=类型, headers={"Cache-Control": "public, max-age=604800"})


# ============ 站点图标托管（2026-10-24，网站 favicon / PWA 图标） ============
# 白名单式单段路由：只认这几个文件名，其余一律 404，因此不会影响任何既有路径
# （全站没有任何单段路由，注册顺序也无冲突）。浏览器标签页、收藏夹、手机主屏幕、
# 桌面安装都走这里；图标文件放在 app/ 根下，与 index.html 同级。

_SITE_ICON_MIME = {
    "favicon.ico": "image/x-icon",
    "favicon.svg": "image/svg+xml",
    "apple-touch-icon.png": "image/png",
    "icon-192.png": "image/png",
    "icon-512.png": "image/png",
    "manifest.webmanifest": "application/manifest+json",
}


@app.get("/{icon_name}")
def site_icon(icon_name: str):
    if icon_name not in _SITE_ICON_MIME:
        raise HTTPException(status_code=404, detail="not found")
    路径 = (APP_DIR / icon_name).resolve()
    if 路径.parent != APP_DIR.resolve() or not 路径.is_file():
        raise HTTPException(status_code=404, detail="not found")
    # 图标长期不变，给一天缓存；index.html 本身仍是 no-store，换图标即时生效
    return FileResponse(str(路径), media_type=_SITE_ICON_MIME[icon_name],
                        headers={"Cache-Control": "public, max-age=86400"})


# ============ 媒体文件托管（2026-10-22，帮助中心视频/图片/文档） ============

@app.get("/media/{fname:path}")
def media_file(fname: str):
    """托管 app/media 目录下的媒体文件（含子目录，如 lib/three.min.js——三维演示页依赖）。
    FileResponse 原生支持 Range 请求（视频可拖动进度条）。"""
    媒体目录 = (APP_DIR / "media").resolve()
    路径 = (媒体目录 / fname).resolve()
    # 防目录穿越：解析后的真实路径必须位于 media 目录内（含子目录）
    if 媒体目录 not in 路径.parents or not 路径.is_file():
        raise HTTPException(status_code=404, detail="not found")
    类型 = {
        ".mp4": "video/mp4", ".webm": "video/webm", ".mov": "video/quicktime", ".m4v": "video/x-m4v",
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml",
        ".pdf": "application/pdf",
        ".html": "text/html", ".htm": "text/html",
        ".js": "text/javascript", ".css": "text/css", ".json": "application/json",
    }.get(路径.suffix.lower(), "application/octet-stream")
    return FileResponse(str(路径), media_type=类型, headers={"Cache-Control": "no-cache"})


# ============ 帮助中心 · 原文档只读（2026-10-29） ============
# 帮助中心「项目说明书」里的导读文章，文末都附「原文档」链接；点开走这里，
# 直接把 需求文档/ 下对应的 .md 原文以纯文本回给浏览器（浏览器原生显示，不下载）。
DOC_SRC_DIR = (APP_DIR.parent / "需求文档").resolve()

@app.get("/api/docs/raw")
def docs_raw(name: str = ""):
    """只读返回 需求文档/*.md 原文。白名单式：只认该目录下的 .md 文件，其余一律 404。"""
    from urllib.parse import unquote
    文件名 = unquote(name or "").strip().replace("\\", "/")
    if not 文件名 or "/" in 文件名 or not 文件名.lower().endswith(".md"):
        raise HTTPException(status_code=404, detail="not found")
    路径 = (DOC_SRC_DIR / 文件名).resolve()
    if 路径.parent != DOC_SRC_DIR or not 路径.is_file():
        raise HTTPException(status_code=404, detail="not found")
    return FileResponse(str(路径), media_type="text/plain; charset=utf-8",
                        headers={"Cache-Control": "no-cache"})


def _apd22_prewarm():
    """启动预热：后台建 enriched 缓存 + 默认筛选响应缓存，首个用户进3号箱APD及安全钳看板 即命中热路径。"""
    def _go():
        import time as _t
        _t.sleep(1.0)   # 等服务先把端口起来，避免拖慢启动
        try:
            _apd22_enriched()
            apd22_data(Apd22Filter())     # 默认响应缓存
            apd22_options()               # 筛选选项缓存
            apd22_detail("")              # 最新下达日明细缓存
            print("[apd22 prewarm] 3号箱APD及安全钳看板 缓存已预热（data/options/detail）")
        except Exception as e:
            print("[apd22 prewarm] 预热失败：", e)
    threading.Thread(target=_go, daemon=True).start()


_apd22_prewarm()


# ============ 订单下达记录 → 03_apd 同步 + 看板自动刷新（2026-10-23 用户明令改造） ============
# 需求（用户 2026-10-23 原话）：「看板根据我 mysql 数据库数据变化自动刷新，不需要看板能触发
#   数据入库这种神奇操作」——数据入库只能人工触发，绑定「立即同步」按钮（手工同步数据并更新看板）。
#
# 机制（改造后）：
#   ① 后台守护线程只做「MySQL 03_apd 行数巡检」：行数变化（立即同步/定时任务/外部写入所致）
#      → 清看板缓存 → 重建 03#箱databoard 落地表 → version 自增 → 前端 30s 轮询感知后自动重载看板。
#      **守护线程绝不读取源目录、绝不自动入库**。
#   ② 数据入库唯一人工入口：看板「立即同步」按钮（POST /api/apd22/auto-sync/run → _apd_sync_run，
#      45 天窗口 + 只取正式下达文件 a26-xxxx.xlsx，跳过中间态 A钳完-日期日期.xlsx 口径）。
#   ③ 定时触发器（apd_source_sync_5d，近 5 个工作日窗口）是另一条人工配置的入库路径，保留。
#   ④ 中间态口径说明：中间态内容次日会以正式文件出现，若先入库会把推算补齐的日期固化成事实数据
#      （历史库中所有 A钳完 文件夹均为 0 行，即该口径的证据）。

_APD_SYNC = {
    "version": 0,            # 数据版本号：库数据变化并刷新后自增，供前端判断是否重载
    "last_check": "",        # 最近一次库检查时间
    "last_sync": "",         # 最近一次手工同步完成时间
    "last_inserted": 0,      # 最近一次手工同步新增行数
    "total_rows": 0,         # 最近一次同步后 03_apd 总行数
    "db_rows": None,         # 03_apd 行数巡检基线：与实测值不一致即触发看板刷新
    "running": False,        # 手工同步进行中
    "error": "",             # 最近一次错误
    "log": [],               # 最近 20 条运行日志
    "interval_sec": 30,      # 兜底巡检间隔（秒）；主要检测由前端 3s 轮询 status 接口内联完成
    "window_days": 45,       # 手工同步窗口：回溯多少天扫描正式文件
    "auto": True,            # 看板自动刷新开关（仅刷新看板，不自动入库）
}
_APD_SYNC_LOCK = threading.Lock()


def _apd_sync_log(msg: str):
    line = f"[{datetime.now().strftime('%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    _APD_SYNC["log"].append(line)
    if len(_APD_SYNC["log"]) > 20:
        del _APD_SYNC["log"][:-20]


def _apd_sync_invalidate():
    """源数据入库后，让看板缓存失效（下次请求即重算）。"""
    with _APD_LOCK:
        _APD_CACHE.update(df=None, ts=0.0)
    _APD22_ENRICH.update(df=None, ts=0.0)
    _APD22_OPTS_CACHE.update(ver=-1, body=None)
    _APD22_DATA_CACHE.update(ver=-1, body=None)


def _apd_db_count() -> int:
    """03_apd 当前行数（巡检基线用，只读、极轻量）。"""
    import pymysql
    import apd_source_sync as _S
    conn = pymysql.connect(**_S.DB)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM `03_apd`")
            return int(cur.fetchone()[0])
    finally:
        conn.close()


_APD_REBUILD_LOCK = threading.Lock()


def _apd_rebuild_async():
    """后台异步重建 03#箱databoard 落地表（不阻塞 status 响应；看板实时数据不依赖落地表）。"""
    def _go():
        with _APD_REBUILD_LOCK:
            try:
                counts = _apd22_rebuild_tables()
                _apd_sync_log("03#箱databoard 落地表已重建：" +
                              " · ".join(f"{k} {v}" for k, v in (counts or {}).items()))
            except Exception as e:
                _apd_sync_log(f"落地表重建失败（不影响看板实时数据）：{e}")
    threading.Thread(target=_go, daemon=True).start()


def _apd_check_db_change(source: str = "巡检") -> bool:
    """检测 MySQL 03_apd 行数是否变化（毫秒级 COUNT）。有变化 → 看板缓存失效 + version 自增
    + 异步重建落地表；返回是否检测到变化。status 轮询接口与守护线程共用此检测。"""
    try:
        cnt = _apd_db_count()
    except Exception as e:
        _APD_SYNC["error"] = str(e)
        return False
    _APD_SYNC["last_check"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _APD_SYNC_LOCK:
        base = _APD_SYNC["db_rows"]
        if base is None:
            _APD_SYNC["db_rows"] = cnt
            return False
        if cnt == base:
            return False
        old = base
        _APD_SYNC["db_rows"] = cnt
        _APD_SYNC["total_rows"] = cnt
        _APD_SYNC["version"] += 1
        ver = _APD_SYNC["version"]
    _apd_sync_log(f"检测到 03_apd 数据变化（{old} → {cnt} 行，{source}），刷新看板 v{ver}")
    _apd_sync_invalidate()
    _apd_rebuild_async()
    return True


def _apd_sync_run(trigger: str = "定时") -> dict:
    """【唯一数据入库入口】执行一次源→03_apd 增量同步 + 缓存失效 + 落地表重建。返回本次结果摘要。
    由看板「立即同步」按钮（手工同步数据并更新看板）或排障接口调用；后台守护线程不会调用本函数。"""
    with _APD_SYNC_LOCK:
        if _APD_SYNC["running"]:
            return {"ok": False, "msg": "上一轮同步尚未结束"}
        _APD_SYNC["running"] = True
    try:
        if str(APP_DIR) not in sys.path:
            sys.path.insert(0, str(APP_DIR))
        import apd_source_sync as S

        root = S.SOURCE_ROOT
        if not os.path.isdir(root):
            raise RuntimeError(f"源目录不可达：{root}")
        today = datetime.now().date()
        dfrom = (today - dt.timedelta(days=_APD_SYNC["window_days"])).strftime("%Y.%m.%d")
        dto = (today + dt.timedelta(days=1)).strftime("%Y.%m.%d")
        _apd_sync_log(f"手工同步开始（{trigger}）：扫描 {dfrom} ~ {dto} 正式文件")
        stat = S.sync(root, dfrom=dfrom, dto=dto, dry_run=False,
                      verbose=False, formal_only=True)
        inserted = int(stat.get("inserted", 0))
        _apd_sync_log(f"源文件 {stat.get('files', 0)} 个 / 唯一键 {stat.get('unique', 0)} "
                      f"/ 跳过已存在 {stat.get('skipped_exists', 0)} / 新增 {inserted}")
        if inserted:
            _apd_sync_invalidate()
            try:
                counts = _apd22_rebuild_tables()
                _apd_sync_log("03#箱databoard 落地表已重建："
                              + " · ".join(f"{k} {v}" for k, v in (counts or {}).items()))
            except Exception as e:
                _apd_sync_log(f"落地表重建失败（不影响看板实时数据）：{e}")
            _APD_SYNC["version"] += 1
        _APD_SYNC["last_sync"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _APD_SYNC["last_inserted"] = inserted
        _APD_SYNC["total_rows"] = int(stat.get("total_after", 0) or 0)
        _APD_SYNC["db_rows"] = _APD_SYNC["total_rows"]   # 巡检基线同步更新，避免重复刷新看板
        _APD_SYNC["error"] = ""
        return {"ok": True, "inserted": inserted, "version": _APD_SYNC["version"],
                "total_rows": _APD_SYNC["total_rows"], "files": stat.get("files", 0)}
    except Exception as e:
        _APD_SYNC["error"] = str(e)
        _apd_sync_log(f"同步失败：{e}")
        return {"ok": False, "error": str(e)}
    finally:
        _APD_SYNC["running"] = False
        _APD_SYNC["last_check"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _apd_sync_loop():
    """守护线程：兜底巡检 MySQL 03_apd 行数变化 → 清缓存、version 自增，看板自动刷新。
    看板页打开时，主要检测由前端 3s 轮询 /api/apd22/auto-sync/status 内联完成（延迟≈3s）；
    本线程负责看板无人打开时的兜底（用户 2026-10-23 明令：只感知库变化并刷新看板，
    绝不读取源目录、绝不自动入库——看板自动触发数据入库属于「神奇操作」，已废除）。"""
    def _go():
        import time as _t
        _t.sleep(20)   # 等服务与预热完成后再开始，避免与预热抢 MySQL
        try:
            _APD_SYNC["db_rows"] = _apd_db_count()
            _apd_sync_log(f"看板自动刷新兜底巡检已启动：每 {_APD_SYNC['interval_sec']}s 检查一次 "
                          f"03_apd 数据变化（当前 {_APD_SYNC['db_rows']} 行；数据入库请用「立即同步」按钮，"
                          f"本进程不会自动入库）")
        except Exception as e:
            _apd_sync_log(f"03_apd 行数基线初始化失败（稍后重试）：{e}")
        while True:
            _t.sleep(_APD_SYNC["interval_sec"])
            if not _APD_SYNC.get("auto", True):
                continue
            _apd_check_db_change("兜底巡检")
    threading.Thread(target=_go, daemon=True).start()


@app.get("/api/apd22/auto-sync/status")
def apd22_auto_sync_status():
    """看板自动刷新状态：数据版本号 / 最近同步时间 / 最近新增行数 / 库行数巡检基线。
    前端 3s 轮询此接口；每次轮询顺带做一次库变化检测（毫秒级 COUNT），
    有变化当次响应即带新 version，前端立即重载——端到端延迟≈前端轮询间隔。
    本接口与巡检线程均不触发数据入库。"""
    _apd_check_db_change("轮询检测")
    with _APD_SYNC_LOCK:
        s = dict(_APD_SYNC)
    s["log"] = list(s.get("log", []))
    return {"ok": True, **s}


@app.post("/api/apd22/auto-sync/run")
def apd22_auto_sync_run():
    """「立即同步」按钮后端：手工同步数据并更新看板（唯一数据入库入口）。"""
    r = _apd_sync_run("手动触发")
    if not r.get("ok"):
        raise HTTPException(status_code=500, detail=r.get("error") or r.get("msg") or "同步失败")
    return r


threading.Thread(target=_apd_sync_loop, daemon=True).start()


# ============ 触发器计划任务 · 后端应用注册表 + 到点调度（2026-10-21） ============
# 需求：把「抓取最近 5 个工作日 A 开头正式文件 → 去重 + 补齐日期 → 入库 03_apd」
#       做成个人工作网站「触发器 → 计划任务」里的一条定时任务（每天 07:30）。
# 现状：触发器的「条件」列只在浏览器里推算"最近几次触发时间"（纯预览），
#       真正到点执行需要后端调度器 —— 本区块即为此而加。
# 机制：
#   ① _TRIGGER_APP_REGISTRY 注册可被后端真实执行的应用（目前 1 个：apd_source_sync_5d）
#   ② _trigger_scheduler_loop 每 30s 读一次触发器存储「计划任务」sheet，
#      命中的启用定时任务到点即执行，并把结果写进「运行日志」sheet（与手动运行同结构）
#   ③ POST /api/trigger/app/run 供网页「立即运行」对这类应用走真实后端执行（不再模拟）
_TRIGGER_APP_REGISTRY = {
    "apd_source_sync_5d": {
        "名称": "订单下达记录同步（近5个工作日→03_apd）",
        "描述": "抓取订单下达记录目录下最近 5 个工作日的 A 开头正式下达文件，"
                "按「合同箱号&安全钳组件」去重、按工厂日历补齐两列日期后追加写入 sap_table.03_apd，"
                "并自动刷新 03#箱APD及安全钳看板",
    },
    # 2026-10-26 新增：维护物料模块自动化（原 box_level_task.py）接入定时调度，
    # 与网页「维护物料模块自动化」页的「运行」按钮同一条链路（box_task_api 子进程 + 单实例栅栏）。
    "box_level_task": {
        "名称": "维护物料模块自动化（箱子及装箱级物料未维护）",
        "描述": "打开 ECS 物料维护流程页 → 清空操作人员筛选 → 查询 → 提取并筛选"
                "箱子及装箱级物料未维护的数据 → 生成 Excel → 按抬头三通道（钉钉消息/群消息/邮箱）通知",
    },
    # 2026-10-26 新增：看板定时推送（模块 app/board_push.py）    # 整页长截图（当日下达箱数 / 近7天完工日波动 / 小型化每日下达 三张 KPI 卡的悬浮面板展开）
    # + 看板 Excel 附件，走 dws_mail 邮件；收件人与开关在 app/board_push_store/config.json。
    "apd22_board_push": {
        "名称": "3号箱APD及安全钳看板 · 定时推送",
        "描述": "打开「看板展示 → 3号箱APD及安全钳看板」整页长截图（含三张 KPI 悬浮面板），"
                "同时导出看板 Excel，邮件推送（内联截图 + Excel 附件）",
    },
    # 2026-10-26 新增：条件触发器的执行体（类型='条件' 的任务用这个应用）
    # 指标 >/< 阈值 成立才发一封告警邮件（该指标那张 KPI 卡截图 + 明细 Excel），不成立只写「未达阈值」。
    "apd22_metric_alert": {
        "名称": "看板指标阈值提醒（条件触发器用）",
        "描述": "取看板指标（小型化每日下达 / 当日下达箱数 / 近7天完工合计），按「>」或「<」与阈值比较，"
                "命中才单独发一封告警邮件（KPI 卡截图 + 明细 Excel）",
    },
    # 2026-10-29 新增：SAPGUI 自动化脚本 · Co28（订单信息查询/批量下达）接入定时调度。
    # 与网页「SAPGUI 自动化脚本 → Co28」页点「执行查询」同一条链路（同一个 /api/vbs/run 执行函数），
    # 查询条件用页面默认口径，两个日期按当天滚动（今天-40 天 / 今天+100 天）。
    "co28_sap_query": {
        "名称": "SAPGUI · Co28 订单信息查询（执行查询 + 结果截图邮件）",
        "描述": "按页面默认条件（订单类别 2 / 参数文件 000001 / 物料 * / 工厂 0001 / MRP 060 / "
                "生产调度员 24* / 基本开始日期 今天-40 天 ~ 今天+100 天）执行 CO28 查询，"
                "跑完把结果截图发邮件（收件人默认 demo.user72@example.com）",
    },
    # 2026-10-29 新增：钉钉邮箱巡检 —— 「小于阀值异常工单明细」邮件（每天约 02:05 到达）
    # 把最近 5 封的 Excel 附件都读一遍，D 列（实际条目数）为 0 的行按 A 列生产订单号去重后发钉钉消息。
    # ★ 排在 07:20 而不是 02:00：凌晨内网无外网出口，dws 读邮件/发消息都会失败（邮件在云端不丢）。
    # 2026-11-04 新增：安全钳+限速器合并同步（用户需求：两步合一次调取）
    # 图号=两表汇总且前后加 * 通配（P115007A000 → *P115007A000*）；部件名称两套包含合并；
    # 品目=115,125；查回按「品目」分流：115→限速器物料号（JSON store）、125→安全钳物料号（SQLite 大表）
    "materials_sync_daily": {
        "名称": "安全钳+限速器物料号每日同步（合并查询→按品目分流）",
        "描述": "一次跑 /nzb02（ZLSTMA）：图号=安全钳总图号分类+限速器参数全部图号（前后加 * 通配），"
                "部件名称 *安全钳*,*限速器* 排除 通用件/标签/座/构/机构/铭牌/提拉，品目 115,125，品/群 G*，L 番 L*，"
                "创建日期 今天-7~今天 → 导出清单 → 按「品目」分流去重：115 进限速器物料号、125 进安全钳物料号，"
                "各自按创建日期倒序插到最前",
    },
    # 2026-11-03 新增：安全钳物料号每日同步（ZLSTMA 清单 → 去重 → 倒序插入「安全钳物料号」SQLite 大表）
    "safclamp_material_sync": {
        "名称": "安全钳物料号每日同步（ZLSTMA→去重→倒序插入）",
        "描述": "跑 SafClamp_Sync.vbs（/nzb02 数据浏览器 ZLSTMA：部件名称 *安全钳* 排除通用件/标签/座/构/铭牌/提拉、"
                "品目 125、图号/变量=安全钳总图号分类全部图号、品/群 G*、L 番 L*、创建日期 今天-7~今天）→ "
                "导出清单 → 解析 → 按「物料」列与「安全钳物料号」大表去重 → 新物料号整行按 创建日期 倒序插到最前",
    },
    # 2026-11-03 新增：限速器物料号每日同步（ZLSTMA 清单 → 去重 → 倒序插入「限速器物料号」基础数据）
    "limiter_material_sync": {
        "名称": "限速器物料号每日同步（ZLSTMA→去重→倒序插入）",
        "描述": "跑 Zlstma_Sync.vbs（/nzb02 数据浏览器 ZLSTMA：部件名称 *限速器* 排除通用件/标签/座/机构/铭牌、"
                "品目 115、图号/变量=限速器参数表全部图号、品/群 G*、L 番 L*、创建日期 今天-7~今天）→ "
                "导出清单 → 解析 → 按「物料」列与「限速器物料号」基础数据去重 → 新物料号整行按 创建日期 倒序插到最前",
    },
    "yuzhi_mail_check": {
        "名称": "钉钉邮箱巡检 · 小于阀值异常工单明细",
        "描述": "读收件箱最近 5 封「小于阀值异常工单明细」邮件（发件人 U2632@systemmail.example.com），"
                "下载 Excel 附件，D 列「实际条目数」为 0（0/00/000/0.0 都算）的行按 A 列生产订单号去重后，"
                "整行发钉钉消息（默认发给当前登录用户本人）",
    },
    # 2026-11-05 用户：每天把两个「汇总」页签整表覆盖进 sap_table 的对应表（MESE_125 / MESE_115）。
    # 只覆盖这两张表，不碰「手工备份到 MySQL」写到 03#箱databoard 的同名页签表（页/任务互不扩散）。
    "mese_summary_mysql_sync": {
        "名称": "MESE安全钳汇总 / MESE限速器汇总 每日覆盖 MySQL",
        "描述": "读「MESE安全钳汇总」「MESE限速器汇总」两个页签的本地存储（行号 + 网站全部字段，按行号升序），"
                "全量重建 sap_table.MESE_125 / sap_table.MESE_115：先建临时表灌数并逐格回读比对，"
                "全部通过才替换旧表；汇总行数异常偏少时护栏拦截，旧表一张都不动",
    },
}
_TRIGGER_SCHED_STATE_PATH = APP_DIR / "trigger_store" / "sched_state.json"
_TRIGGER_SCHED_LOCK = threading.Lock()
_TRIGGER_STORE_LOCK = threading.Lock()   # 调度器写「运行日志」与前端写入互斥
_TRIGGER_SCHED_STATE = {"fired": {}}     # {任务id: {"last_day": "yyyy-mm-dd", "task": 名称, "last_ts": ...}}
_TRIGGER_WEEK_CN = {"周一": 0, "周二": 1, "周三": 2, "周四": 3, "周五": 4, "周六": 5, "周日": 6}


def _trig_sched_state_load() -> dict:
    global _TRIGGER_SCHED_STATE
    try:
        if _TRIGGER_SCHED_STATE_PATH.exists():
            d = json.loads(_TRIGGER_SCHED_STATE_PATH.read_text(encoding="utf-8"))
            if isinstance(d, dict) and isinstance(d.get("fired"), dict):
                _TRIGGER_SCHED_STATE["fired"] = d["fired"]
    except Exception:
        pass
    return _TRIGGER_SCHED_STATE


def _trig_sched_state_save():
    try:
        _TRIGGER_SCHED_STATE_PATH.write_text(
            json.dumps({"fired": _TRIGGER_SCHED_STATE["fired"]}, ensure_ascii=False),
            encoding="utf-8")
    except Exception as e:
        _apd_sync_log(f"调度状态落盘失败（不影响本次执行）：{e}")


def _trig_due_datetime(cfg: dict, now: dt.datetime):
    """按 config.freq 计算「今天是否到点」；到点则返回今天的触发时刻，否则 None。

    支持：每天 / 每工作日（工厂日历）/ 每周 / 月份 / 高级（cron 五段式，需 config.cron）。
    「高级」若只有中文 cronDesc 而无可执行表达式，则无法调度（返回 None）。"""
    freq = str((cfg or {}).get("freq") or "")
    hhmm = str((cfg or {}).get("time") or "00:00")
    try:
        h, m = (int(x) for x in hhmm.split(":")[:2])
    except Exception:
        return None
    base = now.replace(hour=h, minute=m, second=0, microsecond=0)
    today = now.date()

    def _is_factory_workday(d) -> bool:
        try:
            import apd_source_sync as _S
            return _S.is_workday(d)
        except Exception:
            return d.weekday() < 5

    if freq == "每天":
        return base if now >= base else None
    if freq == "每工作日":
        mode = (cfg or {}).get("calendarMode") or "factory"
        ok_day = (today.weekday() < 5) if mode == "cn" else _is_factory_workday(today)
        return (base if now >= base else None) if ok_day else None
    if freq == "每周":
        wds = [str(x) for x in ((cfg or {}).get("weekdays") or [])]
        if any(_TRIGGER_WEEK_CN.get(w) == today.weekday() for w in wds):
            return base if now >= base else None
        return None
    if freq == "月份":
        try:
            md = int((cfg or {}).get("monthday") or 1)
        except Exception:
            md = 1
        if today.day != md:
            return None
        return base if now >= base else None
    if freq == "高级":
        return _cron_due_today(str((cfg or {}).get("cron") or ""), now)
    return None


_CRON_MON_ALIAS = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
                   "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
_CRON_DOW_ALIAS = {"SUN": 0, "MON": 1, "TUE": 2, "WED": 3, "THU": 4, "FRI": 5, "SAT": 6}


def _cron_field(field: str, min_v: int, max_v: int, alias: dict = None) -> set:
    """解析 cron 单段：* / a / a-b / a-b/n / */n / 逗号列表；周字段 0 与 7 均=周日。"""
    out = set()
    for part in str(field).split(","):
        part = part.strip()
        if not part:
            continue
        step = 1
        if "/" in part:
            part, s = part.split("/", 1)
            step = max(1, int(s) or 1)
        if alias:
            for k, v in alias.items():
                part = part.upper().replace(k, str(v))
        if part == "*":
            lo, hi = min_v, max_v
        elif "-" in part:
            a, b = part.split("-", 1)
            lo, hi = int(a), int(b)
        else:
            lo = hi = int(part)
        for v in range(lo, hi + 1, step):
            out.add(v % 7 if max_v == 7 else v)
    return out


def _cron_due_today(expr: str, now: dt.datetime):
    """cron 五段式（分 时 日 月 周）：今天是否已到点 → 到点则返回今天该时刻，否则 None。"""
    f = str(expr or "").split()
    if len(f) != 5:
        return None
    try:
        minute = _cron_field(f[0], 0, 59)
        hour = _cron_field(f[1], 0, 23)
        dom = _cron_field(f[2], 1, 31)
        mon = _cron_field(f[3], 1, 12, _CRON_MON_ALIAS)
        dow = _cron_field(f[4], 0, 7, _CRON_DOW_ALIAS)
    except Exception:
        return None
    if now.month not in mon:
        return None
    # cron 惯例：日与周均为非 * 时取并集（任一命中即算当天有效）
    dom_any, dow_any = f[2] == "*", f[4] == "*"
    if not dom_any and now.day not in dom:
        return None
    if not dow_any:
        wd = {6 if d in (0, 7) else d - 1 for d in dow}   # cron 周日=0/7 → Python 周日=6
        if now.weekday() not in wd:
            return None
    for h in sorted(hour):
        for mi in sorted(minute):
            due = now.replace(hour=h, minute=mi, second=0, microsecond=0)
            if now >= due:
                return due
    return None


def _sap任务报表邮件(cfg: dict) -> bool:
    """计划任务 config「报表邮件」：默认 False（只留汇总通知一封），显式打开才恢复脚本侧邮件。"""
    v = (cfg or {}).get("报表邮件", False)
    return str(v).strip().lower() not in ("", "0", "false", "no", "none", "否", "关", "关闭")


def _sap前置回车次数(cfg: dict) -> int:
    """SAP 相关计划任务：主脚本执行前先按几次 Enter（2026-11-05 用户明令「执行 2 次」）。

    任务 config 里可用「前置回车」覆盖次数（0=关闭）；非法值一律回落到 2。
    只被 4 个 SAP 触发器应用使用，其它计划任务完全不碰。
    """
    try:
        return max(0, min(10, int((cfg or {}).get("前置回车", 2))))
    except Exception:
        return 2


def _trig_run_app(app_id: str, cfg: dict = None) -> dict:
    """执行注册表内的后端应用，返回 {ok, 日志:[...], 结果:文本, 耗时s, 明细}。

    cfg = 该条计划任务的 config（条件触发器要读 指标/比较/阈值/收件人）。"""
    app = _TRIGGER_APP_REGISTRY.get(app_id)
    if not app:
        return {"ok": False, "结果": "该应用未注册后端执行器",
                "日志": [f"✗ 应用「{app_id}」没有后端执行器（仅支持网页端已注册任务）"]}
    lines: list = []
    t0 = datetime.now()
    try:
        if app_id == "apd_source_sync_5d":
            r = _trig_apd_sync_5d(lines, t0)
        elif app_id == "box_level_task":
            r = _trig_box_task(lines, t0)
        elif app_id == "apd22_board_push":
            r = _trig_board_push(lines, t0, cfg)
        elif app_id == "apd22_metric_alert":
            r = _trig_metric_alert(lines, t0, cfg)
        elif app_id == "co28_sap_query":
            r = _trig_co28(lines, t0, cfg)
        elif app_id == "yuzhi_mail_check":
            r = _trig_yuzhi_mail(lines, t0, cfg)
        elif app_id == "limiter_material_sync":
            r = _trig_limiter_material_sync(lines, t0, cfg)
        elif app_id == "safclamp_material_sync":
            r = _trig_safclamp_material_sync(lines, t0, cfg)
        elif app_id == "materials_sync_daily":
            r = _trig_materials_sync_daily(lines, t0, cfg)
        elif app_id == "mese_summary_mysql_sync":
            r = _trig_mese_summary_mysql_sync(lines, t0, cfg)
        else:  # pragma: no cover - 注册表扩展点
            r = {"ok": False, "结果": "未实现"}
        r["日志"] = lines
        r.setdefault("耗时s", round((datetime.now() - t0).total_seconds(), 1))
        return r
    except Exception as e:
        lines.append(f"✗ 运行失败：{e}")
        return {"ok": False, "结果": "运行失败", "日志": lines,
                "耗时s": round((datetime.now() - t0).total_seconds(), 1)}


# ---- 2026-11-05 新增：MESE 汇总每日全量覆盖 MySQL（sap_table.MESE_125 / MESE_115）----
# 口径（与用户 11-05 手工建表一致，沿用 /api/mese/mysql-backup 的「先建临时表再 rename」）：
#   ① 数据源＝两个「汇总」页签的本地 store 文件；字段以 data 的键顺序为准（网站加列自动跟随，
#      不会只抄截图里那几列），不触发 resync、不动页面，避开 11-05 踩过的「汇总被洗成几十行」坑；
#   ② 按 row_no 升序，行号做主键；空值保留空串不转 NULL；
#   ③ 护栏：库里已有 ≥100 行而本次读到不足其一半 → 整任务拒绝覆盖，两张表都不动；
#   ④ 两阶段：两张 __tmp 都建好、灌完、逐格回读比对通过，才统一 DROP 旧表 + RENAME，失败不毁旧表。
# 注意：只覆盖本任务指定的两张表，绝不扩散到 03#箱databoard 里「手工备份到 MySQL」的同名页签表。
MESE_SUMMARY_MYSQL_DB = dict(host="127.0.0.1", port=3306, user="u3854",
                             password="u38544253", database="sap_table", charset="utf8mb4")
MESE_SUMMARY_MYSQL_TARGETS = (
    (MESE_SAFETY_STORE_PATH, MESE_SAFETY_SHEET, "MESE_125", "MESE安全钳汇总"),
    (MESE_LIMITER_STORE_PATH, MESE_LIMITER_SHEET, "MESE_115", "MESE限速器汇总"),
)


def _mese_summary_mysql_read(store_path, sheet, table):
    """读一张汇总 sheet → (列名列表, 按行号排好的二维值)；带字段/行号完整性校验。"""
    store = _generic_store_load(store_path) or {}
    rows = ((store.get("sheets") or {}).get(sheet) or [])
    if not rows:
        raise RuntimeError(f"「{sheet}」本地汇总为 0 行，拒绝覆盖 {table}（可能正被其它进程重写）")
    rows = sorted(rows, key=lambda r: int(r.get("row_no") or 0))
    nos = [int(r.get("row_no") or 0) for r in rows]
    if len(set(nos)) != len(nos):
        raise RuntimeError(f"「{sheet}」行号有重复，拒绝覆盖 {table}")
    first = list((rows[0].get("data") or {}).keys())
    if not first:
        raise RuntimeError(f"「{sheet}」首行没有字段，拒绝覆盖 {table}")
    want = set(first)
    for r in rows:
        if set((r.get("data") or {}).keys()) != want:
            raise RuntimeError(f"「{sheet}」各行字段不一致（第 {r.get('row_no')} 行缺列），拒绝覆盖 {table}")
    vals = [[int(r["row_no"])] + [str((r.get("data") or {}).get(c, "") or "") for c in first]
            for r in rows]
    return ["行号"] + first, vals


def _trig_mese_summary_mysql_sync(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 mese_summary_mysql_sync：两张 MESE 汇总 → 全量覆盖 sap_table 对应表。"""
    import pymysql
    prepared, conn = [], None
    try:
        conn = pymysql.connect(**MESE_SUMMARY_MYSQL_DB)
        with conn.cursor() as cur:
            # ---- 阶段 0：只读预检 + 护栏（此刻一张表都不动）----
            for store_path, sheet, table, label in MESE_SUMMARY_MYSQL_TARGETS:
                cols, vals = _mese_summary_mysql_read(store_path, sheet, table)
                cur.execute("SELECT COUNT(*) FROM information_schema.tables "
                            "WHERE table_schema='sap_table' AND LOWER(table_name)=LOWER(%s)", (table,))
                exists = bool(cur.fetchone()[0])
                old_n = 0
                if exists:
                    cur.execute("SELECT COUNT(*) FROM `%s`" % table)
                    old_n = int(cur.fetchone()[0])
                if old_n >= 100 and len(vals) < old_n / 2:
                    raise RuntimeError(
                        f"护栏拦截：{table} 现有 {old_n} 行，而「{sheet}」只读到 {len(vals)} 行（不足一半），"
                        f"疑似汇总被洗过，拒绝覆盖（两张表都不动）")
                lines.append(f"✓ {label}：本地 {len(vals)} 行 × {len(cols) - 1} 字段"
                             f"（行号 {vals[0][0]}~{vals[-1][0]}）｜库里现有 {old_n} 行")
                prepared.append((table, cols, vals, label, old_n))
            # ---- 阶段 1：两张临时表全部建好、灌完、逐格回读比对 ----
            for table, cols, vals, label, old_n in prepared:
                tmp = table + "__tmp"
                maxlens = [max(len(str(v[i])) for v in vals) for i in range(len(cols))]
                ddl = ["`行号` INT NOT NULL"]
                for i, c in enumerate(cols[1:], start=1):
                    ml = maxlens[i]
                    ddl.append("`%s` %s" % (c, "TEXT" if ml > 255
                                            else "VARCHAR(%d)" % max(32, min(255, ml + 32))))
                ddl.append("PRIMARY KEY (`行号`)")
                cur.execute("DROP TABLE IF EXISTS `%s`" % tmp)
                cur.execute("CREATE TABLE `%s` (%s) DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci "
                            "COMMENT='%s（行号顺序全字段，每日自动覆盖）'" % (tmp, ", ".join(ddl), label))
                sql = "INSERT INTO `%s` (%s) VALUES (%s)" % (
                    tmp, ", ".join("`%s`" % c for c in cols), ", ".join(["%s"] * len(cols)))
                for i in range(0, len(vals), 500):
                    cur.executemany(sql, vals[i:i + 500])
                cur.execute("SELECT %s FROM `%s` ORDER BY `行号`" % (
                    ", ".join("`%s`" % c for c in cols), tmp))
                got = [list(r) for r in cur.fetchall()]
                diff = sum(1 for a, b in zip(got, vals) if a != b)
                if len(got) != len(vals) or diff:
                    raise RuntimeError(f"{table} 临时表回读不一致"
                                       f"（{len(got)}/{len(vals)} 行，差异 {diff} 行），拒绝替换旧表")
                lines.append(f"✓ {table} 临时表 {len(got)} 行灌完，逐格比对 0 差异")
            # ---- 阶段 2：全部通过后才统一替换（旧表此刻才动）----
            for table, cols, vals, label, old_n in prepared:
                tmp = table + "__tmp"
                cur.execute("DROP TABLE IF EXISTS `%s`" % table)
                cur.execute("RENAME TABLE `%s` TO `%s`" % (tmp, table))
        conn.commit()
        for table, cols, vals, label, old_n in prepared:
            lines.append(f"✓ 已覆盖 {table}：{old_n} 行 → {len(vals)} 行 × {len(cols)} 列")
        n_map = {t: len(v) for t, c, v, l, o in prepared}
        return {"ok": True, "结果": "MESE 汇总全量覆盖完成",
                "明细": n_map, "total_rows": sum(n_map.values())}
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        lines.append(f"✗ 覆盖失败：{type(e).__name__}: {e}（旧表保持原样）")
        return {"ok": False, "结果": "覆盖失败"}
    finally:
        if conn:
            try:
                conn.close()
            except Exception:
                pass


def _trig_apd_sync_5d(lines: list, t0) -> dict:
    """触发器应用 apd_source_sync_5d 的真实执行体：近 5 个工作日 → 03_apd → 看板刷新。"""
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    import apd_source_sync as S
    root = S.SOURCE_ROOT
    if not os.path.isdir(root):
        raise RuntimeError(f"源目录不可达：{root}")
    folders = S.recent_workday_folders(5)
    lines.append(f"抓取窗口：最近 5 个工作日（工厂日历）"
                 f"{folders[0]} ~ {folders[-1]}，共 {len(folders)} 个日期文件夹")
    lines.append("口径：只取正式下达文件 A26-xxxx.xlsx；跳过中间态 A钳完-日期日期.xlsx")
    stat = S.sync(root, folders=folders, dry_run=False, verbose=False, formal_only=True)
    inserted = int(stat.get("inserted", 0))
    lines.append(f"源正式文件 {stat.get('files', 0)} 个 / 源行 {stat.get('source_rows', 0)} / "
                 f"批内去重 {stat.get('batch_dup', 0)} / 唯一键 {stat.get('unique', 0)}")
    lines.append(f"跳过已存在 {stat.get('skipped_exists', 0)} / 新增 {inserted}"
                 f" / 日期补齐 {stat.get('filled_dates', 0)} / 仍为空 {stat.get('still_empty', 0)}")
    if inserted:
        _apd_sync_invalidate()
        try:
            counts = _apd22_rebuild_tables()
            lines.append("03#箱databoard 落地表已重建：" +
                         " · ".join(f"{k} {v}" for k, v in (counts or {}).items()))
        except Exception as e:
            lines.append(f"[警告] 落地表重建失败（不影响看板实时数据）：{e}")
        _APD_SYNC["version"] += 1
        _APD_SYNC["last_sync"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _APD_SYNC["last_inserted"] = inserted
    _APD_SYNC["total_rows"] = int(stat.get("total_after", 0) or 0)
    _APD_SYNC["db_rows"] = _APD_SYNC["total_rows"]   # 巡检基线同步更新，避免重复刷新看板
    _APD_SYNC["last_check"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _APD_SYNC["error"] = ""
    _apd_sync_log(f"触发器同步（近5个工作日）：文件 {stat.get('files', 0)} / 新增 {inserted} / "
                  f"总行 {_APD_SYNC['total_rows']}")
    lines.append(f"✓ 运行完成，耗时 {(datetime.now() - t0).total_seconds():.1f}s；"
                 f"03_apd 总行 {_APD_SYNC['total_rows']}")
    return {"ok": True, "结果": "运行成功", "inserted": inserted,
            "files": stat.get("files", 0), "total_rows": _APD_SYNC["total_rows"],
            "folders": folders}


def _trig_box_task(lines: list, t0, 超时秒: int = 2400) -> dict:
    """触发器应用 box_level_task 的真实执行体：起 box_task_api 子进程跑维护物料脚本，并等它跑完。

    刻意复用网页「维护物料模块自动化」页的运行链路（box_task_api.box_run/box_status）：
      · 同一把单实例栅栏 —— 到点触发与人工点「运行」不会同时开两个 Chrome 抢 9222；
      · 同一套日志窗口 —— 运行日志 sheet 里看到的行与页面底部日志一致；
      · 通知复用抬头三通道（跑完由 box_task_api 自己按 钉钉消息/群消息/邮箱 开关发出）。
    脚本自带 --no-push（只采集 + 出 Excel），故发送动作不会重复。
    """
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    try:
        import box_task_api as B
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 维护物料接口不可用：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "维护物料接口未挂载"}
    if not B.SCRIPT_PATH.is_file():
        lines.append(f"✗ 脚本不存在：{B.SCRIPT_PATH}")
        return {"ok": False, "结果": "脚本不存在"}

    try:
        启动 = B.box_run(B.BoxRunRequest())
    except HTTPException as e:
        if e.status_code == 409:                                  # 已有运行中的任务
            在跑 = (B.box_state() or {}).get("run_id") or ""
            lines.append(f"⏭ 已有维护物料任务在运行（run_id={在跑}），本轮不重复启动")
            lines.append("   它跑完会照常生成 Excel 并按抬头三通道发通知，本轮视为已覆盖")
            return {"ok": True, "结果": "跳过（已有运行中的任务）"}
        lines.append(f"✗ 启动失败：{e.detail}")
        return {"ok": False, "结果": "启动失败"}
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 启动失败：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "启动失败"}

    run_id = 启动.get("run_id") or ""
    lines.append(f"子进程已启动（run_id={run_id}，解释器 {启动.get('interpreter')}）")
    游标, 全部行, 退出码 = 0, [], None
    截止 = time.time() + max(60, int(超时秒))
    while True:
        try:
            状态 = B.box_status(run_id=run_id, after=游标)
        except HTTPException as e:
            lines.append(f"✗ 读取运行状态失败：{e.detail}")
            return {"ok": False, "结果": "读取运行状态失败"}
        except Exception as e:                                    # noqa: BLE001
            lines.append(f"✗ 读取运行状态失败：{type(e).__name__}: {e}")
            return {"ok": False, "结果": "读取运行状态失败"}
        全部行.extend(状态.get("lines") or [])
        游标 = 状态.get("next") or 游标
        退出码 = 状态.get("exit_code")
        if not 状态.get("running") and 退出码 is not None:
            break
        if time.time() > 截止:
            lines.append(f"✗ 运行超过 {max(60, int(超时秒)) // 60} 分钟仍未结束，已强制中断")
            try:
                B.box_stop(B.BoxStopRequest(run_id=run_id))
            except Exception:                                     # noqa: BLE001
                pass
            return {"ok": False, "结果": "运行超时被中断"}
        time.sleep(2)

    尾部 = 全部行[-400:]                    # 运行日志单元格只留尾部，别一条任务塞几千行
    if len(全部行) > len(尾部):
        lines.append(f"（脚本输出 {len(全部行)} 行，这里只留最后 {len(尾部)} 行；"
                     f"完整日志见 {B.LOG_FILE}）")
    lines.extend(尾部)

    摘要 = B.parse_summary()
    条数 = 摘要.get("筛选后")
    excel = 摘要.get("Excel") or ""
    ok = (退出码 == 0) and not any("[FAIL]" in x for x in 全部行[-200:])
    if ok:
        结果 = "运行成功" + (f"：筛选后 {条数} 条" if 条数 is not None else "")
    else:
        结果 = f"运行失败（退出码 {退出码}）"
        if 摘要.get("失败"):
            结果 += f"：{摘要['失败']}"
    _apd_sync_log(f"触发器维护物料：{结果}｜run_id={run_id}｜excel={excel or '无'}")
    产物 = ""
    try:
        _p = B.latest_excel()
        if _p is not None:
            产物 = str(_p)          # 通知里直接把 Excel 发给收件人（不只是文件名）
    except Exception:                                         # noqa: BLE001
        pass
    return {"ok": ok, "结果": 结果, "rows": 条数, "excel": excel, "xlsx": 产物, "run_id": run_id}


def _trig_co28(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 co28_sap_query：复刻网页「SAPGUI 自动化脚本 → Co28」页点「执行查询」那条链路。

    参数口径与页面默认值完全一致（app/index.html 的 CO28_DEFAULTS）：
      订单类别 2 / 参数文件 000001 / 物料 * / 工厂 0001 / MRP 控制器 060 / 生产调度员 24*
      基本开始日期 = 今天 -40 天、基本完成日期 = 今天 +100 天。
    **日期必须每次按当天重算**（页面 co28OffsetDate 就是这个口径）——写死日期过几天就查不到数据。
    config 里给同名字段即可覆盖任意一项（catLow / prof / matLow / matHigh / werksLow / … / 收件人 / mail）。
    直接调 /api/vbs/run 的执行函数，复用它的单实例栅栏、运行时 VBS 注入、多屏截图与结果邮件，
    不另起一套 SAP 交互逻辑（保证「手动跑」和「定时跑」结果一致）。
    """
    cfg = cfg or {}

    def _相对日期(天: int) -> str:
        return (datetime.now() + dt.timedelta(days=int(天))).strftime("%Y.%m.%d")

    def _取(键: str, 默认: str = "") -> str:
        v = cfg.get(键)
        return 默认 if v in (None, "") else str(v)

    params = {
        "autfp": _取("catLow", "2"),
        "prof": _取("prof", "000001"),
        "matLow": _取("matLow", "*"), "matHigh": _取("matHigh"),
        "werksLow": _取("werksLow", "0001"), "werksHigh": _取("werksHigh"),
        "dispoLow": _取("dispoLow", "060"), "dispoHigh": _取("dispoHigh"),
        "fevorLow": _取("fevorLow", "24*"), "fevorHigh": _取("fevorHigh"),
        "gstrpLow": _取("gstrpLow", _相对日期(-40)), "gstrpHigh": _取("gstrpHigh", _相对日期(100)),
        "gltrpLow": _取("gltrpLow"), "gltrpHigh": _取("gltrpHigh"),
    }
    发邮件 = cfg.get("mail")
    发邮件 = True if 发邮件 is None else bool(发邮件)
    收件人 = str(cfg.get("收件人") or "")
    lines.append("查询条件：" + " · ".join(f"{k}={v}" for k, v in params.items() if v))
    lines.append("发邮件=" + (f"开（{收件人 or '默认收件人 demo.user73@example.com'}）" if 发邮件 else "关"))

    try:
        r = vbs_run(VbsRunRequest(script="co28", params=params, mail=发邮件, 收件人=收件人,
                                  前置回车=_sap前置回车次数(cfg)))
    except HTTPException as e:
        if e.status_code == 409:
            lines.append(f"⏭ 已有 SAP GUI 脚本在运行，本轮不重复启动：{e.detail}")
            return {"ok": True, "结果": "跳过（已有 VBS 脚本在运行）"}
        lines.append(f"✗ SAP GUI 执行失败：{e.detail}")
        return {"ok": False, "结果": "SAP GUI 执行失败"}
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ SAP GUI 执行异常：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "SAP GUI 执行异常"}

    if r.get("前置回车"):
        lines.append("  " + str(r["前置回车"]))
    正常 = bool(r.get("ok"))
    尾部 = [ln.strip() for ln in (r.get("output") or "").splitlines() if ln.strip()]
    for ln in 尾部[-12:]:
        lines.append("  " + ln)
    批量失败 = r.get("batchFail") or []
    if 批量失败:
        lines.append(f"⚠ SAP 里有 {len(批量失败)} 项未写入：{'；'.join(批量失败[:5])}")
    邮件 = r.get("mail")
    if 邮件:
        lines.append(f"结果邮件：{str(邮件)[:200]}")
    耗时 = r.get("cost")
    结果 = ("运行成功" if 正常 else "运行失败") + (f"（耗时 {耗时}s）" if 耗时 else "")
    # 「无数据行可下达」时，CO28 原本准备发的是「失败邮件」，但要靠前端回传网页截图才能发 ——
    # 定时任务没有前端，于是整次运行完全静默。无人值守最忌讳这个（第二天分不清"没数据"还是"没跑"），
    # 故这里由触发器补一封简短告知邮件（config.无数据发邮件=false 可关）。
    if 正常 and isinstance(邮件, dict) and 邮件.get("无数据") and bool(cfg.get("无数据发邮件", True)):
        收 = 收件人 or "demo.user74@example.com"
        dws = _rpa导dws()
        if dws is None:
            lines.append("⚠ 无数据告知邮件未发（dws_mail 不可用）")
        else:
            try:
                正文 = ("本次 Co28 定时查询已执行完成，但当前筛选条件下没有需要下达的数据行，"
                        "因此未执行下达、也未生成截图。\n\n查询条件：\n"
                        + "\n".join(f"  {k} = {v}" for k, v in params.items() if v)
                        + f"\n\n执行耗时：{耗时}s")
                状, 说 = dws.发送邮件("【Co28】自动查询完成：今日无需下达的数据行", 正文, 收件人=收)
                lines.append(f"无数据告知邮件：{状}（{str(说)[:60]}）")
            except Exception as e:                                # noqa: BLE001
                lines.append(f"⚠ 无数据告知邮件发送异常：{type(e).__name__}: {e}")
    _apd_sync_log(f"触发器 CO28：{结果}｜邮件={str(邮件)[:120]}")
    return {"ok": 正常, "结果": 结果, "耗时s": 耗时, "batchFail": 批量失败}


def _trig_safclamp_material_sync(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 safclamp_material_sync：安全钳物料号每日增量同步（2026-11-03 用户需求）。

    与 _trig_limiter_material_sync 同一工艺，差异：
      · 脚本 = SafClamp_Sync.vbs（品目 125、部件名称 *安全钳* 排除 通用件/标签/座/构/铭牌/提拉）；
      · 图号/变量来源 =「安全钳总图号分类」表的「图号」列（用户口径）；
      · 去重目标 =「安全钳物料号」SQLite 大表（safety_material.db，40 列 c01..c40）；
      · 插入 = 新物料整行（40 列按清单表头↔meta 显示名对应）+ **倒序重排**：
        SQLite 大表 22 万行没法像小表那样整表内存排序后重写 row_no —— 用「窗口期内新行按
        创建日期倒序生成 row_no 2..n、原表全体 row_no 顺延」的位移式插入（只 UPDATE 受影响行）。
    """
    cfg = cfg or {}

    def _取(键, 默认=""):
        v = cfg.get(键)
        return 默认 if v in (None, "") else str(v)

    今天 = datetime.now()
    天窗口 = int(_取("天数窗口", "7"))
    d_from = (今天 - dt.timedelta(days=天窗口)).strftime("%Y.%m.%d")
    d_to = 今天.strftime("%Y.%m.%d")

    # 图号/变量：「安全钳总图号分类」的「图号」列（去重）
    图号们 = []
    for r in ((store_safety := json.loads((APP_DIR / "safety_store" / "local_store.json").read_text(encoding="utf-8"))).get("sheets") or {}).get("安全钳总图号分类") or []:
        v = str((r.get("data") or {}).get("图号") or "").strip()
        if v and v not in 图号们:
            图号们.append(v)
    lines.append(f"图号/变量来源：安全钳总图号分类 {len(图号们)} 个（去重后）")
    if not 图号们:
        lines.append("✗ 安全钳总图号分类没有图号，无法构造查询条件")
        return {"ok": False, "结果": "安全钳总图号分类表为空"}

    部件名称条件 = _取("部件名称", "*安全钳*,!*通用件*,!*标签*,!*座*,!*构*,!*铭牌*,!*提拉*")

    批量们 = [图号们[i:i + 180] for i in range(0, len(图号们), 180)] or [[]]
    lines.append(f"图号/变量 {len(图号们)} 个 → 分 {len(批量们)} 批查询（每批 ≤180，创建日期 {d_from}~{d_to}）")

    下载目录 = _sapq确保下载目录()
    新行们 = []
    seen_mats = set()
    任何失败 = False
    for 批次, 批 in enumerate(批量们, 1):
        批串 = ",".join(批)
        params = {
            "zb02MatLow": "",
            "zb02PirLow": _取("品目", "125"),
            "zb02PartNameLow": 部件名称条件,
            "zb02DrawVarLow": 批串,
            "zb02BrandGrpLow": _取("品群", "G*"),
            "zb02LnoLow": _取("L番", "L*"),
            "zb02ErdatLow": d_from,
            "zb02ErdatHigh": d_to,
        }
        lines.append(f"— 第 {批次}/{len(批量们)} 批（图号 {len(批)} 个）—")
        try:
            r = vbs_run(VbsRunRequest(script="safclamp_sync", params=params,
                                      download=True, 报表转换=False,
                                      下载目录=str(下载目录),
                                      前置回车=_sap前置回车次数(cfg),
                                      报表邮件=_sap任务报表邮件(cfg)))
        except HTTPException as e:
            if e.status_code == 409:
                lines.append(f"⏭ 已有 SAP GUI 脚本在运行，本轮不重复启动：{e.detail}")
                return {"ok": True, "结果": "跳过（已有 VBS 脚本在运行）"}
            lines.append(f"✗ SAP GUI 执行失败：{e.detail}")
            return {"ok": False, "结果": "SAP GUI 执行失败"}
        except Exception as e:
            lines.append(f"✗ SAP GUI 执行异常：{type(e).__name__}: {e}")
            return {"ok": False, "结果": "SAP GUI 执行异常"}

        if r.get("前置回车"):
            lines.append("  " + str(r["前置回车"]))
        if not r.get("ok"):
            lines.append(f"✗ 第 {批次} 批 VBS 运行失败（耗时 {r.get('cost')}s），输出尾部：")
            for ln in (r.get("output") or "").splitlines()[-8:]:
                lines.append("  " + ln)
            任何失败 = True
            break
        dl = r.get("download") or {}
        if not dl.get("ok") or not dl.get("路径"):
            lines.append(f"✗ 第 {批次} 批没拿到导出报表：{dl.get('vbs') or dl}")
            任何失败 = True
            break
        报表 = Path(dl["路径"])
        lines.append(f"  报表已落盘：{报表.name}")

        # ---- 解析清单（与 6.xls 灌库同口径：第4行表头、第5行空、第6行起数据、真实记录 集团=800）----
        try:
            原始 = 报表.read_bytes().decode("gbk", errors="replace")
        except Exception as e:
            lines.append(f"✗ 报表读取失败：{e}")
            任何失败 = True
            break
        行文本 = 原始.split("\r\n")
        if len(行文本) < 6:
            lines.append(f"  第 {批次} 批无数据行（清单 {len(行文本)} 行）")
            continue
        表头原 = [c.strip() for c in 行文本[3].split("\t")]
        # 与 meta.json 同款列名规范：空白列→(空白列N)、重复名→物料(2)/标记(2)
        meta_cols = _safety_material_cols()
        列名映射 = []          # (表头位置, c键) —— 按显示名对回 meta
        seen_names = {}
        disp_to_key = {str(c["display"]).strip(): c["key"] for c in meta_cols}
        for j, nm in enumerate(表头原):
            if not nm:
                continue
            if nm in seen_names:
                seen_names[nm] += 1
                disp = f"{nm}({seen_names[nm]})"
            else:
                seen_names[nm] = 1
                disp = nm
            k = disp_to_key.get(disp) or disp_to_key.get(nm)
            if k:
                列名映射.append((j, k, nm))
        本批 = 0
        for ln in 行文本[5:]:
            if not ln.strip():
                continue
            cells = ln.split("\t")
            集团 = cells[1].strip() if len(cells) > 1 else ""
            if 集团 != "800":
                continue
            data = {}
            for j, k, nm in 列名映射:
                data[k] = cells[j].strip() if len(cells) > j else ""
            物料 = data.get("c03", "")
            if not 物料 or 物料 in seen_mats:
                continue
            seen_mats.add(物料)
            新行们.append(data)
            本批 += 1
        lines.append(f"  解析出物料 {本批} 个（累计 {len(新行们)}）")

    if 任何失败:
        return {"ok": False, "结果": "VBS 查询/导出失败（详见日志），本轮不写数据"}

    if not 新行们:
        lines.append("✓ 窗口内没有查到任何安全钳物料（无需同步）")
        return {"ok": True, "结果": "运行成功（无新数据）", "新增": 0}

    # ---- 去重 + 倒序插入（SQLite 大表：位移式插入，新行按创建日期倒序占 row_no 2..n+1）----
    import sqlite3
    新物料 = [str(d.get("c03") or "").strip() for d in 新行们]
    conn = sqlite3.connect(str(SAFETY_MATERIAL_DB))
    conn.row_factory = sqlite3.Row
    try:
        cur = conn.cursor()
        ph = ",".join(["?"] * len(新物料))
        已有 = {r[0] for r in cur.execute(f'SELECT DISTINCT "c03" FROM "material" WHERE "c03" IN ({ph})', 新物料)}
        真新增 = [d for d in 新行们 if str(d.get("c03") or "").strip() not in 已有]
        lines.append(f"查回物料 {len(新行们)} 个，其中新物料 {len(真新增)} 个（已存在 {len(新行们) - len(真新增)} 个跳过）")
        if not 真新增:
            conn.commit()
            lines.append("✓ 没有新物料号，现有表不变")
            return {"ok": True, "结果": "运行成功（无新增）", "新增": 0}

        def _日期键(d):
            for k in ("c18", "c32"):
                v = str(d.get(k) or "").strip()
                try:
                    return dt.datetime.strptime(v, "%Y.%m.%d")
                except Exception:
                    continue
            return dt.datetime(1970, 1, 1)
        真新增.sort(key=_日期键, reverse=True)     # 倒序：最新在最前

        n_ins = len(真新增)
        max_no = cur.execute('SELECT MAX("row_no") FROM "material"').fetchone()[0] or 0
        # 1) 原表全体 row_no 顺延 n_ins —— row_no 有 UNIQUE 约束，直接 +n 会边移边撞；
        #    两步走：先整体加一个大偏移（全部越过 MAX），再减回 n_ins，两步都不会撞已占用的号段
        _偏移 = max_no + n_ins
        cur.execute('UPDATE "material" SET "row_no" = "row_no" + ?', [_偏移])
        cur.execute('UPDATE "material" SET "row_no" = "row_no" - ?', [n_ins])
        # 2) 新行写 row_no 2..n_ins+1（row_no=1 空缺？——不：最新那行给 row_no 1）
        cols = [c["key"] for c in _safety_material_cols()]
        for i, d in enumerate(真新增):
            vals = [i + 1] + [str(d.get(k) or "") for k in cols]
            ph2 = ",".join(["?"] * (len(cols) + 1))
            sel = ", ".join([f'"{k}"' for k in cols])
            cur.execute(f'INSERT INTO "material" ("row_no",{sel}) VALUES ({ph2})', vals)
        conn.commit()
        lines.append(f"✓ 已按 创建日期 倒序插入 {n_ins} 行到「安全钳物料号」最前（原表全体顺延，总行数 {max_no + n_ins}）")
    except Exception as e:
        conn.rollback()
        lines.append(f"✗ 写库失败（已回滚，本轮不写任何数据）：{type(e).__name__}: {e}")
        return {"ok": False, "结果": f"写库失败：{e}"}
    finally:
        conn.close()
    return {"ok": True, "结果": f"运行成功（新增 {len(真新增)} 行）", "新增": len(真新增),
            "查询回": len(新行们)}


# 限速器物料号页面列头 ↔ ZLSTMA 清单显示名 ↔ SQLite c键（合并同步按品目分流时用）
SF_DEFS_HEADERS_LIMITER = ['集团', '物料', '物料(类型)', '品目', '部件名称', '图号/变量', '品/群', 'L 番', '创建日期', '创建者', '物料传输标记', 'BOM 传输标志', '生产科使用', '变量', '物料类型', '物料号长参数表', '标记', '工艺更改版本编号', '工艺生效日期', '工序数量', '可选工艺数量', '并行工艺数量', '输入日期', '标记2', '印板工艺属性', '材料', 'Changed at', 'RDC标记', '海菱用物料标记', '海菱物料传输标记', '海菱BOM传输标记']
# 页面显示名 → SQLite meta 显示名（物料(类型)→物料(2)、标记2→标记(2)；其余同名）
SF_LIMITER_KEY_TO_DISPLAY = {
    '物料(类型)': '物料(2)',
    '标记2': '标记(2)',
}
# 115 插入限速器物料号（JSON store，键=页面显示名）时的取值来源：meta 显示名 → c键
# （解析清单时数据落在 c键 上；页面列头先经 SF_LIMITER_KEY_TO_DISPLAY 转成 meta 显示名再取 c键）
def _sf_limiter_display_row(cdata: dict) -> dict:
    """c键数据 → 限速器物料号页面行（显示名键）。"""
    meta = {}
    try:
        meta = {str(c["display"]).strip(): c["key"] for c in _safety_material_cols()}
    except Exception:
        pass
    out = {}
    for 页面列 in SF_DEFS_HEADERS_LIMITER:
        meta显示 = SF_LIMITER_KEY_TO_DISPLAY.get(页面列, 页面列)
        c键 = meta.get(meta显示, "")
        out[页面列] = str(cdata.get(c键) or "").strip() if c键 else ""
    return out


def _trig_materials_sync_daily(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 materials_sync_daily：安全钳+限速器合并调取，按品目分流落表（2026-11-04 用户需求）。

    用户口径：
      · 一次调取 —— 图号=「安全钳总图号分类」+「限速器参数」两表汇总，且**每个图号前后加 * 通配**
        （P115007A000 → *P115007A000*，防止 SAP 精确匹配漏掉带后缀的变体图号）；
      · 部件名称 = *安全钳*,*限速器* + 排除（两套合并）：!*通用件*,!*标签*,!*座*,!*构*,!*机构*,!*铭牌*,!*提拉*；
      · 品目 = 115,125（多值）；品/群 G*；L 番 L*；创建日期 今天-7 ~ 今天；
      · 落表按「品目」分类：**115 → 限速器物料号**（JSON store），**125 → 安全钳物料号**（SQLite 大表），
        插入逻辑与各自原链路完全一致（去重 → 按创建日期倒序插最前）。
    脚本 = SafClamp_Sync.vbs（同一份 Zb02 拷贝版；条件区全部由 Python 按次注入，脚本内默认值仅兜底）。
    """
    cfg = cfg or {}

    def _取(键, 默认=""):
        v = cfg.get(键)
        return 默认 if v in (None, "") else str(v)

    今天 = datetime.now()
    天窗口 = int(_取("天数窗口", "7"))
    d_from = (今天 - dt.timedelta(days=天窗口)).strftime("%Y.%m.%d")
    d_to = 今天.strftime("%Y.%m.%d")

    # ---- 图号：两表汇总 + 前后加 * 通配（用户 2026-11-04 口径）----
    def _加星(v):
        v = str(v or "").strip()
        if not v:
            return ""
        return v if v.startswith("*") else "*" + (v.rstrip("*") if v.endswith("*") else v) + "*"

    图号们 = []
    def _收(源名, 键名):
        try:
            st = json.loads((APP_DIR / "safety_store" / "local_store.json").read_text(encoding="utf-8"))
        except Exception:
            return
        for r in (st.get("sheets") or {}).get(源名) or []:
            v = str((r.get("data") or {}).get(键名) or "").strip()
            if v and v not in 图号们:
                图号们.append(v)
    _收("安全钳总图号分类", "图号")
    _收("限速器参数", "图号/变量")
    图号星 = []
    for v in 图号们:
        s = _加星(v)
        if s not in 图号星:
            图号星.append(s)
    lines.append(f"图号来源：安全钳总图号分类 + 限速器参数 合计 {len(图号们)} 个（去重后）→ 加 * 通配 {len(图号星)} 个")

    部件名称条件 = _取("部件名称",
        "*安全钳*,*限速器*,!*通用件*,!*标签*,!*座*,!*构*,!*机构*,!*铭牌*,!*提拉*")
    品目条件 = _取("品目", "115,125")

    批量们 = [图号星[i:i + 180] for i in range(0, len(图号星), 180)] or [[]]
    lines.append(f"分 {len(批量们)} 批查询（每批 ≤180，创建日期 {d_from}~{d_to}，品目 {品目条件}）")

    下载目录 = _sapq确保下载目录()
    新行们 = []          # [{data(c键或显示名→值), 品目: "115"/"125"}]
    seen_mats = set()
    任何失败 = False
    报表们 = []
    for 批次, 批 in enumerate(批量们, 1):
        params = {
            "zb02MatLow": "",
            "zb02PirLow": "~" + 品目条件.replace(",", ",~"),   # 115,125 → ~115,~125 强制走多重选择框（3位数字串会被千分位判定拦下）
            "zb02PartNameLow": 部件名称条件,     # 含/排除 → 多重选择框
            "zb02DrawVarLow": ",".join(批),      # *P115007A000* … → 多重选择框
            "zb02BrandGrpLow": _取("品群", "G*"),
            "zb02LnoLow": _取("L番", "L*"),
            "zb02ErdatLow": d_from,
            "zb02ErdatHigh": d_to,
        }
        lines.append(f"— 第 {批次}/{len(批量们)} 批（图号 {len(批)} 个）—")
        try:
            r = vbs_run(VbsRunRequest(script="safclamp_sync", params=params,
                                      download=True, 报表转换=False,
                                      下载目录=str(下载目录),
                                      前置回车=_sap前置回车次数(cfg),
                                      报表邮件=_sap任务报表邮件(cfg)))
        except HTTPException as e:
            if e.status_code == 409:
                lines.append(f"⏭ 已有 SAP GUI 脚本在运行，本轮不重复启动：{e.detail}")
                return {"ok": True, "结果": "跳过（已有 VBS 脚本在运行）"}
            lines.append(f"✗ SAP GUI 执行失败：{e.detail}")
            return {"ok": False, "结果": "SAP GUI 执行失败"}
        except Exception as e:
            lines.append(f"✗ SAP GUI 执行异常：{type(e).__name__}: {e}")
            return {"ok": False, "结果": "SAP GUI 执行异常"}

        if r.get("前置回车"):
            lines.append("  " + str(r["前置回车"]))
        if not r.get("ok"):
            lines.append(f"✗ 第 {批次} 批 VBS 运行失败（耗时 {r.get('cost')}s），输出尾部：")
            for ln in (r.get("output") or "").splitlines()[-8:]:
                lines.append("  " + ln)
            任何失败 = True
            break
        dl = r.get("download") or {}
        if not dl.get("ok") or not dl.get("路径"):
            lines.append(f"✗ 第 {批次} 批没拿到导出报表：{dl.get('vbs') or dl}")
            任何失败 = True
            break
        报表 = Path(dl["路径"])
        报表们.append(报表)
        lines.append(f"  报表已落盘：{报表.name}")

        # ---- 解析清单（第4行表头/第6行起数据、真实记录 集团=800；列名按两套目标的并集映射）----
        try:
            原始 = 报表.read_bytes().decode("gbk", errors="replace")
        except Exception as e:
            lines.append(f"✗ 报表读取失败：{e}")
            任何失败 = True
            break
        行文本 = 原始.split("\r\n")
        if len(行文本) < 6:
            lines.append(f"  第 {批次} 批无数据行（清单 {len(行文本)} 行）")
            continue
        表头原 = [c.strip() for c in 行文本[3].split("\t")]
        meta_cols = _safety_material_cols()
        disp_to_key = {str(c["display"]).strip(): c["key"] for c in meta_cols}
        seen_names, 映射 = {}, []
        for j, nm in enumerate(表头原):
            if not nm:
                continue
            if nm in seen_names:
                seen_names[nm] += 1
                disp = f"{nm}({seen_names[nm]})"
            else:
                seen_names[nm] = 1
                disp = nm
            k = disp_to_key.get(disp) or disp_to_key.get(nm)
            if k:
                映射.append((j, k, nm))
        本批 = 0
        for ln in 行文本[5:]:
            if not ln.strip():
                continue
            cells = ln.split("\t")
            集团 = cells[1].strip() if len(cells) > 1 else ""
            if 集团 != "800":
                continue
            data = {}
            品目 = ""
            for j, k, nm in 映射:
                v = cells[j].strip() if len(cells) > j else ""
                data[k] = v
                if nm == "品目":
                    品目 = v
            物料 = data.get("c03", "")
            if not 物料 or 物料 in seen_mats:
                continue
            seen_mats.add(物料)
            新行们.append({"data": data, "品目": 品目})
            本批 += 1
        lines.append(f"  解析出物料 {本批} 个（累计 {len(新行们)}）")

    if 任何失败:
        return {"ok": False, "结果": "VBS 查询/导出失败（详见日志），本轮不写数据"}
    if not 新行们:
        lines.append("✓ 窗口内没有查到任何新物料（无需同步）")
        return {"ok": True, "结果": "运行成功（无新数据）", "新增限速器": 0, "新增安全钳": 0}

    # ---- 按品目分流落表（115→限速器 JSON store / 125→安全钳 SQLite 大表；插入逻辑不变）----
    def _日期键(c键组, d):
        for k in c键组:
            v = str(d.get(k) or "").strip()
            try:
                return dt.datetime.strptime(v, "%Y.%m.%d")
            except Exception:
                continue
        return dt.datetime(1970, 1, 1)

    # 115 → 限速器物料号（JSON store：显示名键）
    行115 = [x for x in 新行们 if str(x.get("品目") or "").strip() == "115"]
    新增115 = 0
    明细115 = []          # 真新增行（组装通知清单用）
    if 行115:
        with _TRIGGER_STORE_LOCK:
            s2 = _safety_store_load()
            rows = s2["sheets"].setdefault("限速器物料号", [])
            现有 = {str((r.get("data") or {}).get("物料") or "").strip() for r in rows}
            真115 = [x for x in 行115 if str(x["data"].get("c03") or "").strip() not in 现有]
            if 真115:
                def _日期键115(d):
                    return _日期键(("创建日期", "输入日期"), d)
                真115.sort(key=lambda x: _日期键115(x["data"]), reverse=True)
                nid = s2.get("next_id", 1)
                头部 = []
                for x in 真115:
                    头部.append({"id": nid, "row_no": 0, "data": _sf_limiter_display_row(x["data"])})
                    nid += 1
                s2["next_id"] = nid
                全部 = 头部 + rows
                for i, r in enumerate(全部, 1):
                    r["row_no"] = i
                s2["sheets"]["限速器物料号"] = 全部
                _safety_store_save(s2)
                新增115 = len(真115)
                明细115 = [x["data"] for x in 真115]
        lines.append(f"品目 115（限速器）：查回 {len(行115)}，新增 {新增115}（倒序插最前）")
    else:
        lines.append("品目 115（限速器）：本批无数据")

    # 125 → 安全钳物料号（SQLite 大表：c 键直写 + 两步偏移位移式插入）
    行125 = [x for x in 新行们 if str(x.get("品目") or "").strip() == "125"]
    新增125 = 0
    明细125 = []          # 真新增行（组装通知清单用）
    if 行125:
        import sqlite3
        conn = sqlite3.connect(str(SAFETY_MATERIAL_DB))
        conn.row_factory = sqlite3.Row
        try:
            cur = conn.cursor()
            新物料 = [str(x["data"].get("c03") or "").strip() for x in 行125]
            ph = ",".join(["?"] * len(新物料))
            已有 = {r[0] for r in cur.execute(f'SELECT DISTINCT "c03" FROM "material" WHERE "c03" IN ({ph})', 新物料)}
            真125 = [x for x in 行125 if str(x["data"].get("c03") or "").strip() not in 已有]
            if 真125:
                def _日期键125(x):
                    return _日期键(("c18", "c32"), x["data"])
                真125.sort(key=_日期键125, reverse=True)
                n_ins = len(真125)
                max_no = cur.execute('SELECT MAX("row_no") FROM "material"').fetchone()[0] or 0
                _偏移 = max_no + n_ins
                cur.execute('UPDATE "material" SET "row_no" = "row_no" + ?', [_偏移])
                cur.execute('UPDATE "material" SET "row_no" = "row_no" - ?', [n_ins])
                cols = [c["key"] for c in _safety_material_cols()]
                for i, x in enumerate(真125):
                    vals = [i + 1] + [str(x["data"].get(k) or "") for k in cols]
                    ph2 = ",".join(["?"] * (len(cols) + 1))
                    sel = ", ".join([f'"{k}"' for k in cols])
                    cur.execute(f'INSERT INTO "material" ("row_no",{sel}) VALUES ({ph2})', vals)
                conn.commit()
                新增125 = n_ins
                明细125 = [x["data"] for x in 真125]
            else:
                conn.commit()
        except Exception as e:
            conn.rollback()
            lines.append(f"✗ 安全钳写库失败（已回滚）：{type(e).__name__}: {e}")
            return {"ok": False, "结果": f"安全钳写库失败：{e}"}
        finally:
            conn.close()
        lines.append(f"品目 125（安全钳）：查回 {len(行125)}，新增 {新增125}（倒序插最前）")
    else:
        lines.append("品目 125（安全钳）：本批无数据")

    # ---- 通知正文（用户 2026-11-04 口径）：xxx个115新生物料以及yyy个125新生物料已汇总插入，具体为：…----
    def _清单文本(明细们, 显示名模式):
        """明细行 → 按创建日期倒序的清单文本；显示名模式=True（115）用显示名键，False（125）用 c键。"""
        if not 明细们:
            return "（无新增）"
        def _键(d):
            for k in (("创建日期", "输入日期") if 显示名模式 else ("c18", "c32")):
                v = str(d.get(k) or "").strip()
                try:
                    return dt.datetime.strptime(v, "%Y.%m.%d")
                except Exception:
                    continue
            return dt.datetime(1970, 1, 1)
        排序后 = sorted(明细们, key=_键, reverse=True)
        行文本 = []
        for d in 排序后:
            if 显示名模式:
                行文本.append("· " + str(d.get("物料") or "").strip() + "　" + str(d.get("图号/变量") or "").strip() + "　" + str(d.get("创建日期") or "").strip())
            else:
                行文本.append("· " + str(d.get("c03") or "").strip() + "　" + str(d.get("c15") or "").strip() + "　" + str(d.get("c18") or "").strip())
        return "（共 " + str(len(排序后)) + " 个，按创建日期倒序）" + chr(10) + chr(10).join(行文本)

    _新115 = 新增115 if 行115 else 0
    _新125 = 新增125 if 行125 else 0
    通知正文 = ("安全钳+限速器物料同步完成（查询窗口 " + d_from + " ~ " + d_to + "）。" + chr(10) + chr(10)
               + str(_新115) + "个115新生物料以及" + str(_新125) + "个125新生物料已汇总插入，具体为：" + chr(10) + chr(10)
               + "115：" + chr(10) + _清单文本(明细115, True) + chr(10) + chr(10)
               + "125：" + chr(10) + _清单文本(明细125, False))
    通知主题 = "【安全钳+限速器物料同步】115新增 " + str(_新115) + " 个 / 125新增 " + str(_新125) + " 个"
    通知钉钉 = ("【安全钳+限速器物料同步】" + str(_新115) + "个115新生物料以及" + str(_新125) + "个125新生物料已汇总插入" + chr(10) + chr(10)
               + "115：" + chr(10) + _清单文本(明细115, True) + chr(10) + chr(10)
               + "125：" + chr(10) + _清单文本(明细125, False))
    if not 新增115 and not 新增125:
        lines.append("✓ 没有新物料号，两表均不变")
        return {"ok": True, "结果": "运行成功（无新增）", "新增限速器": 0, "新增安全钳": 0,
                "邮件主题": f"【安全钳+限速器物料同步】今日无新增（{d_from}~{d_to}）",
                "邮件主题": "【安全钳+限速器物料同步】今日无新增（" + d_from + "~" + d_to + "）",
                "邮件正文": "安全钳+限速器物料同步完成（查询窗口 " + d_from + " ~ " + d_to + "）。" + chr(10) + "今日无新增物料，两表均不变。",
                "钉钉正文": f"【安全钳+限速器物料同步】今日无新增物料（{d_from}~{d_to}），两表均不变。"}

    # MySQL 落地表同步（失败不影响主流程）
    try:
        counts = _apd22_rebuild_tables()
        lines.append("03#箱databoard 落地表已重建：" + " · ".join(f"{k} {v}" for k, v in (counts or {}).items()))
    except Exception as e:
        lines.append(f"[警告] 落地表重建失败（下次刷新会重建）：{e}")
    return {"ok": True, "结果": f"运行成功（限速器 +{新增115}，安全钳 +{新增125}）",
            "新增限速器": 新增115, "新增安全钳": 新增125, "查询回": len(新行们),
            "邮件主题": 通知主题, "邮件正文": 通知正文, "钉钉正文": 通知钉钉}


def _trig_yuzhi_mail(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 yuzhi_mail_check：钉钉邮箱巡检「小于阀值异常工单明细」（模块 app/yuzhi_mail_check.py）。

    重活（dws 检索邮件 / 下载附件 / 解析 Excel / 判零去重）都在模块里，这里只负责：
      · 模块 import 失败不炸调度器，改成一行可读的失败原因；
      · 把模块产出的正文挂到 r["钉钉正文"]，由 config.通知 统一投递（钉钉消息通道用这段正文）。
    """
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    try:
        import yuzhi_mail_check as Y
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 巡检模块不可用：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "巡检模块未加载"}
    try:
        r = Y.巡检(lines, cfg or {})
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 巡检异常：{type(e).__name__}: {e}")
        return {"ok": False, "结果": f"巡检异常：{e}"}
    _apd_sync_log(f"触发器 钉钉邮箱巡检：{r.get('结果')}｜命中 {r.get('命中行数')} 行")
    return r


def _trig_limiter_material_sync(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 limiter_material_sync：限速器物料号每日增量同步（2026-11-03 用户需求）。

    链路（用户口径）：
      ① 跑 vbs_scripts/Zlstma_Sync.vbs（从 Zb02.vbs 整本拷贝改造，原脚本不动）：
         /nzb02 数据浏览器 ZLSTMA 选择屏，条件按用户指定 + 创建日期=今天-7 自然日 ~ 今天（按次注入），
         图号/变量 =「限速器参数」基础数据全部图号/变量（动态生成，逗号串走多重选择框），
         执行后导出清单（功能名_时间戳.xls，与页面「报表下载」同一导出链）。
      ② Python 解析清单（GBK 制表符文本：第4行表头/第5行起数据，真实记录 集团=800）。
      ③ 按「物料」列与「限速器物料号」基础数据去重：已存在跳过；新的整行插入，按 创建日期
         倒序（最新的在最前）。多列值按字段名一一对应（清单表头=页面列头，重复列名同 2026-11-03 口径）。
    不驱动 SAP 之外的任何系统；VBS 失败/无报表时如实报失败，绝不写半截数据。
    """
    cfg = cfg or {}

    def _取(键, 默认=""):
        v = cfg.get(键)
        return 默认 if v in (None, "") else str(v)

    今天 = datetime.now()
    天窗口 = int(_取("天数窗口", "7"))
    d_from = (今天 - dt.timedelta(days=天窗口)).strftime("%Y.%m.%d")
    d_to = 今天.strftime("%Y.%m.%d")

    # 图号/变量：取「限速器参数」表全部图号/变量（用户口径：限速器参数标题栏 中的图号/变量）
    store = _safety_store_seed_if_empty()
    图号们 = []
    for r in (store["sheets"].get("限速器参数") or []):
        v = str((r.get("data") or {}).get("图号/变量") or "").strip()
        if v and v not in 图号们:
            图号们.append(v)
    lines.append(f"图号/变量来源：限速器参数 {len(图号们)} 个（去重后）")
    if not 图号们:
        lines.append("✗ 限速器参数表没有图号/变量，无法构造查询条件")
        return {"ok": False, "结果": "限速器参数表为空"}

    部件名称条件 = _取("部件名称", "*限速器*,!*通用件*,!*标签*,!*座*,!*机构*,!*铭牌*")

    # 图号超过 180 个就分批查（单字段多重选择框一次放 200 值以内最稳）
    批量们 = [图号们[i:i + 180] for i in range(0, len(图号们), 180)] or [[]]
    lines.append(f"图号/变量 {len(图号们)} 个 → 分 {len(批量们)} 批查询（每批 ≤180，创建日期 {d_from}~{d_to}）")

    下载目录 = _sapq确保下载目录()
    新行们 = []                    # 累计各批解析出的整行 dict（字段名 → 值）
    seen_mats = set()
    任何失败 = False
    for 批次, 批 in enumerate(批量们, 1):
        # 同一图号多批时兜底也查图号前缀：不做（用户条件就是精确图号列表）
        批串 = ",".join(批)
        params = {
            "zb02MatLow": "",
            "zb02PirLow": _取("品目", "115"),
            "zb02PartNameLow": 部件名称条件,          # 多值（含/排除）→ BatchWrite 多重选择框
            "zb02DrawVarLow": 批串,                  # 多值 → 多重选择框
            "zb02BrandGrpLow": _取("品群", "G*"),
            "zb02LnoLow": _取("L番", "L*"),
            "zb02ErdatLow": d_from,
            "zb02ErdatHigh": d_to,
        }
        lines.append(f"— 第 {批次}/{len(批量们)} 批（图号 {len(批)} 个）—")
        try:
            r = vbs_run(VbsRunRequest(script="zlstma_sync", params=params,
                                      download=True, 报表转换=False,
                                      下载目录=str(下载目录),
                                      前置回车=_sap前置回车次数(cfg),
                                      报表邮件=_sap任务报表邮件(cfg)))
        except HTTPException as e:
            if e.status_code == 409:
                lines.append(f"⏭ 已有 SAP GUI 脚本在运行，本轮不重复启动：{e.detail}")
                return {"ok": True, "结果": "跳过（已有 VBS 脚本在运行）"}
            lines.append(f"✗ SAP GUI 执行失败：{e.detail}")
            return {"ok": False, "结果": "SAP GUI 执行失败"}
        except Exception as e:
            lines.append(f"✗ SAP GUI 执行异常：{type(e).__name__}: {e}")
            return {"ok": False, "结果": "SAP GUI 执行异常"}

        if r.get("前置回车"):
            lines.append("  " + str(r["前置回车"]))
        if not r.get("ok"):
            lines.append(f"✗ 第 {批次} 批 VBS 运行失败（耗时 {r.get('cost')}s），输出尾部：")
            for ln in (r.get("output") or "").splitlines()[-8:]:
                lines.append("  " + ln)
            任何失败 = True
            break
        dl = r.get("download") or {}
        if not dl.get("ok") or not dl.get("路径"):
            lines.append(f"✗ 第 {批次} 批没拿到导出报表：{dl.get('vbs') or dl}")
            任何失败 = True
            break
        报表 = Path(dl["路径"])
        lines.append(f"  报表已落盘：{报表.name}")

        # ---- 解析清单（GBK 制表符文本，第4行表头/第5行起数据；与 2026-11-03 种子同口径）----
        try:
            原始 = 报表.read_bytes().decode("gbk", errors="replace")
        except Exception as e:
            lines.append(f"✗ 报表读取失败：{e}")
            任何失败 = True
            break
        行文本 = [x for x in 原始.split("\r\n")]
        if len(行文本) < 5:
            lines.append(f"  第 {批次} 批无数据行（清单 {len(行文本)} 行）")
            continue
        表头原 = [c.strip() for c in 行文本[3].split("\t")]
        列位 = [j for j, c in enumerate(表头原) if c]
        # 重复列名去重（与种子同口径：物料(类型)/标记2）
        seen_names = {}
        列名 = []
        for j in 列位:
            nm = 表头原[j]
            if nm in seen_names:
                seen_names[nm] += 1
                列名.append((j, f"{nm}(类型)" if nm == "物料" else f"{nm}{seen_names[nm]}"))
            else:
                seen_names[nm] = 1
                列名.append((j, nm))
        本批 = 0
        for ln in 行文本[4:]:
            if not ln.strip():
                continue
            cells = ln.split("\t")
            集团 = cells[1].strip() if len(cells) > 1 else ""
            if 集团 != "800":
                continue                      # 换行碎片行/异常行
            data = {}
            for j, nm in 列名:
                data[nm] = cells[j].strip() if len(cells) > j else ""
            物料 = data.get("物料", "")
            if not 物料 or 物料 in seen_mats:
                continue
            seen_mats.add(物料)
            新行们.append(data)
            本批 += 1
        lines.append(f"  解析出物料 {本批} 个（累计 {len(新行们)}）")

    if 任何失败:
        return {"ok": False, "结果": "VBS 查询/导出失败（详见日志），本轮不写数据"}

    if not 新行们:
        lines.append("✓ 窗口内没有查到任何限速器物料（无需同步）")
        return {"ok": True, "结果": "运行成功（无新数据）", "新增": 0}

    # ---- 去重 + 倒序插入（用户口径：查回的物料号在现有「物料」列里找，没有才插；最新插最前）----
    现有 = {str((r.get("data") or {}).get("物料") or "").strip() for r in (store["sheets"].get("限速器物料号") or [])}
    新增 = [d for d in 新行们 if str(d.get("物料") or "").strip() not in 现有]
    lines.append(f"查回物料 {len(新行们)} 个，其中新物料 {len(新增)} 个（已存在 {len(新行们) - len(新增)} 个跳过）")
    if not 新增:
        lines.append("✓ 没有新物料号，现有表不变")
        return {"ok": True, "结果": "运行成功（无新增）", "新增": 0}

    def _日期键(d):
        # 创建日期形如 2026.11.03；解析不了排最后（新物料里创建日期为空的兜底按输入日期/今天）
        for k in ("创建日期", "输入日期"):
            v = str(d.get(k) or "").strip()
            try:
                return dt.datetime.strptime(v, "%Y.%m.%d")
            except Exception:
                continue
        return dt.datetime(1970, 1, 1)

    新增.sort(key=_日期键, reverse=True)      # 倒序：最新在最前
    with _TRIGGER_STORE_LOCK:
        s2 = _safety_store_load()
        rows = s2["sheets"].setdefault("限速器物料号", [])
        nid = s2.get("next_id", 1)
        头部 = []
        for d in 新增:
            头部.append({"id": nid, "row_no": 0, "data": d})
            nid += 1
        s2["next_id"] = nid
        全部 = 头部 + rows
        for i, r in enumerate(全部, 1):
            r["row_no"] = i
        s2["sheets"]["限速器物料号"] = 全部
        _safety_store_save(s2)
    lines.append(f"✓ 已按 创建日期 倒序插入 {len(新增)} 行到「限速器物料号」最前（总行数 {len(全部)}）")
    # MySQL 落地表异步跟上（失败不影响主流程，下次刷新会重建）
    try:
        counts = _apd22_rebuild_tables()
        lines.append("03#箱databoard 落地表已重建：" + " · ".join(f"{k} {v}" for k, v in (counts or {}).items()))
    except Exception as e:
        lines.append(f"[警告] 落地表重建失败（下次 /api/apd22/refresh 会重建）：{e}")
    return {"ok": True, "结果": f"运行成功（新增 {len(新增)} 行）", "新增": len(新增),
            "查询回": len(新行们)}


def _trig_metric_alert(lines: list, t0, cfg: dict = None) -> dict:
    """条件触发器执行体：看板指标 vs 阈值，命中才发告警邮件（模块 app/board_push.py:run_alert）。

    阈值/指标/比较符来自任务 config（触发器编辑器里填），没填就回落 board_push 的默认值。
    """
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    try:
        import board_push as B
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 看板推送模块不可用：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "看板推送模块未加载"}
    覆盖 = {}
    for k in ("指标", "比较", "阈值", "收件人"):
        v = (cfg or {}).get(k)
        if v not in (None, "", []):
            覆盖[k] = v
    try:
        # 交由通知=True：命中时只产出（告警正文 + KPI 卡截图 + 明细 Excel），不自己发邮件；
        # 发不发、发给谁、走哪个通道，全由本任务的「通知」配置决定。
        return B.run_alert(lines, 覆盖, 交由通知=True)
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 阈值判断异常：{type(e).__name__}: {e}")
        return {"ok": False, "结果": f"阈值判断异常：{e}"}


def _trig_board_push(lines: list, t0, cfg: dict = None) -> dict:
    """触发器应用 apd22_board_push 的执行体：看板长截图 + Excel（模块 app/board_push.py）。

    重活（无头浏览器截图、dws 发信）都在 board_push 里，这里只保证：
      · 模块 import 失败不炸调度器，改成一行可读的失败原因；
      · 模块自己往 lines 里写流水（带 HH:MM:SS 前缀），运行日志 sheet 直接落库。

    交由通知=True（2026-10-26 用户口径「我只配了钉钉消息，凭什么还收到邮件」）：
      本模块只产出截图/Excel/正文，**不自己发邮件**；投递一律由任务 config.通知 决定。
    """
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    try:
        import board_push as B
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 看板推送模块不可用：{type(e).__name__}: {e}")
        return {"ok": False, "结果": "看板推送模块未加载"}
    try:
        return B.run_push(lines, 交由通知=True)
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"✗ 看板推送异常：{type(e).__name__}: {e}")
        return {"ok": False, "结果": f"推送异常：{e}"}


def _trig_artifact_shots(r: dict) -> list:
    """从执行结果里挑出**真实产物**（看板长截图 / KPI 卡截图），组装成「截图记录」。

    2026-10-26 用户反馈：「运行完成截图怎么都是计划任务列表这个页面，我要的是跑完的实际情况」。
    → 产物本身就是最真实的"完成画面"，直接登记进来；前端就不再截那个没用的页面。

    返回 [{t, title, url}]，url 指向 /api/trigger/artifact（受白名单保护）。
    """
    out: list = []
    r = r or {}
    # 只认产物路径字段（png = 看板长截图 / KPI 卡截图）；
    # 「截图」键是日志行里的截图记录（字典列表），不是路径，不要混进来。
    候选 = [(r.get("png"), "运行产物截图")]
    for 路径, 标题 in 候选:
        if not 路径:
            continue
        try:
            p = Path(str(路径))
            if not p.is_absolute():
                p = (APP_DIR / p)
            p = p.resolve()
            if not p.is_file():
                continue
            相对 = p.relative_to(APP_DIR.resolve())
        except Exception:
            continue
        out.append({"t": datetime.now().strftime("%H:%M:%S"), "title": 标题,
                    "url": "/api/trigger/artifact?path=" + str(相对).replace("\\", "/")})
    return out


def _trig_evidence_shot(应用: str, lines: list) -> list:
    """失败现场截图：运行失败且没有任何产物时，截一张相关页面当证据。

    2026-10-27 用户反馈「运行日志截图也没有」——失败详情里空着一块啥也看不到。
    这里尽力截一张：看板类应用截看板页；其它应用截触发器页（运行日志所在页）。
    截不到就返回空列表（绝不因截图失败影响主流程）。
    """
    落盘 = APP_DIR / "trigger_store" / "shots"
    try:
        落盘.mkdir(parents=True, exist_ok=True)
        fname = "trigshot_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f") + ".png"
        目标 = 落盘 / fname
        # 看板类应用 → 截看板页；其它 → 截触发器页
        if "看板" in (应用 or "") or "阈值" in (应用 or ""):
            try:
                import board_push as B
                png = B.capture_png(lambda m: lines.append(f"[截图] {m}"), B.load_config())
                if png and Path(png).is_file():
                    # 复制进 shots 目录（artifact 接口只放行 out 与 shots 两个目录）
                    import shutil
                    shutil.copy2(str(png), str(目标))
                    return [{"t": datetime.now().strftime("%H:%M:%S"), "title": "失败现场截图（看板）",
                             "url": "/api/trigger/shot?file=" + fname}]
            except Exception as e:                            # noqa: BLE001
                lines.append(f"[截图] 失败现场截图未取到：{type(e).__name__}: {e}")
        return []
    except Exception:
        return []


def _trig_write_log(应用: str, 触发方式: str, 结果: str, 日志行: list,
                    主题: str = "", 失败: bool = False, 截图: Optional[list] = None) -> int:
    """把一次运行写进触发器存储「运行日志」sheet（与前端手动运行同结构），返回行 id。"""
    with _TRIGGER_STORE_LOCK:
        s = _trigger_load()
        rows = s["sheets"].setdefault("运行日志", [])
        rid = s.get("next_id", 1)
        s["next_id"] = rid + 1
        row_no = (max([r.get("row_no") or 0 for r in rows]) + 1) if rows else 1
        data = {
            "应用": 应用, "开始时间": datetime.now().strftime("%Y年%m月%d日 %H:%M"),
            "触发方式": 触发方式, "运行结果": 结果,
            "主题": 主题 or (应用 + " · " + 触发方式),
            "错误次数": 1 if 失败 else 0, "日志": list(日志行), "截图": list(截图 or []),
        }
        rows.append({"id": rid, "row_no": row_no, "data": data})
        _generic_store_save(TRIGGER_STORE_PATH, s)
        return rid


def _trig_notify_failure(应用: str, 主题: str, 日志行: list):
    """运行失败按「运行错误处理」设置发钉钉邮箱通知（收件箱/内容以设置为准）。"""
    try:
        st = (_trigger_load().get("settings") or {})
        if not st.get("邮件通知", True):
            return
        dws = _rpa导dws()
        if dws is None:
            return
        收件 = [x for x in (st.get("收件箱") or []) if str(x).strip()]
        状态, 说明 = dws.发送邮件(
            f"【触发器运行失败】{主题}", "\n".join(日志行),
            收件人=", ".join(收件) if 收件 else None)
        _apd_sync_log(f"失败通知：{状态}（{说明}）")
    except Exception as e:
        _apd_sync_log(f"失败通知发送异常：{e}")


def _trig_artifacts(r: dict) -> tuple:
    """从执行结果里挑出能随通知发出去的产物 → (附件列表, 图片列表)。

    看板推送返回 png/xlsx；维护物料返回 excel；其它应用没有产物就是空。
    路径不存在时由 trigger_notify._路径们 过滤掉，不会发出空文件。
    """
    r = r or {}
    图 = [x for x in (r.get("png"),) if x]
    附 = [x for x in (r.get("xlsx"), r.get("excel"), r.get("附件"), r.get("附件列表")) if x]
    return 附, 图


def _trig_send_notify(d: dict, 显示: str, 触发方式: str, 结果: str, lines: list,
                      附件=None, 图片=None, 主题: str = "", 正文: str = "",
                      用应用正文: bool = False, 跳过: bool = False,
                      钉钉正文: str = "") -> dict:
    """按计划任务 config.通知（开关 + 对象）发一次通知；开关关着就静默跳过。

    对象串可混合通道：计划科（群消息）／王锦志（钉钉消息）／abc@x.com（邮箱）——
    解析与发送都在 app/trigger_notify.py（复用维护物料那套 dws_mail / dws_msg）。
    `附件 / 图片` 是该次运行产出的文件（Excel、长截图），会随通知一起发出去。
    `主题 / 正文` + `用应用正文=True`：通知正文直接用应用产出的内容（如看板日报正文），
      不套「触发器/应用/运行结果」模板。默认 False（其余链路保持原样）。
    `跳过=True`：条件触发器未命中阈值时用——没到阈值本来就不该打扰人，只在运行日志记一行。
    发送结果会写进本次运行的日志行，页面上「运行详情」能直接看到发没发、发到哪。

    ★ 2026-10-26 用户口径「我只配置了钉钉消息，凭什么还收到邮件？」——
      通知是**唯一出口**：配了钉钉就只发钉钉，没配邮箱就绝不发邮件。
        看板推送 / 阈值告警模块不再自作主张发邮件，只产出内容交到这里投递。
    """
    cfg = (d or {}).get("config") or {}
    通知 = cfg.get("通知") or {}
    if 跳过:
        lines.append("[通知] 未达阈值，本次不发送通知（按本任务「通知」配置）")
        return {}
    if not bool(通知.get("启用")):
        return {}
    try:
        import trigger_notify as N
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"[通知] ⚠ 通知模块不可用：{type(e).__name__}: {e}")
        return {}
    try:
        r = N.send(名称=(d or {}).get("名称") or "", 应用=显示 or "", 触发方式=触发方式 or "",
                   运行结果=结果 or "", 日志=list(lines or []), 通知=通知,
                   附件=附件, 图片=图片, 主题=主题 or "", 正文=正文 or "",
                   用应用正文=用应用正文, 钉钉正文=钉钉正文 or "")
    except Exception as e:                                        # noqa: BLE001
        lines.append(f"[通知] ⚠ 发送异常：{type(e).__name__}: {e}")
        return {}
    行 = "[通知] " + str(r.get("汇总") or "未发送")
    for it in (r.get("结果") or []):
        行 += f"｜{it.get('通道')}→{it.get('对象')}：{it.get('状态')}"
    附 = sorted({n for it in (r.get("结果") or []) for n in (it.get("附件") or [])})
    if 附:
        行 += "｜附件 " + "、".join(附)
    lines.append(行)
    _apd_sync_log(f"触发器通知（{(d or {}).get('名称') or 显示}）：{r.get('汇总')}")
    return r


def _trig_notify_pending(r: dict) -> bool:
    """这次通知是不是「只差投递」（dws 授权不可用），需要登记自动补发？

    2026-10-31：看板推送 07:40 撞上 dws 授权抖动（保活线程正在重登），产物全齐、
    只差投递；授权 07:59 自己就恢复了，却没人再去发。用户明确拒绝「手工点立即运行」。
    """
    if not isinstance(r, dict) or not r:
        return False
    if r.get("ok"):
        return False
    串 = f"{r.get('跳过') or ''} {r.get('汇总') or ''}"
    return "授权不可用" in 串


def _trig_patch_log(rid, 追加行: str, ok: bool = True) -> None:
    """给指定运行日志追加一行（补发结果回写，页面「运行详情」能看见补发这件事）。"""
    if not rid:
        return
    try:
        追加行 = str(追加行 or "").strip()
        with _TRIGGER_STORE_LOCK:
            s = _trigger_load()
            for row in (s["sheets"].get("运行日志") or []):
                if row.get("id") != rid:
                    continue
                data = row.setdefault("data", {})
                data["日志"] = list(data.get("日志") or []) + [f"{datetime.now():%H:%M:%S}  {追加行}"]
                if ok:
                    data["运行结果"] = "运行成功（通知已自动补发）"
                    data["错误次数"] = 0
                _generic_store_save(TRIGGER_STORE_PATH, s)
                return
    except Exception as e:                                        # noqa: BLE001
        _apd_sync_log(f"补发回写运行日志失败：{type(e).__name__}: {e}")


def _trig_register_retry(d: dict, 显示: str, 触发方式: str, 结果: str, lines: list,
                         附件=None, 图片=None, 主题: str = "", 正文: str = "",
                         钉钉正文: str = "", rid=None) -> None:
    """把「产物已生成、通知未发出」登记为待补发（授权恢复后由后台线程自动发）。"""
    try:
        import notify_retry as R
        R.登记({
            "名称": (d or {}).get("名称") or 显示,
            "应用": 显示, "触发方式": 触发方式, "结果": 结果,
            "日志": [str(x) for x in (lines or [])][-20:],
            "通知": ((d or {}).get("config") or {}).get("通知") or {},
            "附件": list(附件 or []), "图片": list(图片 or []),
            "主题": 主题 or "", "正文": 正文 or "", "用应用正文": bool(正文),
            "钉钉正文": 钉钉正文 or "", "运行日志id": rid,
        })
    except Exception as e:                                        # noqa: BLE001
        _apd_sync_log(f"登记自动补发失败：{type(e).__name__}: {e}")


def _trig_retry_notify(p: dict) -> tuple:
    """补发处理器：把队列里那条「只差投递」的通知真的发出去，并回写运行日志。

    由 app/notify_retry.py 的后台线程调用；这里不 import 主站以外的东西（回调注册，
    避免 notify_retry 反向依赖 main / RAG 抢 sys.modules['main']）。
    """
    try:
        import trigger_notify as N
        r = N.send(名称=p.get("名称") or "", 应用=p.get("应用") or "", 触发方式=p.get("触发方式") or "",
                   运行结果=p.get("结果") or "", 日志=list(p.get("日志") or []),
                   通知=p.get("通知") or {}, 附件=p.get("附件") or [], 图片=p.get("图片") or [],
                   主题=p.get("主题") or "", 正文=p.get("正文") or "",
                   用应用正文=bool(p.get("用应用正文")), 钉钉正文=p.get("钉钉正文") or "")
    except Exception as e:                                        # noqa: BLE001
        return False, f"{type(e).__name__}: {e}"
    ok = bool(r.get("ok"))
    汇总 = str(r.get("汇总") or "未发送")
    _trig_patch_log(p.get("运行日志id"), f"[通知·补发] {汇总}", ok=ok)
    _apd_sync_log(f"通知补发（{p.get('名称')}）：{汇总}")
    return ok, 汇总



def _trigger_fire_task(task: dict, trigger_label: str, catch_up: bool = False):
    """执行一条定时任务：跑应用 → 写运行日志 → 失败按设置通知。

    catch_up=True：今天已到点但本进程启动时还没跑过（如电脑 07:30 关机、08:00 才开机），
    重启后补跑一次。同步本身幂等（去重键存在即跳过），补跑只会补进真正的新行。

    2026-10-21 增：同时把运行结果写进 MES 三表 SanHaoXiangApdRun_RedoLog
    （运行时间/是否成功/结果说明/各类条数；插入 0 条也如实记录）。"""
    d = task.get("data") or {}
    应用 = d.get("应用") or ""
    名称 = d.get("名称") or 应用
    app = _TRIGGER_APP_REGISTRY.get(应用) or {}
    显示 = app.get("名称") or 应用
    触发方式 = f"[定时触发] {trigger_label}"
    lines = [f"{datetime.now().strftime('%H:%M:%S')}  触发器「{名称}」到点触发（{trigger_label}）"]
    if catch_up:
        lines.append(f"{datetime.now().strftime('%H:%M:%S')}  今日到点后服务才启动，执行补跑"
                     "（同步幂等：已存在的行自动跳过，只会补进真正的新行）")
    lines.append(f"{datetime.now().strftime('%H:%M:%S')}  应用启动：{显示}")
    _run_t0 = datetime.now()
    r = _trig_run_app(应用, d.get("config") or {})
    _run_t1 = datetime.now()
    lines.extend(r.get("日志") or [])
    # 结果口径（2026-10-26）：产物都出了、只差邮件 → 「部分成功」，别一枪打成运行失败
    结果 = "运行成功" if r.get("ok") else ("部分成功" if r.get("部分成功") else "运行失败")
    # 通知（config.通知 开关 + 对象）：成功/失败都发，日志里能看见发到哪几个对象；产物（Excel/长截图）一起带上
    附, 图 = _trig_artifacts(r)
    通知返回 = _trig_send_notify(d, 显示, 触发方式, 结果, lines, 附, 图,
                      主题=r.get("邮件主题") or "", 正文=r.get("邮件正文") or "",
                      用应用正文=bool(r.get("邮件正文")),
                      钉钉正文=r.get("钉钉正文") or "",
                      跳过=(r.get("命中") is False))
    # 2026-10-31 用户「我不要手工授权」：产物齐、只差 dws 投递 → 登记自动补发，
    # 授权一恢复后台就发出去（不再需要谁去点「立即运行」）。
    需补发 = _trig_notify_pending(通知返回)
    if 需补发:
        lines.append("[通知] 已登记「授权恢复后自动补发」（后台每 3 分钟检查，无需手工操作）")
    # 截图记录（2026-10-27 用户反馈「运行日志截图也没有」）：
    # 优先用本次运行的真实产物（看板长截图 / KPI 卡截图）；完全没有产物时，
    # 截一张看板页面当失败现场证据 —— 免得运行详情里空着一块，用户啥也看不到。
    截图记录 = _trig_artifact_shots(r)
    if not 截图记录 and not r.get("ok"):
        截图记录 = _trig_evidence_shot(应用, lines)
    运行日志id = _trig_write_log(显示, 触发方式, 结果, lines, 主题=f"{显示} · {trigger_label}", 失败=not r.get("ok"),
                    截图=截图记录)
    if 需补发:
        _trig_register_retry(d, 显示, 触发方式, 结果, lines, 附件=附, 图片=图,
                             主题=r.get("邮件主题") or "", 正文=r.get("邮件正文") or "",
                             钉钉正文=r.get("钉钉正文") or "", rid=运行日志id)
    _trigger_write_mes_run_log(应用, r, _run_t0, _run_t1, trigger_mode="TIMER")
    if not r.get("ok") and not 需补发:
        # 2026-10-31：「产物齐、只差投递」不算运行失败（已登记自动补发），
        # 否则授权一恢复就会给用户补一封「触发器运行失败」的邮件 —— 07:59 就发过一封，
        # 内容还是 07:43 的旧日志，纯噪音且误导。
        _trig_notify_failure(显示, f"{显示} · {trigger_label}", lines)
    return r


def _trigger_write_mes_run_log(应用: str, r: dict, t0, t1, trigger_mode: str = "TIMER"):
    """把一次运行结果写进 MES 三表日志（SanHaoXiangApdRun_RedoLog）。

    只对注册表里已接入的 APD 同步应用生效（目前 apd_source_sync_5d）。
    写库失败不影响业务主流程，只在服务日志里告警。
    """
    if 应用 != "apd_source_sync_5d":
        return
    try:
        import apd_run_log as _L
        ok = bool(r.get("ok"))
        stat = {
            "inserted": r.get("inserted", 0),
            "skipped_exists": r.get("skipped_exists"),
            "failed_rows": r.get("failed_rows", 0),
            "total_rows": r.get("total_rows"),
            "结果": r.get("结果") or "",
        }
        # skipped 没单独返回时，从日志行里兜底解析「跳过已存在 N」
        if stat["skipped_exists"] is None:
            for ln in (r.get("日志") or []):
                if "跳过已存在" in ln:
                    try:
                        stat["skipped_exists"] = int(ln.split("跳过已存在")[1].split("/")[0].strip())
                    except Exception:
                        pass
        detail = "\n".join((r.get("日志") or [])[-40:])
        wrote, msg = _L.write_run_log(ok=ok, result=stat, trigger_mode=trigger_mode,
                                      start_time=t0, end_time=t1, detail=detail,
                                      operator="scheduler")
        _apd_sync_log(f"运行日志入库{'成功' if wrote else '失败'}：{msg}")
    except Exception as e:  # noqa: BLE001
        _apd_sync_log(f"运行日志入库异常（不影响业务）：{type(e).__name__}: {e}")


def _trigger_scheduler_loop():
    """守护线程：每 30s 扫一次「计划任务」，启用的定时任务到点即真实执行（同一天只跑一次）。"""
    def _go():
        import time as _t
        _t.sleep(30)   # 等服务与预热完成
        _trig_sched_state_load()
        _apd_sync_log("触发器调度器已启动：每 30s 检查一次计划任务到点情况")
        while True:
            _t.sleep(30)
            try:
                s = _trigger_load()
                tasks = (s.get("sheets") or {}).get("计划任务") or []
                if not tasks:
                    continue
                now = datetime.now()
                today = now.strftime("%Y-%m-%d")
                fired = _TRIGGER_SCHED_STATE["fired"]
                for task in tasks:
                    d = task.get("data") or {}
                    # 2026-10-26：新增「条件触发器」（类型=条件）——到点先算指标，命中才通知。
                    # 调度口径与定时完全一致（同一个 _trig_due_datetime），差别只在执行体自己判阈值。
                    if d.get("类型") not in ("定时", "条件") or not d.get("启用"):
                        continue
                    应用 = d.get("应用") or ""
                    if 应用 not in _TRIGGER_APP_REGISTRY:
                        continue        # 仅调度已注册后端执行器的应用（网页端任务由网页触发）
                    cfg = d.get("config") or {}
                    due = _trig_due_datetime(cfg, now)
                    if due is None or now < due:
                        continue
                    key = str(task.get("id"))
                    if fired.get(key, {}).get("last_day") == today:
                        continue        # 今天已跑过（重启后由 sched_state.json 记忆）
                    with _TRIGGER_SCHED_LOCK:
                        if fired.get(key, {}).get("last_day") == today:
                            continue
                        fired[key] = {"last_day": today, "task": d.get("名称") or 应用,
                                      "last_ts": now.strftime("%Y-%m-%d %H:%M:%S")}
                        _trig_sched_state_save()
                    if d.get("类型") == "条件":
                        label = (f"{cfg.get('指标') or '指标'} {cfg.get('比较') or '>'} "
                                 f"{cfg.get('阈值') if cfg.get('阈值') not in (None, '') else ''}"
                                 f"｜{cfg.get('freq')} {cfg.get('time')}")
                    else:
                        label = (f"{cfg.get('freq')} {cfg.get('time')}"
                                 if cfg.get("freq") in ("每天", "每工作日") else
                                 f"{cfg.get('freq')} {cfg.get('time') or ''}".strip())
                    # 到点超过 10 分钟才执行 = 本进程启动后的补跑（正常到点在 30s 扫描精度内）
                    catch_up = (now - due).total_seconds() > 600
                    # 2026-10-26：改为「各自线程执行」——维护物料要等 ECS 的任务进程跑完（几十分钟级），
                    # 串行执行会把同窗口内其它到点任务（如 07:30 的 03_apd 同步）一起顶住。
                    # 「今天已跑过」在起线程之前就落盘（上面 fired[key]），所以不会重复触发。
                    try:
                        threading.Thread(
                            target=_trigger_fire_task, args=(task, label),
                            kwargs={"catch_up": catch_up}, daemon=True,
                            name=f"trig-{task.get('id')}",
                        ).start()
                    except Exception as e:
                        _apd_sync_log(f"触发器「{d.get('名称')}」执行异常：{e}")
            except Exception as e:
                _APD_SYNC["error"] = str(e)
    threading.Thread(target=_go, daemon=True).start()


@app.get("/api/dws/diagnose")
def dws_diagnose(recover: bool = False):
    """钉钉连接诊断（2026-10-27，2026-10-30 增补长期凭据）。

    recover=true 时会顺手尝试一次自动恢复（重新授权）。
    返回 {可用, 原因, 说明, 建议, 剩余分钟, 客户端在跑, profile,
          长期凭据剩余天, 长期凭据到期, 短期令牌到期, 最后登录}。
    注意「剩余分钟」是钉钉固定 2 小时的短期令牌余量（会自动换新），
    「长期凭据剩余天」（refresh token，30 天）才是真会到期的那个。
    """
    try:
        import dws_conn_guard as G
    except Exception as e:                                        # noqa: BLE001
        return {"可用": False, "原因": "未知", "说明": f"模块不可用：{type(e).__name__}: {e}",
                "建议": "重启服务后重试", "剩余分钟": None, "客户端在跑": False, "profile": ""}
    if recover:
        try:
            G.确保可用(强制=True)
        except Exception:
            pass
    d = G.诊断()
    d["ok"] = bool(d.get("可用"))
    return d


@app.get("/api/trigger/apps")
def trigger_apps():
    """可被后端真实执行的应用清单（供网页「立即运行」走真实链路 + 下拉候选项）。"""
    return {"ok": True, "apps": [
        {"id": k, "name": v["名称"], "desc": v["描述"], "run_url": "/api/trigger/app/run"}
        for k, v in _TRIGGER_APP_REGISTRY.items()]}


class TriggerAppRun(BaseModel):
    app: str
    名称: str = ""
    config: dict = {}          # 条件触发器：指标 / 比较 / 阈值 / 收件人


@app.post("/api/trigger/app/run")
def trigger_app_run(payload: TriggerAppRun):
    """网页「立即运行」入口：对注册表内的应用执行真实后端逻辑并返回日志行。"""
    if payload.app not in _TRIGGER_APP_REGISTRY:
        raise HTTPException(status_code=400, detail=f"应用「{payload.app}」未注册后端执行器")
    _t0 = datetime.now()
    r = _trig_run_app(payload.app, payload.config or {})
    _trigger_write_mes_run_log(payload.app, r, _t0, datetime.now(), trigger_mode="MANUAL")
    显示 = (_TRIGGER_APP_REGISTRY.get(payload.app) or {}).get("名称") or payload.app
    行 = list(r.get("日志") or [])
    结果文本 = "运行成功" if r.get("ok") else ("部分成功" if r.get("部分成功") else "运行失败")
    # 手动运行同样按该触发器的「通知」配置发（前端对后端应用不再重复发，见 trigger.js runNow）；产物一起带上
    附, 图 = _trig_artifacts(r)
    _trig_send_notify({"名称": payload.名称 or 显示, "config": payload.config or {}},
                      显示, "[手动运行]", 结果文本, 行, 附, 图,
                      主题=r.get("邮件主题") or "", 正文=r.get("邮件正文") or "",
                      用应用正文=bool(r.get("邮件正文")),
                      钉钉正文=r.get("钉钉正文") or "",
                      跳过=(r.get("命中") is False))
    return {"ok": bool(r.get("ok")), "部分成功": bool(r.get("部分成功")), "结果": r.get("结果"),
            "日志": 行,
            # 真实产物截图（看板长截图 / KPI 卡）：前端直接用它们当「运行完成画面」，
            # 不用再截触发器页面那种没用的图（2026-10-26 用户反馈）。
            "截图": _trig_artifact_shots(r),
            "inserted": r.get("inserted"), "files": r.get("files"),
            "total_rows": r.get("total_rows"), "folders": r.get("folders")}


@app.get("/api/trigger/notify/options")
def trigger_notify_options():
    """触发器「通知对象」候选清单（群消息 / 钉钉消息 / 邮箱）。

    与「维护物料模块自动化」抬头同一个数据源（钉钉邮箱群消息配置模块），
    维护一次，两处候选都跟着变。
    """
    try:
        import box_task_api as _B
        d = _B.notify_options()
        return {"ok": True, "群消息": d.get("群消息") or [], "钉钉消息": d.get("钉钉消息") or [],
                "邮箱": d.get("邮箱") or [], "全部": d.get("全部") or []}
    except Exception as e:                                        # noqa: BLE001
        return {"ok": False, "detail": f"{type(e).__name__}: {e}",
                "群消息": [], "钉钉消息": [], "邮箱": [], "全部": []}


class TriggerNotifyReq(BaseModel):
    """网页侧（前端模拟 / 前端流水线类型）跑完的通知请求。"""
    名称: str = ""
    应用: str = ""
    触发方式: str = "[手动运行]"
    运行结果: str = ""
    日志: list = []
    通知: dict = {}
    附件: list = []          # 该次运行产出的文件（Excel 等）
    图片: list = []          # 截图（邮箱内联 / 钉钉按图片文件发）


@app.post("/api/trigger/run-notify")
def trigger_notify_send(payload: TriggerNotifyReq):
    """按触发器「通知」配置发一次（文件/热键/邮件/钉钉类触发器的运行由网页侧发起，故由网页回调本接口）。

    注意：路径不能叫 /api/trigger/notify —— 那是「运行错误处理 → 发送测试邮件」的老接口（见文件前部）。
    """
    try:
        import trigger_notify as _N
        return _N.send(名称=payload.名称, 应用=payload.应用, 触发方式=payload.触发方式,
                       运行结果=payload.运行结果,
                       日志=[str(x) for x in (payload.日志 or [])],
                       通知=payload.通知 or {},
                       附件=payload.附件 or None, 图片=payload.图片 or None)
    except Exception as e:                                        # noqa: BLE001
        return {"ok": False, "汇总": f"通知发送异常：{type(e).__name__}: {e}", "结果": []}


@app.get("/api/trigger/schedule-state")
def trigger_schedule_state():
    """调度器状态：各定时任务最近一次到点执行记录（排障用）。"""
    return {"ok": True, "fired": _TRIGGER_SCHED_STATE.get("fired", {})}


# ============ 维护物料模块自动化 · 接口（2026-10-24 集成 box_level_task 脚本）============
# 蓝图见 app/box_task_api.py：配置读写 / 运行(子进程) / 状态轮询 / 运行日志(尾N行+超上限自动压缩) / 结果与 Excel 下载。
# 与既有 rag_mount 同款「挂载失败不影响主站」写法。
try:
    from box_task_api import router as _box_router
    app.include_router(_box_router)
except Exception as _box_e:
    print(f"[BOX] 维护物料模块接口挂载异常，已跳过：{type(_box_e).__name__}: {_box_e}")


# ============ 修改模块自动化 · 接口（2026-10-24 移植「批量提交待办」参考脚本） ============
# 蓝图见 app/modify_todo_auto.py：run（同步 def + 专用 daemon 线程跑 playwright）/ status
# （after 游标取增量日志，窗口内不丢行）/ stop（协作式取消，不强杀）。
# 与既有 box_task_api 同款「挂载失败不影响主站」写法。
try:
    from modify_todo_auto import router as _modify_router
    app.include_router(_modify_router)
    # 启动自检：把「当前解释器 + playwright 可用性」打出来。
    # playwright 只装在 C:/Python311，用别的解释器（如托管 3.13）启动本服务时，
    # 路由能挂上但 run 会在点击时报 ModuleNotFoundError——提前暴露，别等用户点按钮才发现。
    import modify_todo_auto as _mod_auto
    if getattr(_mod_auto, "_PLAYWRIGHT_OK", False):
        print(f"[MODIFY] 修改模块自动化接口已挂载；playwright 可用（解释器 {sys.executable}）")
    else:
        # 2026-10-25：不再只喊「请用 Python311 重启」——模块自身会后台自愈安装（见
        # modify_todo_auto._确保playwright），这里只把状态说清楚，并把可用解释器点出来。
        print(f"[MODIFY][警告] 当前解释器 {sys.executable} 暂无 playwright："
              f"{getattr(_mod_auto, '_PLAYWRIGHT_ERR', 'unknown')}——已启动后台自动安装自愈")
        _可用 = _mod_auto._找可用解释器() if hasattr(_mod_auto, "_找可用解释器") else None
        print(f"[MODIFY][警告] 已知装了 playwright 的解释器：{_可用 or '无（自愈安装依赖网络）'}")
    # 登录态保活（2026-10-26）：access_token 只有 2 小时寿命，过期后 refresh 也会一起失效
    # → 只能人工扫码。这里起一个后台线程临期就静默续期，做到「扫一次码管很久」。
    if hasattr(_mod_auto, "start_auth_keepalive"):
        try:
            _mod_auto.start_auth_keepalive()
            print("[MODIFY] 登录态保活线程已启动（临期自动续期，失败只记日志不影响主站）")
        except Exception as _ka_e:
            print(f"[MODIFY] 登录态保活启动失败（忽略）：{type(_ka_e).__name__}: {_ka_e}")
except Exception as _modify_e:
    print(f"[MODIFY] 修改模块自动化接口挂载异常，已跳过：{type(_modify_e).__name__}: {_modify_e}")


# dws（钉钉 CLI）授权保活（2026-10-27，2026-10-30 重写为「定期用一次 API」）：
# 钉钉的短期令牌 2 小时是服务端写死的（谁也改不了），但**过期会自动刷新**；
# 长期凭据 30 天、**用一次就轮转**（= 又续 30 天）。所以让钉钉「永久在线」的正解不是反复 login
# （实测 login 必然弹扫码页、没有静默模式），而是**每 20 分钟调一次最便宜的只读 API**。
def _gbk安全print(文本: str, 前缀: str = ""):
    """控制台 print 的 GBK 兜底：Windows 隐藏控制台是 cp936，⚠/emoji 这类字符会让 print
    直接抛 UnicodeEncodeError —— 2026-11-03 现场：就因为这个异常，[DWS] 保活线程连续 16 代
    服务进程没启动过，令牌无人续期、14:55 掉线。凡启动期 print 一律走这里。"""
    try:
        print(前缀 + 文本)
    except UnicodeEncodeError:
        try:
            print((前缀 + 文本).encode("gbk", errors="replace").decode("gbk"))
        except Exception:
            print(前缀 + str(文本).encode("ascii", errors="replace").decode("ascii"))


try:
    import dws_auth_guard as _dws_guard
    # 2026-10-31 自检：全线只许一个 dws exe（多版本共用凭据会互相把登录态写成「旧版」，
    # 这才是「授权老是掉」的机制性根因，不是钉钉那边的问题）。
    # 2026-11-03 重排：**先启动保活再做诊断打印** —— 之前 ⚠ 打印崩溃会连累保活一起不启动。
    if _dws_guard.启动保活(间隔秒=600, 提前分钟=60, 日志=lambda m: _gbk安全print(m)):
        _gbk安全print("[DWS] 授权保活线程已启动（每 10 分钟用一次钉钉 API，令牌自动滚动续期）")
    _gbk安全print(f"[DWS] 统一入口 exe = {_dws_guard.找dws()}｜版本 {_dws_guard.版本() or '未知'}"
          f"（只允许这一个；smec-harness 那份已停用为 dws.exe.disabled）")
    _残留 = _dws_guard.旧版残留()
    if _残留:
        _gbk安全print(f"[DWS][警告] 发现第二份 dws：{_残留} —— 多版本共用同一份凭据会把登录态写成"
              f"「旧版」，导致授权周期性失效，请把它改名停用")
except Exception as _dws_e:
    _gbk安全print(f"[DWS] 授权保活启动失败（忽略）：{type(_dws_e).__name__}: {_dws_e}")


# 通知补发（2026-10-31，用户「我不要手工授权，你能不能靠谱点」）：
# dws 授权是外部依赖（钉钉客户端掉登录/被顶号时只能等它自己重登），我们控制不了；
# 能控制的是「授权恢复之后，把当时没发出去的那条自动补发掉」。
# 队列落在 app/trigger_store/pending_notify.json（服务重启也还在），后台每 3 分钟检查一次。
try:
    import notify_retry as _notify_retry

    def _补发日志(m: str):
        print(m)
        # 只有「放弃」才值得吵人（授权长期不可用、12 小时都没补上）；
        # 平时的「还在等授权」全部静默，别再往用户钉钉里塞噪音。
        if "放弃" in str(m):
            try:
                _trig_notify_failure("通知自动补发", "通知自动补发失败（dws 授权长期不可用）", [str(m)])
            except Exception:
                pass

    if _notify_retry.启动补发(_trig_retry_notify, 间隔秒=180, 日志=_补发日志):
        print("[补发] 通知补发线程已启动（dws 授权恢复后自动补发，无需手工点「立即运行」）")
except Exception as _nr_e:
    print(f"[补发] 通知补发启动失败（忽略）：{type(_nr_e).__name__}: {_nr_e}")


def _bounce_watch_loop():
    """守护线程：每 20 分钟扫一次收件箱退信，自动重发被 MX 偶发弹回的邮件（纯 DWS，不需授权码）。

    2026-10-28（用户「邮件连续退回」）：退信以 no-reply@mailsupport.aliyun.com 邮件落收件箱，
    其正文含原主题/收件地址/502 原因；扫描→解析→匹配 sent_mail_queue.json→自动重发（同主题上限 2 次）。
    可用环境变量 APS_MAIL_BOUNCE_WATCH=0 关闭。
    """
    import time as _t
    _t.sleep(90)  # 等服务预热
    while True:
        _t.sleep(20 * 60)
        try:
            if os.environ.get("APS_MAIL_BOUNCE_WATCH", "1") != "1":
                continue
            import mail_bounce_watch as _BW
            _BW.检查退信并重发()
        except Exception as _e:  # noqa: BLE001
            try:
                import log_kit as _lk
                _lk.记录错误(f"退信巡检循环异常：{_e}")
            except Exception:
                pass


# ============ 通用批量写 /store/bulk（2026-11-04） ============
# 为什么要它：单元格「拖动填充 / 区域粘贴 / Delete 清空」原先走「每格一次 upsert + 每格整表重载」——
# 27 格 = 27 次往返 + 27 次重绘，肉眼可见地狂闪；慢网下相邻写入还会互相覆盖
# （真机实测：清空 7 格，后端只落下 5 格 → 用户看到「删了、又好像没删干净、还撤不回来」）。
# 现在统一成一次请求、一次落盘、行数不变：前端只刷变动的那几格，一次操作也只产生一个撤销项。
# 此前只有 /api/mese/limiter 有 bulk，其它族缺这条路由 → 那些页面的 Delete 清空/批量填充
# 一律 404（页面上表现为「清空失败」toast）。这里把各族补齐。


def _bulk_apply(path, sheet: str, updates: list) -> int:
    """把 [{id, data}] 批量写进某 store 的某 sheet：只改 data，不动 id / row_no / 行数。"""
    if not sheet or not isinstance(updates, list) or not updates:
        return 0
    s = _generic_store_load(path)
    rows = s.setdefault("sheets", {}).setdefault(sheet, [])
    by_id = {r.get("id"): r for r in rows}
    n = 0
    for u in updates[:5000]:
        r = by_id.get((u or {}).get("id"))
        data = (u or {}).get("data")
        if r is None or not isinstance(data, dict):
            continue
        data = _停产供方_清洗(sheet, data)   # 只「停产供方」生效：去供应商代号前导 0
        # 2026-11-04：合并而不是整块替换 —— 撤销补丁只带「改过的那几列」，
        # 整块替换会把没带上的列全清空（大表撤销的致命细节）。
        cur = dict(r.get("data") or {})
        cur.update(data)
        r["data"] = cur
        n += 1
    if n:
        _generic_store_save(path, s)
    return n


def _bulk_body(payload: dict):
    p = payload or {}
    return str(p.get("sheet") or ""), (p.get("updates") or [])


@app.post("/api/forecast/store/bulk")
def forecast_store_bulk(payload: dict):
    """预投工作表：批量改多行 data（拖动填充/区域粘贴/Delete 清空）。body: {sheet, updates:[{id,data}]}"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(FORECAST_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"预投工作表批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/safety/store/bulk")
def safety_store_bulk(payload: dict):
    """安全钳基础数据：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(SAFETY_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"安全钳本地存储批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/modify/store/bulk")
def modify_store_bulk(payload: dict):
    """修改模块：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(MODIFY_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"修改模块批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/perm/store/bulk")
def perm_store_bulk(payload: dict):
    """权限管控：批量改多行 data。密码列前端是掩码且只读（readonlyCols），这里再兜一层：
    带回来的密码值若是掩码/空，一律沿用库里原哈希，绝不把掩码当密码写回去。"""
    sheet, updates = _bulk_body(payload)
    if not sheet or not isinstance(updates, list) or not updates:
        return {"ok": True, "count": 0}
    s = _generic_store_load(PERM_STORE_PATH)
    rows = s.setdefault("sheets", {}).setdefault(sheet, [])
    by_id = {r.get("id"): r for r in rows}
    n = 0
    for u in updates[:5000]:
        r = by_id.get((u or {}).get("id"))
        data = (u or {}).get("data")
        if r is None or not isinstance(data, dict):
            continue
        old = r.get("data") or {}
        data = dict(data)
        if "密码" in old and (not str(data.get("密码") or "").strip() or str(data.get("密码")) == str(old.get("密码"))):
            data["密码"] = old.get("密码")
        r["data"] = data
        n += 1
    if n:
        _generic_store_save(PERM_STORE_PATH, s)
        UI信息(f"权限管控批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/loginlog/store/bulk")
def loginlog_store_bulk(payload: dict):
    """登录日志：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(LOGINLOG_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"登录日志批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/ddconfig/store/bulk")
def ddconfig_store_bulk(payload: dict):
    """钉钉邮箱配置：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(DDCONFIG_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"钉钉邮箱配置批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/trigger/store/bulk")
def trigger_store_bulk(payload: dict):
    """触发器配置：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(TRIGGER_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"触发器配置批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


@app.post("/api/mese/safety/store/bulk")
def mese_safety_store_bulk(payload: dict):
    """MESE安全钳汇总：批量改多行 data（同上）。"""
    sheet, updates = _bulk_body(payload)
    n = _bulk_apply(MESE_SAFETY_STORE_PATH, sheet, updates)
    if n:
        UI信息(f"MESE安全钳汇总批量修改 {sheet}：{n} 行")
    return {"ok": True, "count": n}


threading.Thread(target=_trigger_scheduler_loop, daemon=True).start()
threading.Thread(target=_bounce_watch_loop, daemon=True).start()


if __name__ == "__main__":
    import socket
    import uvicorn

    端口 = int(os.environ.get("PORT", 8000))

    def _端口可用(端口号: int) -> bool:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("0.0.0.0", 端口号))
            return True
        except OSError:
            return False
        finally:
            s.close()

    if not _端口可用(端口):
        print(f"[错误] 端口 {端口} 已被其他程序占用。")
        print(f"        换一个端口启动：set PORT=8010 && python main.py")
        sys.exit(1)

    print(f"SAP 表格数据处理工具（网页版）启动中 → http://127.0.0.1:{端口}")
    uvicorn.run(app, host="0.0.0.0", port=端口)
