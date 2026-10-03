# -*- coding: utf-8 -*-
"""
PPI高维宏观试提数据：多来源自动抓取 V2.0

本程序严格按照《PPI试提清单与提取记录模板_V1.0.xlsx》组织：
1 个预测目标 + 33 个候选变量 = 34 条论文试提序列。

设计原则：
- AKShare 优先抓取已有稳定接口的中国宏观数据；
- FRED 直接抓取美国工业生产、VIX、Brent、国际铜；
- BIS SDMX API 直接抓取中国 REER；
- Wind / 中国债券信息网 / 国家统计局细项数据等暂时不强行猜接口，
  程序会保留空列并生成“待人工/备用来源”清单，避免把错误序列当成正确数据；
- 不制造“假实时”数据：程序记录本次抓取时间，但不会把当前修订历史值
  冒充成历史 vintage；
- 高频数据的月均值、月末值同时保存；正式预测时再根据预测截止日筛选可用信息；
- 不再粗暴生成统一“滞后1期”列，因为不同指标的实际发布日期不同，统一滞后1期
  会掩盖真实的信息可得性。

运行前：
    pip install -U akshare pandas numpy openpyxl requests

运行：
    python PPI高维试提_多源自动抓取_V2.0.py

输出目录：
    ./PPI_试提_V2_输出/
"""

from __future__ import annotations

import json
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import numpy as np
import pandas as pd
import requests

# =========================
# 基本配置
# =========================
START_DATE = "2000-01"
OUTPUT_DIR = Path(__file__).resolve().parent / "PPI_试提_V2_输出"
RAW_DIR = OUTPUT_DIR / "01_原始"
PROC_DIR = OUTPUT_DIR / "02_处理后"
LOG_DIR = OUTPUT_DIR / "03_日志"
TIMEOUT = 30

# True：允许用 AKShare 的现货备用接口尝试提取钢材/LNG等。
# False：严格按模板主来源（Wind）处理，避免来源口径悄悄变化。
USE_AKSHARE_SPOT_FALLBACK = False

# =========================
# 34条目标序列：严格对应 V1.0
# =========================
TARGETS = [
    (1, "PPI（预测目标）"),
    (2, "CPI"),
    (3, "PPIRM"),
    (4, "PPIRM—燃料、动力类"),
    (5, "工业增加值"),
    (6, "粗钢产量"),
    (7, "发电量"),
    (8, "社会消费品零售总额"),
    (9, "服务业生产指数"),
    (10, "固定资产投资"),
    (11, "制造业投资"),
    (12, "房地产开发投资"),
    (13, "出口金额"),
    (14, "铁矿石进口数量"),
    (15, "美元兑人民币汇率"),
    (16, "人民币实际有效汇率（REER）"),
    (17, "M1"),
    (18, "M2"),
    (19, "DR007"),
    (20, "10年期国债收益率"),
    (21, "3个月国债收益率"),
    (22, "制造业PMI"),
    (23, "PMI新订单指数"),
    (24, "PMI主要原材料购进价格指数"),
    (25, "美国工业生产"),
    (26, "VIX恐慌指数"),
    (27, "工业企业产成品存货"),
    (28, "建筑业商务活动指数"),
    (29, "螺纹钢价格"),
    (30, "国内铜现货价格"),
    (31, "LNG价格"),
    (32, "Brent原油价格"),
    (33, "国际铜价"),
    (34, "波罗的海干散货指数（BDI）"),
]

TARGET_COLUMNS = [x[1] for x in TARGETS]


# =========================
# 通用工具
# =========================
def ensure_dirs() -> None:
    for p in (OUTPUT_DIR, RAW_DIR, PROC_DIR, LOG_DIR):
        p.mkdir(parents=True, exist_ok=True)


def save_raw(name: str, df: pd.DataFrame) -> None:
    safe = (
        name.replace("/", "_")
        .replace("\\", "_")
        .replace(":", "_")
        .replace("?", "_")
        .replace("*", "_")
    )
    df.to_csv(RAW_DIR / f"{safe}_原始.csv", index=False, encoding="utf-8-sig")


def save_processed(name: str, df: pd.DataFrame) -> None:
    safe = name.replace("/", "_").replace("\\", "_")
    df.to_csv(PROC_DIR / f"{safe}_处理后.csv", index=False, encoding="utf-8-sig")


def parse_numeric(x: Any) -> float:
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "")
    if s in {"", "--", "-", "None", "nan", "NaN"}:
        return np.nan
    # 例如 "-2.40%"
    s = s.replace("%", "")
    try:
        return float(s)
    except ValueError:
        return np.nan


def parse_month_series(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()

    # 优先处理 YYYY-MM
    out = pd.to_datetime(s, format="%Y-%m", errors="coerce")

    # 再处理 YYYYMM
    mask = out.isna()
    if mask.any():
        out.loc[mask] = pd.to_datetime(s.loc[mask], format="%Y%m", errors="coerce")

    # 再处理类似 2020年07月份 / 2020年7月
    mask = out.isna()
    if mask.any():
        cleaned = (
            s.loc[mask]
            .str.replace("年", "-", regex=False)
            .str.replace("月份", "", regex=False)
            .str.replace("月", "", regex=False)
        )
        out.loc[mask] = pd.to_datetime(cleaned, errors="coerce")

    return out.dt.to_period("M").astype(str)


def find_col(df: pd.DataFrame, candidates: Iterable[str], required: bool = True) -> str | None:
    cols = [str(c).strip() for c in df.columns]
    mapping = {c: c for c in cols}

    for cand in candidates:
        if cand in mapping:
            return mapping[cand]

    # 模糊匹配
    for cand in candidates:
        for col in cols:
            if cand in col:
                return mapping[col]

    if required:
        raise KeyError(f"没有找到字段。候选字段：{list(candidates)}；实际字段：{cols}")
    return None


def standard_monthly(
    df: pd.DataFrame,
    date_candidates: Iterable[str],
    value_candidates: Iterable[str],
    output_name: str,
) -> pd.DataFrame:
    date_col = find_col(df, date_candidates)
    value_col = find_col(df, value_candidates)

    out = df[[date_col, value_col]].copy()
    out["月份"] = parse_month_series(out[date_col])
    out[output_name] = out[value_col].map(parse_numeric)
    out = out[["月份", output_name]].dropna(subset=["月份"])
    out = out.sort_values("月份").drop_duplicates("月份", keep="last")
    return out.reset_index(drop=True)


def daily_to_monthly(
    df: pd.DataFrame,
    date_candidates: Iterable[str],
    value_candidates: Iterable[str],
    output_name: str,
) -> pd.DataFrame:
    date_col = find_col(df, date_candidates)
    value_col = find_col(df, value_candidates)

    out = df[[date_col, value_col]].copy()
    out["日期"] = pd.to_datetime(out[date_col], errors="coerce")
    out[output_name] = out[value_col].map(parse_numeric)
    out = out.dropna(subset=["日期"])
    out["月份"] = out["日期"].dt.to_period("M").astype(str)

    month_mean = out.groupby("月份")[output_name].mean().rename(f"{output_name}月均")
    month_end = (
        out.sort_values("日期")
        .groupby("月份")[output_name]
        .last()
        .rename(f"{output_name}月末")
    )

    result = pd.concat([month_mean, month_end], axis=1).reset_index()
    return result.sort_values("月份").reset_index(drop=True)


def make_single_primary(df: pd.DataFrame, primary_col: str, keep_cols: list[str] | None = None) -> pd.DataFrame:
    if keep_cols is None:
        keep_cols = ["月份", primary_col]
    result = df[keep_cols].copy()
    result = result.sort_values("月份").drop_duplicates("月份", keep="last")
    return result.reset_index(drop=True)


# =========================
# 运行日志
# =========================
STATUS: list[dict[str, Any]] = []


def record_status(
    name: str,
    status: str,
    source: str,
    interface: str = "",
    start: str = "",
    end: str = "",
    missing_rate: float | None = None,
    note: str = "",
) -> None:
    STATUS.append(
        {
            "序号": next((i for i, n in TARGETS if n == name), None),
            "变量": name,
            "状态": status,
            "来源": source,
            "接口": interface,
            "实际起始期": start,
            "实际结束期": end,
            "缺失率": None if missing_rate is None else round(missing_rate, 6),
            "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "备注": note,
        }
    )


def try_job(
    name: str,
    source: str,
    interface_names: list[str],
    loader: Callable[[Any, str], pd.DataFrame],
) -> pd.DataFrame | None:
    try:
        import akshare as ak
    except ImportError:
        record_status(name, "环境缺少依赖", source, note="未安装 akshare")
        return None

    fn_name = next((x for x in interface_names if hasattr(ak, x)), None)
    if fn_name is None:
        record_status(
            name,
            "接口不可用",
            source,
            interface=" / ".join(interface_names),
            note="当前AKShare版本未发现候选接口；程序未猜测替代序列。",
        )
        return None

    try:
        df = loader(ak, fn_name)
        if df is None or df.empty:
            raise ValueError("接口返回为空")
        save_raw(name, df)
        record_status(
            name,
            "成功",
            source,
            interface=fn_name,
            note=f"原始数据已保存：{name}_原始.csv",
        )
        return df
    except Exception as exc:
        record_status(
            name,
            "抓取失败",
            source,
            interface=fn_name,
            note=f"{type(exc).__name__}: {exc}",
        )
        return None


# =========================
# 中国数据：AKShare
# =========================
def fetch_ppi() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        p = standard_monthly(
            df,
            ["月份", "日期"],
            ["当月同比增长", "当月同比", "同比增长"],
            "PPI同比",
        )
        index_col = find_col(df, ["当月", "指数"], required=False)
        if index_col:
            idx = df[[find_col(df, ["月份", "日期"]), index_col]].copy()
            idx["月份"] = parse_month_series(idx.iloc[:, 0])
            idx["PPI原始指数"] = idx.iloc[:, 1].map(parse_numeric)
            p = p.merge(idx[["月份", "PPI原始指数"]], on="月份", how="left")
        return p

    return try_job("PPI（预测目标）", "国家统计局/AKShare", ["macro_china_ppi"], loader)


def fetch_cpi() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["全国-同比增长", "当月同比", "全国同比", "同比增长"],
            "CPI",
        )

    return try_job("CPI", "国家统计局/AKShare", ["macro_china_cpi"], loader)


def fetch_ppirm() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)(indicator="购进价格指数")
        mask = df.astype(str).apply(lambda col: col.str.contains("工业生产者购进价格指数|总计|综合", na=False)).any(axis=1)
        sub = df.loc[mask].copy()
        if sub.empty:
            sub = df.copy()
        return standard_monthly(
            sub,
            ["月份", "日期"],
            ["当月同比增长", "当月同比", "同比增长"],
            "PPIRM",
        )

    return try_job(
        "PPIRM",
        "国家统计局/AKShare",
        ["macro_china_ppi_category", "macro_china_purchasing_price_index"],
        loader,
    )


def fetch_ppirm_fuel() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)(indicator="购进价格指数")
        name_col = find_col(df, ["指标名称", "项目", "分类"], required=False)
        if name_col:
            df = df[df[name_col].astype(str).str.contains("燃料、动力类", na=False)].copy()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["当月同比增长", "当月同比", "同比增长"],
            "PPIRM燃料动力类同比",
        )

    return try_job(
        "PPIRM—燃料、动力类",
        "国家统计局/AKShare",
        ["macro_china_ppi_category"],
        loader,
    )


def fetch_industrial_added() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "同比", "value", "值"],
            "工业增加值",
        )

    return try_job(
        "工业增加值",
        "国家统计局/AKShare",
        ["macro_china_gyzjz", "macro_china_industrial_production_yoy", "macro_china_industrial_added_value"],
        loader,
    )


def fetch_products_output() -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    def loader(ak, fn_name):
        return getattr(ak, fn_name)()

    df = try_job(
        "粗钢产量",
        "国家统计局/AKShare",
        ["macro_china_main_products_output"],
        loader,
    )
    if df is None:
        # 尝试另一个常见命名；记录在两个变量各自结果中
        return None, None

    product_col = find_col(df, ["产品名称", "产品", "指标名称"], required=False)
    value_col = find_col(df, ["产量", "数值", "当月"], required=False)
    date_col = find_col(df, ["月份", "日期"], required=False)
    if not all([product_col, value_col, date_col]):
        return None, None

    df = df.copy()
    steel = df[df[product_col].astype(str).str.contains("粗钢", na=False)].copy()
    power = df[df[product_col].astype(str).str.contains("发电量", na=False)].copy()

    out_steel = standard_monthly(steel, [date_col], [value_col], "粗钢产量") if not steel.empty else None
    out_power = standard_monthly(power, [date_col], [value_col], "发电量") if not power.empty else None

    if out_steel is not None:
        record_status("粗钢产量", "成功", "国家统计局/AKShare", "macro_china_main_products_output")
    if out_power is not None:
        record_status("发电量", "成功", "国家统计局/AKShare", "macro_china_main_products_output")
    return out_steel, out_power


def fetch_retail() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "当月同比", "全国同比"],
            "社会消费品零售总额",
        )

    return try_job(
        "社会消费品零售总额",
        "国家统计局/AKShare",
        ["macro_china_consumer_goods_retail", "macro_china_retail_sales_yoy"],
        loader,
    )


def fetch_service_index() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "同比", "服务业生产指数"],
            "服务业生产指数",
        )

    return try_job(
        "服务业生产指数",
        "国家统计局/AKShare",
        ["macro_china_service_production_index"],
        loader,
    )


def fetch_fai() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "当月同比", "同比"],
            "固定资产投资",
        )

    return try_job("固定资产投资", "国家统计局/AKShare", ["macro_china_gdzctz"], loader)


def fetch_manufacturing_investment() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        # 优先寻找同比字段；如果接口返回累计值，则不擅自做1-2月伪拆分
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "同比", "制造业投资同比"],
            "制造业投资",
        )

    return try_job(
        "制造业投资",
        "国家统计局/AKShare",
        ["macro_china_manufacturing_investment", "macro_china_mfg_investment"],
        loader,
    )


def fetch_real_estate_investment() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["月份", "日期"],
            ["同比增长", "同比", "房地产开发投资同比"],
            "房地产开发投资",
        )

    return try_job(
        "房地产开发投资",
        "国家统计局/AKShare",
        ["macro_china_real_estate_development_investment", "macro_china_real_estate_investment"],
        loader,
    )


def fetch_exports() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["日期", "月份", "date"],
            ["同比", "数值", "value"],
            "出口金额",
        )

    return try_job("出口金额", "海关/AKShare", ["macro_china_exports_yoy"], loader)


def fetch_rmb() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        return standard_monthly(
            df,
            ["日期", "月份"],
            ["美元/人民币_中间价", "美元兑人民币中间价"],
            "美元兑人民币汇率",
        )

    return try_job("美元兑人民币汇率", "国家外汇管理局/AKShare", ["macro_china_rmb"], loader)


def fetch_money_supply() -> pd.DataFrame | None:
    def loader(ak, fn_name):
        df = getattr(ak, fn_name)()
        date_col = find_col(df, ["月份", "日期"])
        cols = list(df.columns)
        m1_col = find_col(df, ["货币(M1)同比增长", "M1同比", "M1同比增长"], required=False)
        m2_col = find_col(df, ["货币和准货币(M2)同比增长", "M2同比", "M2同比增长"], required=False)
        if not m1_col or not m2_col:
            raise KeyError(f"M1/M2同比字段未找到，实际字段：{cols}")
        out = df[[date_col, m1_col, m2_col]].copy()
        out["月份"] = parse_month_series(out[date_col])
        out["M1"] = out[m1_col].map(parse_numeric)
        out["M2"] = out[m2_col].map(parse_numeric)
        return out[["月份", "M1", "M2"]].sort_values("月份").drop_duplicates("月份", keep="last")

    return try_job("M1", "中国人民银行/AKShare", ["macro_china_money_supply"], loader)


# =========================
# FRED
# =========================
def fetch_fred_series(symbol: str) -> pd.DataFrame:
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={symbol}"
    r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": "Mozilla/5.0 PPI-thesis-data-collector"})
    r.raise_for_status()
    from io import StringIO
    df = pd.read_csv(StringIO(r.text))
    return df


def add_fred_monthly(
    name: str,
    symbol: str,
    out_col: str,
    mode: str,
) -> pd.DataFrame | None:
    try:
        df = fetch_fred_series(symbol)
        save_raw(name, df)

        date_col = df.columns[0]
        value_col = symbol if symbol in df.columns else df.columns[1]
        d = df[[date_col, value_col]].copy()
        d["日期"] = pd.to_datetime(d[date_col], errors="coerce")
        d[value_col] = pd.to_numeric(d[value_col], errors="coerce")
        d = d.dropna(subset=["日期"])

        if mode == "monthly_index_yoy":
            d = d.sort_values("日期")
            d["值"] = d[value_col]
            d["月份"] = d["日期"].dt.to_period("M")
            m = d.groupby("月份")["值"].last().sort_index()
            yoy = m.pct_change(12) * 100
            result = pd.DataFrame({"月份": yoy.index.astype(str), out_col: yoy.values}).dropna()

        elif mode == "daily_month_mean":
            d["月份"] = d["日期"].dt.to_period("M").astype(str)
            result = d.groupby("月份")[value_col].mean().rename(out_col).reset_index()

        elif mode == "monthly_level":
            d["月份"] = d["日期"].dt.to_period("M").astype(str)
            result = d.groupby("月份")[value_col].last().rename(out_col).reset_index()

        else:
            raise ValueError(f"未知FRED处理模式：{mode}")

        record_status(
            name,
            "成功",
            "FRED",
            f"{symbol}",
            result["月份"].min() if not result.empty else "",
            result["月份"].max() if not result.empty else "",
            result[out_col].isna().mean() if not result.empty else None,
            "FRED当前历史序列；不是历史vintage。下载时间已记录。",
        )
        save_processed(name, result)
        return result

    except Exception as exc:
        record_status(name, "抓取失败", "FRED", symbol, note=f"{type(exc).__name__}: {exc}")
        return None


# =========================
# BIS REER
# =========================
def fetch_bis_reer() -> pd.DataFrame | None:
    url = (
        "https://stats.bis.org/api/v1/data/WS_EER/M.R.B.CN/all"
        "?startPeriod=1994-01&endPeriod=2100-12"
    )
    try:
        r = requests.get(
            url,
            timeout=TIMEOUT,
            headers={"User-Agent": "Mozilla/5.0 PPI-thesis-data-collector"},
        )
        r.raise_for_status()

        # BIS 默认可返回 SDMX-XML；这里用本地XML解析，避免依赖额外库。
        import xml.etree.ElementTree as ET

        root = ET.fromstring(r.text)
        obs = []
        for elem in root.iter():
            if elem.tag.endswith("Obs"):
                period = elem.attrib.get("TIME_PERIOD")
                value = elem.attrib.get("OBS_VALUE")
                if period is not None and value is not None:
                    obs.append((period, parse_numeric(value)))

        if not obs:
            raise ValueError("BIS返回中没有解析到OBS观测值")

        out = pd.DataFrame(obs, columns=["日期", "人民币实际有效汇率"])
        out["日期"] = pd.to_datetime(out["日期"], errors="coerce")
        out["月份"] = out["日期"].dt.to_period("M").astype(str)
        out = out[["月份", "人民币实际有效汇率"]].dropna()
        out = out.drop_duplicates("月份", keep="last").sort_values("月份")

        save_raw("人民币实际有效汇率（REER）", pd.DataFrame(obs, columns=["日期", "值"]))
        save_processed("人民币实际有效汇率（REER）", out)
        record_status(
            "人民币实际有效汇率（REER）",
            "成功",
            "BIS",
            "WS_EER / M.R.B.CN",
            out["月份"].min(),
            out["月份"].max(),
            out["人民币实际有效汇率"].isna().mean(),
            "中国广义篮子实际有效汇率；Index 2020=100。",
        )
        return out

    except Exception as exc:
        record_status(
            "人民币实际有效汇率（REER）",
            "抓取失败",
            "BIS",
            "WS_EER / M.R.B.CN",
            note=f"{type(exc).__name__}: {exc}",
        )
        return None


# =========================
# 保守占位：严格来源未验证时不猜
# =========================
def register_manual(name: str, source: str, note: str) -> None:
    record_status(name, "待人工/备用来源", source, note=note)


# =========================
# 主流程
# =========================
def main() -> None:
    ensure_dirs()
    start_run = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("=" * 72)
    print("PPI高维试提数据：多来源自动抓取 V2.0")
    print("目标：1条PPI + 33条候选变量 = 34条")
    print("抓取时间：", start_run)
    print("=" * 72)

    try:
        import akshare as ak
        print("AKShare版本：", getattr(ak, "__version__", "未知"))
    except ImportError:
        print("未安装AKShare。请执行：pip install -U akshare")
        sys.exit(1)

    # -------------------------
    # 1. 中国宏观
    # -------------------------
    ppi = fetch_ppi()
    cpi = fetch_cpi()
    ppirm = fetch_ppirm()
    ppirm_fuel = fetch_ppirm_fuel()
    industrial = fetch_industrial_added()
    steel, electricity = fetch_products_output()
    retail = fetch_retail()
    service = fetch_service_index()
    fai = fetch_fai()
    manufacturing_invest = fetch_manufacturing_investment()
    real_estate = fetch_real_estate_investment()
    exports = fetch_exports()
    rmb = fetch_rmb()
    money = fetch_money_supply()

    # 当前V1.0中仍需要明确口径/来源，不用“猜接口”替代
    register_manual(
        "铁矿石进口数量",
        "海关总署",
        "需要按海关口径核实进口数量序列；程序不把其他铁矿石指标冒充该变量。",
    )
    register_manual(
        "DR007",
        "中国货币网/Wind",
        "需要使用实际DR007历史序列；不以SHIBOR等近似替代。",
    )
    register_manual(
        "10年期国债收益率",
        "中国债券信息网",
        "当前公开AKShare收盘收益率曲线接口存在历史窗口限制，不适合直接拼1990年代以来长期样本；本程序不伪造长期历史。",
    )
    register_manual(
        "3个月国债收益率",
        "中国债券信息网",
        "同上；需要正式长期序列或Wind提取。",
    )
    register_manual(
        "PMI新订单指数",
        "国家统计局",
        "当前AKShare官方PMI接口能稳定取得制造业PMI，但本程序不猜细项字段；建议从NBS细项表正式提取。",
    )
    register_manual(
        "PMI主要原材料购进价格指数",
        "国家统计局",
        "同上；NBS已发布该细项，但需要细项历史序列的稳定提取口径。",
    )
    register_manual(
        "制造业投资",
        "国家统计局",
        "若当前AKShare版本没有可核实的月度当月/同比字段，转官方表或Wind；不从错误累计口径硬拆。",
    )
    register_manual(
        "房地产开发投资",
        "国家统计局",
        "同上；正式序列需要确认当月/累计口径以及1—2月统计期。",
    )
    register_manual(
        "工业企业产成品存货",
        "国家统计局/工业企业经济效益",
        "发布滞后明显；按实际发布日期处理。若接口字段不稳定，转官方表。",
    )
    register_manual(
        "螺纹钢价格",
        "Wind",
        "主序列按模板固定为Wind统一市场现货/价格序列；不把任意现货报价当正式论文序列。",
    )
    register_manual(
        "国内铜现货价格",
        "Wind",
        "同上；需要Wind具体代码并核对单位、历史起点。",
    )
    register_manual(
        "LNG价格",
        "Wind",
        "同上；需要Wind具体代码并核对单位、历史起点。",
    )
    register_manual(
        "波罗的海干散货指数（BDI）",
        "Baltic Exchange",
        "正式来源按Baltic Exchange固定；不使用未经授权的替代网站值直接写入正式面板。",
    )

    # -------------------------
    # 2. FRED
    # -------------------------
    us_ip = add_fred_monthly("美国工业生产", "INDPRO", "美国工业生产", "monthly_index_yoy")
    vix = add_fred_monthly("VIX恐慌指数", "VIXCLS", "VIX恐慌指数", "daily_month_mean")
    brent = add_fred_monthly("Brent原油价格", "DCOILBRENTEU", "Brent原油价格", "daily_month_mean")
    international_copper = add_fred_monthly("国际铜价", "PCOPPUSDM", "国际铜价", "monthly_level")

    # -------------------------
    # 3. BIS REER
    # -------------------------
    reer = fetch_bis_reer()

    # -------------------------
    # 4. 构建34列试提面板
    # -------------------------
    series_map: dict[str, pd.DataFrame | None] = {
        "PPI（预测目标）": ppi,
        "CPI": cpi,
        "PPIRM": ppirm,
        "PPIRM—燃料、动力类": ppirm_fuel,
        "工业增加值": industrial,
        "粗钢产量": steel,
        "发电量": electricity,
        "社会消费品零售总额": retail,
        "服务业生产指数": service,
        "固定资产投资": fai,
        "制造业投资": manufacturing_invest,
        "房地产开发投资": real_estate,
        "出口金额": exports,
        "铁矿石进口数量": None,
        "美元兑人民币汇率": rmb,
        "人民币实际有效汇率（REER）": reer,
        "M1": money[["月份", "M1"]] if money is not None and "M1" in money.columns else None,
        "M2": money[["月份", "M2"]] if money is not None and "M2" in money.columns else None,
        "DR007": None,
        "10年期国债收益率": None,
        "3个月国债收益率": None,
        "制造业PMI": None,
        "PMI新订单指数": None,
        "PMI主要原材料购进价格指数": None,
        "美国工业生产": us_ip,
        "VIX恐慌指数": vix,
        "工业企业产成品存货": None,
        "建筑业商务活动指数": None,
        "螺纹钢价格": None,
        "国内铜现货价格": None,
        "LNG价格": None,
        "Brent原油价格": brent,
        "国际铜价": international_copper,
        "波罗的海干散货指数（BDI）": None,
    }

    # 为“制造业PMI”单独抓取官方PMI总指数；不混同细项
    try:
        def pmi_loader(ak, fn_name):
            df = getattr(ak, fn_name)()
            col = find_col(df, ["制造业-指数", "制造业PMI", "value", "数值"])
            date = find_col(df, ["月份", "日期"])
            return standard_monthly(df, [date], [col], "制造业PMI")

        pmi = try_job("制造业PMI", "国家统计局/AKShare", ["macro_china_pmi"], pmi_loader)
        series_map["制造业PMI"] = pmi
    except Exception:
        pass

    # 建筑业商务活动指数尝试读取非制造业PMI；只有明确命中“建筑业”才使用
    try:
        def construction_loader(ak, fn_name):
            df = getattr(ak, fn_name)()
            name_col = find_col(df, ["指标名称", "项目", "分类"], required=False)
            if name_col:
                sub = df[df[name_col].astype(str).str.contains("建筑业", na=False)].copy()
            else:
                sub = df.copy()
            if sub.empty:
                raise ValueError("未明确识别建筑业商务活动指数")
            return standard_monthly(sub, ["月份", "日期"], ["数值", "指数", "商务活动指数"], "建筑业商务活动指数")

        construction = try_job(
            "建筑业商务活动指数",
            "国家统计局/AKShare",
            ["macro_china_non_man_pmi"],
            construction_loader,
        )
        if construction is not None:
            series_map["建筑业商务活动指数"] = construction
    except Exception:
        pass

    # 可选：仅作为试提备用，不替代Wind主序列
    if USE_AKSHARE_SPOT_FALLBACK:
        register_manual(
            "螺纹钢价格",
            "AKShare备用",
            "仅允许在确认具体品种与单位后用于交叉核对；正式主序列仍按Wind。",
        )
        register_manual(
            "国内铜现货价格",
            "AKShare备用",
            "仅允许在确认具体品种与单位后用于交叉核对；正式主序列仍按Wind。",
        )
        register_manual(
            "LNG价格",
            "AKShare备用",
            "仅允许在确认具体品种与单位后用于交叉核对；正式主序列仍按Wind。",
        )

    # 统一月份：把所有成功序列转成长表再合并
    prepared = []
    for name in TARGET_COLUMNS:
        df = series_map.get(name)
        if df is None or df.empty:
            prepared.append(pd.DataFrame(columns=["月份", name]))
            continue
        if "月份" not in df.columns:
            continue
        cols = [c for c in df.columns if c != "月份"]
        if not cols:
            continue
        value_col = cols[0]
        temp = df[["月份", value_col]].copy()
        if value_col != name:
            temp = temp.rename(columns={value_col: name})
        temp["月份"] = temp["月份"].astype(str)
        temp[name] = temp[name].map(parse_numeric)
        prepared.append(temp)

    panel = pd.DataFrame(columns=["月份"] + TARGET_COLUMNS)
    for temp in prepared:
        if temp.empty:
            continue
        if panel.empty:
            panel = temp
        else:
            panel = pd.merge(panel, temp, on="月份", how="outer")

    panel = panel.sort_values("月份").drop_duplicates("月份", keep="last").reset_index(drop=True)

    # 限定起始月份
    panel = panel[panel["月份"] >= START_DATE].reset_index(drop=True)

    # 质量报告
    quality_rows = []
    for name in TARGET_COLUMNS:
        s = panel[name] if name in panel.columns else pd.Series(dtype=float)
        quality_rows.append(
            {
                "变量": name,
                "列是否存在": name in panel.columns,
                "实际起始期": panel.loc[s.notna(), "月份"].min() if name in panel.columns and s.notna().any() else "",
                "实际结束期": panel.loc[s.notna(), "月份"].max() if name in panel.columns and s.notna().any() else "",
                "观测数": int(s.notna().sum()) if name in panel.columns else 0,
                "缺失数": int(s.isna().sum()) if name in panel.columns else 0,
                "缺失率": float(s.isna().mean()) if name in panel.columns and len(s) else np.nan,
            }
        )

    quality_df = pd.DataFrame(quality_rows)
    status_df = pd.DataFrame(STATUS).drop_duplicates(
        subset=["变量", "状态", "接口", "来源"], keep="last"
    )

    # 保存
    run_tag = datetime.now().strftime("%Y%m%d_%H%M%S")
    panel_path = OUTPUT_DIR / f"PPI试提月度面板_{run_tag}.xlsx"
    quality_path = OUTPUT_DIR / f"PPI试提质量报告_{run_tag}.xlsx"
    status_path = OUTPUT_DIR / f"PPI试提抓取状态_{run_tag}.xlsx"
    meta_path = LOG_DIR / f"运行元数据_{run_tag}.json"

    panel.to_excel(panel_path, index=False)
    quality_df.to_excel(quality_path, index=False)
    status_df.to_excel(status_path, index=False)

    metadata = {
        "程序": "PPI高维试提_多源自动抓取_V2.0.py",
        "运行时间": run_tag,
        "数据起始筛选": START_DATE,
        "目标序列数": len(TARGET_COLUMNS),
        "说明": "当前历史序列+抓取时间记录；不宣称为历史vintage实时数据。",
        "面板文件": str(panel_path),
        "质量报告": str(quality_path),
        "抓取状态": str(status_path),
    }
    meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + "=" * 72)
    print("抓取流程结束")
    print("面板：", panel_path)
    print("质量：", quality_path)
    print("状态：", status_path)
    print("原始数据目录：", RAW_DIR)
    print("本轮共定义：", len(TARGET_COLUMNS), "条论文序列")
    print("实际面板列数（不含月份）：", max(0, len(panel.columns) - 1))
    print("=" * 72)
    print("注意：待人工/备用来源并不等于失败数据，而是程序刻意避免猜错口径。")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户中断。")
        sys.exit(130)
    except Exception as exc:
        ensure_dirs()
        err_path = LOG_DIR / f"严重错误_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        err_path.write_text(traceback.format_exc(), encoding="utf-8")
        print("\n程序发生未处理错误：", exc)
        print("错误日志：", err_path)
        sys.exit(1)
