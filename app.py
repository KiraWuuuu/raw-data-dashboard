import io
import re
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


# =========================
# 页面设置
# =========================
st.set_page_config(page_title="Prada Tracking Builder", layout="wide")
st.title("Prada Tracking Builder")
st.caption("上传 Preview Raw Data + Tracking Template + Mapping Config，自动生成输出 Excel。")


# =========================
# 标准字段
# =========================
STANDARD_COLUMNS = [
    "Media", "Position", "Market", "Date", "Landing page", "Campaign",
    "Cost", "IMP", "CLICK", "CPM", "CPC", "CTR",
    "ENG", "Like", "Forward", "Comment", "Revenue", "Orders",
    "source_file", "target_sheet"
]


# =========================
# 工具函数
# =========================
def norm_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    return s


def clean_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def safe_get_col(df: pd.DataFrame, target_name: str) -> Optional[pd.Series]:
    """
    忽略大小写、空格、下划线、连字符差异来取列
    """
    target_norm = norm_text(target_name)
    for col in df.columns:
        if norm_text(col) == target_norm:
            return df[col]
    return None


def to_number(v):
    if pd.isna(v) or v == "":
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)

    s = str(v).strip().replace(",", "").replace("¥", "").replace("￥", "").replace("$", "")
    if s.endswith("%"):
        try:
            return float(s[:-1]) / 100
        except Exception:
            return np.nan

    try:
        return float(s)
    except Exception:
        return np.nan


def safe_div(a, b):
    if pd.isna(a) or pd.isna(b) or b == 0:
        return np.nan
    return a / b


def allocate_cost_by_imp(df: pd.DataFrame, total_cost) -> pd.DataFrame:
    """
    如果 raw 中没有 Cost，用 config 中配置的总成本按 IMP 比例分摊
    """
    df = df.copy()
    total_cost = to_number(total_cost)

    if pd.isna(total_cost):
        return df

    imp_sum = df["IMP"].sum(skipna=True)
    if pd.isna(imp_sum) or imp_sum == 0:
        df["Cost"] = np.nan
        return df

    df["Cost"] = df["IMP"] / imp_sum * float(total_cost)
    return df


def add_kpis(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "CTR" not in df.columns:
        df["CTR"] = np.nan
    if "CPM" not in df.columns:
        df["CPM"] = np.nan
    if "CPC" not in df.columns:
        df["CPC"] = np.nan

    df["CTR"] = df["CTR"].where(df["CTR"].notna(), df.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1))
    df["CPM"] = df["CPM"].where(df["CPM"].notna(), df.apply(lambda x: safe_div(x["Cost"] * 1000, x["IMP"]), axis=1))
    df["CPC"] = df["CPC"].where(df["CPC"].notna(), df.apply(lambda x: safe_div(x["Cost"], x["CLICK"]), axis=1))

    return df


def map_market_by_keyword(text: str, market_map_df: pd.DataFrame) -> str:
    txt = str(text)
    for _, row in market_map_df.iterrows():
        keyword = str(row["keyword"])
        bucket = str(row["market_bucket"])
        if keyword and keyword in txt:
            return bucket
    return "Unknown"


def read_excel_first_sheet(uploaded_file) -> pd.DataFrame:
    """
    当前版按第一个 sheet 读取：
    - 微信 / 微博第一张就是目标数据
    - 抖音 feeds / opening 第一张也是 *_Net 数据
    """
    data = uploaded_file.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(data), sheet_name=first_sheet, engine="openpyxl")
    return clean_colnames(df)


def load_mapping_config(uploaded_file):
    cfg = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
    if "file_mapping" not in cfg:
        raise ValueError("mapping_config.xlsx 缺少 sheet：file_mapping")
    if "market_mapping" not in cfg:
        raise ValueError("mapping_config.xlsx 缺少 sheet：market_mapping")

    file_map = clean_colnames(cfg["file_mapping"])
    market_map = clean_colnames(cfg["market_mapping"])

    required_cols = [
        "source_file", "target_sheet", "media_name",
        "position_mode", "position_value",
        "landing_page_default",
        "market_mode", "market_value",
        "cost_mode", "cost_total"
    ]
    missing = [c for c in required_cols if c not in file_map.columns]
    if missing:
        raise ValueError(f"file_mapping 缺少字段：{missing}")

    return file_map, market_map


def find_rule(file_map_df: pd.DataFrame, source_file_name: str) -> pd.Series:
    matched = file_map_df[file_map_df["source_file"].astype(str) == str(source_file_name)]
    if matched.empty:
        raise ValueError(f"mapping_config 的 file_mapping 中找不到 source_file = {source_file_name}")
    return matched.iloc[0]


# =========================
# 各媒体解析器
# =========================
def process_wechat(df: pd.DataFrame, cfg: pd.Series, market_map_df: pd.DataFrame) -> pd.DataFrame:
    """
    适配：
    Prada may issue wechat brand gallery.xlsx
    原始字段：
    日期 / 广告名称 / 花费 / 曝光次数 / 点击次数 / 点击率 / 点赞次数 / 分享次数 / 评论次数 / 下单金额 / 下单次数
    """
    out = pd.DataFrame()

    date_col = safe_get_col(df, "日期")
    campaign_col = safe_get_col(df, "广告名称")
    cost_col = safe_get_col(df, "花费")
    imp_col = safe_get_col(df, "曝光次数")
    click_col = safe_get_col(df, "点击次数")
    ctr_col = safe_get_col(df, "点击率")
    like_col = safe_get_col(df, "点赞次数")
    forward_col = safe_get_col(df, "分享次数")
    comment_col = safe_get_col(df, "评论次数")
    revenue_col = safe_get_col(df, "下单金额")
    orders_col = safe_get_col(df, "下单次数")

    out["Date"] = date_col
    out["Campaign"] = campaign_col
    out["Cost"] = cost_col.apply(to_number) if cost_col is not None else np.nan
    out["IMP"] = imp_col.apply(to_number) if imp_col is not None else np.nan
    out["CLICK"] = click_col.apply(to_number) if click_col is not None else np.nan
    out["CTR"] = ctr_col.apply(to_number) if ctr_col is not None else np.nan

    out["Like"] = like_col.apply(to_number) if like_col is not None else np.nan
    out["Forward"] = forward_col.apply(to_number) if forward_col is not None else np.nan
    out["Comment"] = comment_col.apply(to_number) if comment_col is not None else np.nan

    out["Revenue"] = revenue_col.apply(to_number) if revenue_col is not None else np.nan
    out["Orders"] = orders_col.apply(to_number) if orders_col is not None else np.nan

    out["ENG"] = out[["Like", "Forward", "Comment"]].fillna(0).sum(axis=1)

    out["Media"] = cfg["media_name"]

    # Position
    if str(cfg["position_mode"]).lower() == "fixed":
        out["Position"] = cfg["position_value"]
    else:
        raw_pos = safe_get_col(df, str(cfg["position_value"]))
        out["Position"] = raw_pos if raw_pos is not None else "Unknown"

    # Landing page
    out["Landing page"] = cfg["landing_page_default"]

    # Market
    mmode = str(cfg["market_mode"]).lower()
    if mmode == "keyword":
        out["Market"] = out["Campaign"].apply(lambda x: map_market_by_keyword(x, market_map_df))
    elif mmode == "fixed":
        out["Market"] = cfg["market_value"]
    elif mmode == "raw":
        raw_market = safe_get_col(df, str(cfg["market_value"]))
        out["Market"] = raw_market if raw_market is not None else "Unknown"
    else:
        out["Market"] = "Unknown"

    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = add_kpis(out)

    for c in STANDARD_COLUMNS:
        if c not in out.columns:
            out[c] = np.nan

    return out[STANDARD_COLUMNS]


def process_douyin(df: pd.DataFrame, cfg: pd.Series) -> pd.DataFrame:
    """
    适配：
    Prada may issue douyin feeds.xlsx
    Prada may issue douyin opening.xlsx

    原始字段：
    Region / Type / SPID / Website / Channel / Ad Placement / Campaign ID / CampaignName / Date / Impression / Click / CTR
    """
    out = pd.DataFrame()

    date_col = safe_get_col(df, "Date")
    camp_col = safe_get_col(df, "CampaignName")
    imp_col = safe_get_col(df, "Impression")
    click_col = safe_get_col(df, "Click")
    ctr_col = safe_get_col(df, "CTR")

    out["Date"] = date_col
    out["Campaign"] = camp_col
    out["IMP"] = imp_col.apply(to_number) if imp_col is not None else np.nan
    out["CLICK"] = click_col.apply(to_number) if click_col is not None else np.nan
    out["CTR"] = ctr_col.apply(to_number) if ctr_col is not None else np.nan

    out["Media"] = cfg["media_name"]

    # Position
    if str(cfg["position_mode"]).lower() == "fixed":
        out["Position"] = cfg["position_value"]
    else:
        raw_pos = safe_get_col(df, str(cfg["position_value"]))
        out["Position"] = raw_pos if raw_pos is not None else "Unknown"

    # Market
    mmode = str(cfg["market_mode"]).lower()
    if mmode == "raw":
        raw_market = safe_get_col(df, str(cfg["market_value"]))
        out["Market"] = raw_market if raw_market is not None else "Unknown"
    elif mmode == "fixed":
        out["Market"] = cfg["market_value"]
    else:
        out["Market"] = "Unknown"

    out["Landing page"] = cfg["landing_page_default"]

    # 这些字段 raw 没有
    out["Like"] = np.nan
    out["Forward"] = np.nan
    out["Comment"] = np.nan
    out["ENG"] = np.nan
    out["Revenue"] = np.nan
    out["Orders"] = np.nan

    # Cost
    out["Cost"] = np.nan
    if str(cfg["cost_mode"]).lower() == "raw":
        raw_cost = safe_get_col(df, "Cost")
        if raw_cost is not None:
            out["Cost"] = raw_cost.apply(to_number)
    elif str(cfg["cost_mode"]).lower() == "allocate_total":
        out = allocate_cost_by_imp(out, cfg["cost_total"])

    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = add_kpis(out)

    for c in STANDARD_COLUMNS:
        if c not in out.columns:
            out[c] = np.nan

    return out[STANDARD_COLUMNS]


def process_weibo(df: pd.DataFrame, cfg: pd.Series) -> pd.DataFrame:
    """
    适配：
    Prada may issue weibo opening.xlsx

    原始字段：
    日期 / 点位 / PV / Click
    """
    out = pd.DataFrame()

    date_col = safe_get_col(df, "日期")
    pos_col = safe_get_col(df, "点位")
    imp_col = safe_get_col(df, "PV")
    click_col = safe_get_col(df, "Click")

    out["Date"] = date_col
    out["Position"] = pos_col if pos_col is not None else "Unknown"
    out["IMP"] = imp_col.apply(to_number) if imp_col is not None else np.nan
    out["CLICK"] = click_col.apply(to_number) if click_col is not None else np.nan
    out["CTR"] = out.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1)

    out["Media"] = cfg["media_name"]

    # Market
    mmode = str(cfg["market_mode"]).lower()
    if mmode == "fixed":
        out["Market"] = cfg["market_value"]
    elif mmode == "raw":
        raw_market = safe_get_col(df, str(cfg["market_value"]))
        out["Market"] = raw_market if raw_market is not None else "Unknown"
    else:
        out["Market"] = "Unknown"

    out["Landing page"] = cfg["landing_page_default"]
    out["Campaign"] = np.nan

    # 微博没有这些
    out["Like"] = np.nan
    out["Forward"] = np.nan
    out["Comment"] = np.nan
    out["ENG"] = np.nan
    out["Revenue"] = np.nan
    out["Orders"] = np.nan

    out["Cost"] = np.nan
    if str(cfg["cost_mode"]).lower() == "allocate_total":
        out = allocate_cost_by_imp(out, cfg["cost_total"])

    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = add_kpis(out)

    for c in STANDARD_COLUMNS:
        if c not in out.columns:
            out[c] = np.nan

    return out[STANDARD_COLUMNS]


# =========================
# Campaign 汇总页
# =========================
def build_campaign_table(all_std: pd.DataFrame) -> pd.DataFrame:
    df = all_std.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    group_cols = ["Media", "Position", "Market", "Date"]
    value_cols = ["Cost", "IMP", "CLICK", "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"]

    agg = df.groupby(group_cols, dropna=False)[value_cols].sum(min_count=1).reset_index()

    agg["CTR"] = agg.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1)
    agg["CPM"] = agg.apply(lambda x: safe_div(x["Cost"] * 1000, x["IMP"]), axis=1)
    agg["CPC"] = agg.apply(lambda x: safe_div(x["Cost"], x["CLICK"]), axis=1)

    agg = agg[[
        "Media", "Position", "Market", "Date",
        "Cost", "IMP", "CLICK", "CPM", "CPC", "CTR",
        "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"
    ]]
    return agg


# =========================
# 写入模板
# =========================
def clear_sheet(ws):
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None


def write_table(ws, df: pd.DataFrame):
    clear_sheet(ws)

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    bold_font = Font(bold=True)

    # header
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = bold_font

    # body
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    # 简单列宽
    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        max_len = 0
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 50)


def build_output_workbook(template_file, all_std: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(template_file.getvalue()))

    # 1) 写 detail sheets
    for target_sheet, sub_df in all_std.groupby("target_sheet", dropna=False):
        target_sheet = str(target_sheet)
        if target_sheet not in wb.sheetnames:
            continue

        ws = wb[target_sheet]
        export_df = sub_df.copy()
        export_df["Date"] = pd.to_datetime(export_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

        # 微信 detail sheet
        if target_sheet in ["【WeChat Moments】", "【WeChat Banner】"]:
            out = export_df[
                ["Market", "Date", "Campaign", "Cost", "IMP", "CLICK", "CTR", "Like", "Forward", "Comment", "Revenue", "Orders"]
            ].copy()
            out.columns = [
                "Market", "日期", "广告名称", "花费", "曝光次数", "点击次数", "点击率",
                "点赞次数", "分享次数", "评论次数", "下单金额", "下单次数"
            ]
        else:
            # Douyin / Weibo detail sheet
            out = export_df[
                ["Market", "Position", "Campaign", "Date", "IMP", "CLICK", "CTR"]
            ].copy()
            out.columns = [
                "Region", "Ad Placement", "CampaignName", "Date", "Impression", "Click", "CTR"
            ]

        write_table(ws, out)

    # 2) 写 Campaign 汇总页
    if "Campaign" in wb.sheetnames:
        ws_campaign = wb["Campaign"]
        campaign_df = build_campaign_table(all_std)
        write_table(ws_campaign, campaign_df)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================
# 页面上传区
# =========================
raw_files = st.file_uploader(
    "Upload Raw Data",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

template_file = st.file_uploader(
    "Upload Template",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)

config_file = st.file_uploader(
    "Upload Mapping Config",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)


# =========================
# 主按钮
# =========================
if st.button("Generate Dashboard"):
    if not raw_files:
        st.error("请先上传 Preview Raw Data。")
        st.stop()

    if template_file is None:
        st.error("请先上传 Tracking Template。")
        st.stop()

    if config_file is None:
        st.error("请先上传 Mapping Config。")
        st.stop()

    try:
        file_map_df, market_map_df = load_mapping_config(config_file)
    except Exception as e:
        st.error(f"读取 mapping_config 失败：{e}")
        st.stop()

    all_parts = []

    for f in raw_files:
        try:
            cfg = find_rule(file_map_df, f.name)
            raw_df = read_excel_first_sheet(f)
            lower_name = f.name.lower()

            if "wechat" in lower_name or "微信" in lower_name:
                std_df = process_wechat(raw_df, cfg, market_map_df)
            elif "douyin" in lower_name and "feeds" in lower_name:
                std_df = process_douyin(raw_df, cfg)
            elif "douyin" in lower_name and "opening" in lower_name:
                std_df = process_douyin(raw_df, cfg)
            elif "weibo" in lower_name or "微博" in lower_name:
                std_df = process_weibo(raw_df, cfg)
            else:
                st.warning(f"文件 {f.name} 没有匹配到规则，已跳过。")
                continue

            all_parts.append(std_df)

        except Exception as e:
            st.error(f"处理文件 {f.name} 失败：{e}")
            st.stop()

    if not all_parts:
        st.error("没有生成可用数据。")
        st.stop()

    all_std = pd.concat(all_parts, ignore_index=True)

    # 预览
    st.subheader("Standardized Preview")
    preview_df = all_std.copy()
    preview_df["Date"] = pd.to_datetime(preview_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    st.dataframe(preview_df, use_container_width=True)

    try:
        output_bytes = build_output_workbook(template_file, all_std)
    except Exception as e:
        st.error(f"写入模板失败：{e}")
        st.stop()

    st.success("生成成功，可以下载输出文件。")
    st.download_button(
        "Download Output Excel",
        data=output_bytes,
        file_name="Prada_tracking_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
