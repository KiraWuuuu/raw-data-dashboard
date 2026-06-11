import io
import re
from typing import List

import numpy as np
import pandas as pd
import streamlit as st


# =========================
# 页面设置
# =========================
st.set_page_config(page_title="Raw Data Dashboard Generator", layout="wide")

TARGET_COLUMNS = [
    "Media",
    "Position",
    "Market",
    "Date",
    "Landing page",
    "Campaign",
    "Cost",
    "IMP",
    "CLICK",
    "CTR",
    "CPM",
    "CPC",
    "ENG",
    "Like",
    "Forward",
    "Comment",
    "Revenue",
    "Orders",
]

NUMERIC_COLUMNS = [
    "Cost", "IMP", "CLICK", "CTR", "CPM", "CPC",
    "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"
]


# =========================
# 字段映射（最终版：适配微信 / 抖音 / Bilibili）
# =========================
COLUMN_MAP = {
    # ===== 微信中文字段 =====
    "日期": "Date",
    "广告名称": "Campaign",
    "花费": "Cost",
    "曝光次数": "IMP",
    "点击次数": "CLICK",
    "点击率": "CTR",
    "点赞次数": "Like",
    "分享次数": "Forward",
    "评论次数": "Comment",
    "下单金额": "Revenue",
    "下单次数": "Orders",

    # ===== 常见中英文字段 =====
    "媒体": "Media",
    "平台": "Media",
    "媒体名称": "Media",
    "渠道": "Media",
    "publisher": "Media",
    "media": "Media",
    "platform": "Media",
    "channel": "Media",

    "广告位": "Position",
    "版位": "Position",
    "position": "Position",
    "placement": "Position",
    "ad placement": "Position",
    "ad placem": "Position",
    "type": "Position",

    "地区": "Market",
    "城市": "Market",
    "market": "Market",
    "region": "Market",

    "落地页": "Landing page",
    "landing page": "Landing page",
    "landing_page": "Landing page",
    "landingpage": "Landing page",
    "website": "Landing page",
    "url": "Landing page",

    "campaign": "Campaign",
    "campaignn": "Campaign",
    "campaign name": "Campaign",
    "campaignid": "Campaign",
    "campaign id": "Campaign",
    "广告活动": "Campaign",

    "date": "Date",
    "day": "Date",
    "时间": "Date",
    "投放日期": "Date",

    "cost": "Cost",
    "spend": "Cost",
    "消耗": "Cost",
    "actual spend": "Cost",
    "media cost": "Cost",

    "imp": "IMP",
    "impression": "IMP",
    "impressions": "IMP",
    "曝光": "IMP",
    "曝光量": "IMP",

    "click": "CLICK",
    "clicks": "CLICK",
    "点击": "CLICK",
    "点击量": "CLICK",

    "ctr": "CTR",
    "cpm": "CPM",
    "cpc": "CPC",

    "eng": "ENG",
    "engagement": "ENG",
    "互动": "ENG",
    "互动量": "ENG",

    "like": "Like",
    "likes": "Like",
    "点赞": "Like",

    "forward": "Forward",
    "share": "Forward",
    "shares": "Forward",
    "转发": "Forward",

    "comment": "Comment",
    "comments": "Comment",
    "评论": "Comment",

    "revenue": "Revenue",
    "gmv": "Revenue",
    "成交金额": "Revenue",
    "销售额": "Revenue",

    "orders": "Orders",
    "order": "Orders",
    "订单": "Orders",
    "订单量": "Orders",
}


# =========================
# 工具函数
# =========================
def norm_text(v: str) -> str:
    """标准化列名文本"""
    if v is None:
        return ""
    v = str(v).strip().lower()
    v = v.replace("_", " ").replace("-", " ")
    v = re.sub(r"\s+", " ", v)
    return v


def map_column(col: str) -> str:
    return COLUMN_MAP.get(norm_text(col), col)


def detect_media(source_file: str, columns: List[str], df: pd.DataFrame) -> str:
    """
    自动识别媒体
    优先级：
    1. 文件名判断
    2. 列结构判断
    """
    name = str(source_file).lower()

    # 1) 先看文件名
    if "wechat" in name or "微信" in name:
        return "WeChat"
    if "douyin" in name or "抖音" in name:
        return "Douyin"
    if "bili" in name or "b站" in name or "bilibili" in name:
        return "Bilibili"

    cols_raw = [str(c) for c in columns]
    cols = [norm_text(c) for c in columns]

    # 2) 微信特征
    if "广告名称" in cols_raw or "下单金额" in cols_raw or "曝光次数" in cols_raw:
        return "WeChat"

    # 3) B站特征：有 Website / Channel / SPID
    if "website" in cols or "channel" in cols or "spid" in cols:
        return "Bilibili"

    # 4) 抖音特征：Region + Type + Date + Impression + Click
    if "region" in cols and "type" in cols and "impression" in cols and "click" in cols:
        return "Douyin"

    return "Unknown"


def to_number(v):
    """把文本/百分比/货币转成数值"""
    if pd.isna(v) or v == "":
        return np.nan

    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)

    s = str(v).strip()
    s = s.replace(",", "").replace("¥", "").replace("￥", "").replace("$", "")

    # 百分比
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


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    """重命名列，并补齐目标字段"""
    df = df.copy()
    df.columns = [map_column(c) for c in df.columns]

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan

    keep_cols = TARGET_COLUMNS + [c for c in df.columns if c not in TARGET_COLUMNS]
    df = df[keep_cols]
    return df


def read_uploaded_file(uploaded_file) -> List[pd.DataFrame]:
    """读取单个上传文件（支持多个sheet）"""
    suffix = uploaded_file.name.lower().split(".")[-1]
    results = []

    if suffix == "csv":
        df = pd.read_csv(uploaded_file)
        df["source_file"] = uploaded_file.name
        df["source_sheet"] = "csv"
        results.append(df)
    else:
        excel_bytes = uploaded_file.getvalue()
        excel_io = io.BytesIO(excel_bytes)
        sheets = pd.read_excel(excel_io, sheet_name=None, engine="openpyxl")

        for sheet_name, df in sheets.items():
            if df is None or df.empty:
                continue
            df["source_file"] = uploaded_file.name
            df["source_sheet"] = sheet_name
            results.append(df)

    return results


def clean_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    核心清洗逻辑：
    1. 统一字段
    2. 媒体识别
    3. 日期/数值标准化
    4. 自动计算 CTR / CPM / CPC
    """
    original_columns = list(raw_df.columns)
    source_file = (
        raw_df["source_file"].iloc[0]
        if "source_file" in raw_df.columns and len(raw_df) > 0
        else ""
    )

    df = ensure_columns(raw_df)

    # 媒体识别
    detected_media = detect_media(source_file, original_columns, raw_df)

    # 如果 Media 列原本为空，则填识别值
    if df["Media"].isna().all() or (df["Media"].astype(str).str.strip() == "").all():
        df["Media"] = detected_media
    else:
        df["Media"] = df["Media"].replace("", np.nan)
        df["Media"] = df["Media"].fillna(detected_media)

    # 强制按文件特征归类媒体（避免 B站 / 抖音被误判成 Channel）
    if detected_media == "Bilibili":
        df["Media"] = "Bilibili"
    elif detected_media == "Douyin":
        df["Media"] = "Douyin"
    elif detected_media == "WeChat":
        df["Media"] = "WeChat"

    # 日期标准化
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # 数值标准化
    for col in NUMERIC_COLUMNS:
        df[col] = df[col].apply(to_number)

    # 自动计算派生字段
    df["CTR"] = df["CTR"].where(df["CTR"].notna(), df.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1))
    df["CPM"] = df["CPM"].where(df["CPM"].notna(), df.apply(lambda x: safe_div(x["Cost"] * 1000, x["IMP"]), axis=1))
    df["CPC"] = df["CPC"].where(df["CPC"].notna(), df.apply(lambda x: safe_div(x["Cost"], x["CLICK"]), axis=1))

    # 如果 ENG 全空，则用 Like + Forward + Comment 合成
    if df["ENG"].isna().all():
        df["ENG"] = (
            df["Like"].fillna(0) +
            df["Forward"].fillna(0) +
            df["Comment"].fillna(0)
        )

    return df


def build_dashboard_tables(df: pd.DataFrame):
    """构建 Sheet1 Dashboard 数据"""
    total_cost = df["Cost"].sum(skipna=True)
    total_imp = df["IMP"].sum(skipna=True)
    total_click = df["CLICK"].sum(skipna=True)
    total_eng = df["ENG"].sum(skipna=True)
    total_revenue = df["Revenue"].sum(skipna=True)
    total_orders = df["Orders"].sum(skipna=True)

    summary = pd.DataFrame([{
        "Total Cost": total_cost,
        "Total IMP": total_imp,
        "Total CLICK": total_click,
        "Overall CTR": safe_div(total_click, total_imp),
        "Total ENG": total_eng,
        "Total Revenue": total_revenue,
        "Total Orders": total_orders,
        "Overall CPM": safe_div(total_cost * 1000, total_imp),
        "Overall CPC": safe_div(total_cost, total_click),
    }])

    by_media = df.groupby("Media", dropna=False)[
        ["Cost", "IMP", "CLICK", "Revenue", "Orders", "ENG", "Like", "Forward", "Comment"]
    ].sum(min_count=1).reset_index()

    by_market = df.groupby("Market", dropna=False)[
        ["Cost", "IMP", "CLICK", "Revenue", "Orders"]
    ].sum(min_count=1).reset_index()

    by_date_tmp = df.copy()
    by_date_tmp["Date"] = by_date_tmp["Date"].dt.strftime("%Y-%m-%d")
    by_date = by_date_tmp.groupby("Date", dropna=False)[
        ["Cost", "IMP", "CLICK", "Revenue", "Orders"]
    ].sum(min_count=1).reset_index()

    return summary, by_media, by_market, by_date


def export_workbook(df: pd.DataFrame) -> bytes:
    """
    导出 Excel：
    - Sheet1_Dashboard
    - 后续 sheet 按 Media 拆分
    """
    from openpyxl.styles import Font, PatternFill, Alignment

    output = io.BytesIO()

    export_df = df.copy()
    export_df["Date"] = export_df["Date"].dt.strftime("%Y-%m-%d")

    summary, by_media, by_market, by_date = build_dashboard_tables(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dashboard_sheet = "Sheet1_Dashboard"

        # Sheet1
        summary.to_excel(writer, sheet_name=dashboard_sheet, index=False, startrow=0)
        by_media.to_excel(writer, sheet_name=dashboard_sheet, index=False, startrow=4)
        by_market.to_excel(writer, sheet_name=dashboard_sheet, index=False, startrow=4 + len(by_media) + 3)
        by_date.to_excel(writer, sheet_name=dashboard_sheet, index=False, startrow=4 + len(by_media) + len(by_market) + 6)

        ws = writer.sheets[dashboard_sheet]
        ws["A1"] = "Dashboard Summary"
        ws["A5"] = "By Media"
        ws[f"A{5 + len(by_media) + 3}"] = "By Market"
        ws[f"A{5 + len(by_media) + len(by_market) + 6}"] = "By Date"

        title_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        sub_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        title_font = Font(color="FFFFFF", bold=True)
        bold_font = Font(bold=True)

        for cell_ref in ["A1", "A5", f"A{5 + len(by_media) + 3}", f"A{5 + len(by_media) + len(by_market) + 6}"]:
            ws[cell_ref].fill = title_fill
            ws[cell_ref].font = title_font

        # 让所有单元格垂直居中
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")

        # 后续 sheet：按 Media 拆分
        for media_name, sub_df in export_df.groupby("Media", dropna=False):
            sheet_name = str(media_name if pd.notna(media_name) else "Unknown")
            sheet_name = sheet_name.replace("/", "_")[:31]

            media_df = sub_df[TARGET_COLUMNS + ["source_file", "source_sheet"]].copy()
            media_df.to_excel(writer, sheet_name=sheet_name, index=False)

            media_ws = writer.sheets[sheet_name]
            for cell in media_ws[1]:
                cell.fill = sub_fill
                cell.font = bold_font

    output.seek(0)
    return output.getvalue()


def format_num(v, pct=False):
    if pd.isna(v):
        return "-"
    if pct:
        return f"{v:.2%}"
    return f"{v:,.2f}"


# =========================
# 页面样式
# =========================
st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}
.hero {
    background: linear-gradient(135deg, #0F4C81 0%, #2A9FD6 100%);
    padding: 24px 28px;
    border-radius: 18px;
    color: white;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='hero'>
  <h1 style='margin:0'>Raw Data Dashboard Generator</h1>
  <p style='margin:8px 0 0 0'>
    上传多个不同媒体的 raw data Excel，自动按目标字段标准化，生成 Dashboard，并导出 Excel（Sheet1 为 Dashboard，后续 Sheet 按 Media 拆分明细）。
  </p>
</div>
""", unsafe_allow_html=True)

with st.expander("目标字段（网站会统一映射到这套字段）", expanded=False):
    st.code(", ".join(TARGET_COLUMNS))

uploaded_files = st.file_uploader(
    "上传多个 Excel / CSV 文件",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="支持多个文件；Excel 会自动读取全部 sheet。"
)

if uploaded_files:
    raw_parts = []

    for uf in uploaded_files:
        try:
            parts = read_uploaded_file(uf)
            raw_parts.extend(parts)
        except Exception as e:
            st.error(f"读取文件失败：{uf.name}，错误：{e}")

    if raw_parts:
        cleaned_parts = [clean_data(part) for part in raw_parts]
        clean_df = pd.concat(cleaned_parts, ignore_index=True)

        summary, by_media, by_market, by_date = build_dashboard_tables(clean_df)

        # KPI
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.metric("Total Cost", format_num(summary.loc[0, "Total Cost"]))
        with c2:
            st.metric("Total IMP", format_num(summary.loc[0, "Total IMP"]))
        with c3:
            st.metric("Total CLICK", format_num(summary.loc[0, "Total CLICK"]))
        with c4:
            st.metric("Overall CTR", format_num(summary.loc[0, "Overall CTR"], pct=True))
        with c5:
            st.metric("Revenue", format_num(summary.loc[0, "Total Revenue"]))
        with c6:
            st.metric("Orders", format_num(summary.loc[0, "Total Orders"]))

        tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "By Media", "By Market", "Preview Data"])

        with tab1:
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Spend by Media")
                if not by_media.empty:
                    st.bar_chart(by_media.set_index("Media")[["Cost"]])

            with col2:
                st.subheader("Trend by Date")
                date_plot = by_date.copy()
                if not date_plot.empty and "Date" in date_plot.columns:
                    date_plot = date_plot[date_plot["Date"].notna()]
                    if not date_plot.empty:
                        st.line_chart(date_plot.set_index("Date")[["Cost", "IMP", "CLICK"]])
                    else:
                        st.info("没有可用于趋势图的日期数据。")

            st.subheader("Dashboard Tables")
            st.markdown("**Summary**")
            st.dataframe(summary, use_container_width=True)

            st.markdown("**By Media**")
            st.dataframe(by_media, use_container_width=True)

            st.markdown("**By Market**")
            st.dataframe(by_market, use_container_width=True)

        with tab2:
            st.subheader("By Media")
            st.dataframe(by_media, use_container_width=True)

            media_options = sorted(by_media["Media"].dropna().astype(str).unique().tolist()) if not by_media.empty else []
            if media_options:
                selected_media = st.selectbox("选择 Media 查看对应明细", media_options)
                media_df = clean_df[clean_df["Media"].astype(str) == selected_media].copy()
                media_df["Date"] = media_df["Date"].dt.strftime("%Y-%m-%d")
                st.dataframe(media_df[TARGET_COLUMNS + ["source_file", "source_sheet"]], use_container_width=True)

        with tab3:
            st.subheader("By Market")
            st.dataframe(by_market, use_container_width=True)

        with tab4:
            preview_df = clean_df.copy()
            preview_df["Date"] = preview_df["Date"].dt.strftime("%Y-%m-%d")
            st.dataframe(preview_df[TARGET_COLUMNS + ["source_file", "source_sheet"]], use_container_width=True)

        # 下载 Excel
        excel_bytes = export_workbook(clean_df)
        st.download_button(
            label="下载输出 Excel（Sheet1 Dashboard + Media Sheets）",
            data=excel_bytes,
            file_name="tracking_dashboard_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("当前字段映射规则", expanded=False):
            st.json(COLUMN_MAP)

else:
    st.info("请先上传文件后再生成 Dashboard。")
