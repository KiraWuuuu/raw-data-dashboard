
import io
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Raw Data Dashboard Generator", layout="wide")

TARGET_COLUMNS = [
    "Media", "Position", "Market", "Date", "Landing page",
    "Cost", "IMP", "CLICK", "CPM", "CPC", "CTR",
    "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"
]

COLUMN_MAP = {
    # Media
    "media": "Media", "platform": "Media", "媒体": "Media", "渠道": "Media", "publisher": "Media",
    # Position
    "position": "Position", "placement": "Position", "ad format": "Position", "广告位": "Position", "版位": "Position",
    # Market
    "market": "Market", "region": "Market", "城市": "Market", "地区": "Market", "market name": "Market",
    # Date
    "date": "Date", "日期": "Date", "投放日期": "Date", "day": "Date", "时间": "Date",
    # URL
    "landing page": "Landing page", "landing_page": "Landing page", "landingpage": "Landing page", "落地页": "Landing page", "url": "Landing page",
    # Spend
    "cost": "Cost", "spend": "Cost", "花费": "Cost", "消耗": "Cost", "actual spend": "Cost", "media cost": "Cost",
    # IMP
    "imp": "IMP", "impression": "IMP", "impressions": "IMP", "曝光": "IMP", "曝光量": "IMP", "display": "IMP",
    # Click
    "click": "CLICK", "clicks": "CLICK", "点击": "CLICK", "点击量": "CLICK",
    # KPI
    "cpm": "CPM", "cpc": "CPC", "ctr": "CTR", "点击率": "CTR",
    # Engagement
    "eng": "ENG", "engagement": "ENG", "互动": "ENG", "互动量": "ENG",
    "like": "Like", "likes": "Like", "点赞": "Like",
    "forward": "Forward", "share": "Forward", "shares": "Forward", "转发": "Forward",
    "comment": "Comment", "comments": "Comment", "评论": "Comment",
    # Business
    "revenue": "Revenue", "gmv": "Revenue", "销售额": "Revenue", "成交金额": "Revenue",
    "orders": "Orders", "order": "Orders", "订单": "Orders", "订单量": "Orders",
}

PLATFORM_ALIAS = {
    "douyin": "Douyin",
    "抖音": "Douyin",
    "tiktok": "Douyin",
    "wechat": "WeChat",
    "微信": "WeChat",
    "weixin": "WeChat",
    "xhs": "Xiaohongshu",
    "red": "Xiaohongshu",
    "小红书": "Xiaohongshu",
    "xiaohongshu": "Xiaohongshu",
    "bilibili": "Bilibili",
    "b站": "Bilibili",
    "tencent video": "Tencent Video",
    "腾讯视频": "Tencent Video",
    "xiaomi": "Xiaomi",
    "小米": "Xiaomi",
}

NUMERIC_COLUMNS = ["Cost", "IMP", "CLICK", "CPM", "CPC", "CTR", "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"]


def norm_text(v):
    return str(v if v is not None else "").strip().lower().replace("_", " ").replace("-", " ")


def map_column(col: str) -> str:
    key = norm_text(col)
    return COLUMN_MAP.get(key, col)


def normalize_media(v):
    key = norm_text(v)
    if key in PLATFORM_ALIAS:
        return PLATFORM_ALIAS[key]
    value = str(v).strip() if pd.notna(v) else ""
    return value if value else "Unknown"


def to_number(v):
    if pd.isna(v) or v == "":
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return pd.to_numeric(v, errors="coerce")
    s = str(v).strip().replace(",", "").replace("¥", "").replace("￥", "").replace("$", "")
    if s.endswith("%"):
        s = s[:-1]
        try:
            return float(s) / 100
        except:
            return np.nan
    try:
        return float(s)
    except:
        return np.nan


def safe_div(a, b):
    if pd.isna(a) or pd.isna(b) or float(b) == 0:
        return np.nan
    return a / b


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [map_column(c) for c in df.columns]
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan
    df = df[TARGET_COLUMNS + [c for c in df.columns if c not in TARGET_COLUMNS]]
    return df


def read_uploaded_file(uploaded_file) -> List[pd.DataFrame]:
    suffix = uploaded_file.name.lower().split('.')[-1]
    results = []
    if suffix == 'csv':
        df = pd.read_csv(uploaded_file)
        df['source_file'] = uploaded_file.name
        df['source_sheet'] = 'csv'
        results.append(df)
    else:
        data = uploaded_file.getvalue()
        excel_io = io.BytesIO(data)
        sheets = pd.read_excel(excel_io, sheet_name=None, engine='openpyxl')
        for sheet_name, df in sheets.items():
            df['source_file'] = uploaded_file.name
            df['source_sheet'] = sheet_name
            results.append(df)
    return results


def clean_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = ensure_columns(raw_df)

    # normalize main fields
    df['Media'] = df['Media'].apply(normalize_media)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    for col in NUMERIC_COLUMNS:
        df[col] = df[col].apply(to_number)

    # derived KPIs if missing
    df['CTR'] = df['CTR'].where(df['CTR'].notna(), df.apply(lambda x: safe_div(x['CLICK'], x['IMP']), axis=1))
    df['CPM'] = df['CPM'].where(df['CPM'].notna(), df.apply(lambda x: safe_div(x['Cost'] * 1000, x['IMP']), axis=1))
    df['CPC'] = df['CPC'].where(df['CPC'].notna(), df.apply(lambda x: safe_div(x['Cost'], x['CLICK']), axis=1))

    # keep display format for date in exports
    return df


def build_dashboard_tables(df: pd.DataFrame):
    total_cost = df['Cost'].sum(skipna=True)
    total_imp = df['IMP'].sum(skipna=True)
    total_click = df['CLICK'].sum(skipna=True)
    total_eng = df['ENG'].sum(skipna=True)
    total_revenue = df['Revenue'].sum(skipna=True)
    total_orders = df['Orders'].sum(skipna=True)

    summary = pd.DataFrame([{
        'Total Cost': total_cost,
        'Total IMP': total_imp,
        'Total CLICK': total_click,
        'Overall CTR': safe_div(total_click, total_imp),
        'Total ENG': total_eng,
        'Total Revenue': total_revenue,
        'Total Orders': total_orders,
        'Overall CPM': safe_div(total_cost * 1000, total_imp),
        'Overall CPC': safe_div(total_cost, total_click),
    }])

    by_media = df.groupby('Media', dropna=False)[['Cost', 'IMP', 'CLICK', 'Revenue', 'Orders', 'ENG', 'Like', 'Forward', 'Comment']].sum(min_count=1).reset_index()
    by_market = df.groupby('Market', dropna=False)[['Cost', 'IMP', 'CLICK', 'Revenue', 'Orders']].sum(min_count=1).reset_index()

    date_tmp = df.copy()
    date_tmp['Date'] = date_tmp['Date'].dt.strftime('%Y-%m-%d')
    by_date = date_tmp.groupby('Date', dropna=False)[['Cost', 'IMP', 'CLICK', 'Revenue', 'Orders']].sum(min_count=1).reset_index()

    return summary, by_media, by_market, by_date


def export_workbook(df: pd.DataFrame) -> bytes:
    from openpyxl.styles import Font, PatternFill, Alignment

    output = io.BytesIO()
    dashboard_df = df.copy()
    dashboard_df['Date'] = dashboard_df['Date'].dt.strftime('%Y-%m-%d')
    summary, by_media, by_market, by_date = build_dashboard_tables(df)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_name = 'Sheet1_Dashboard'
        summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
        by_media.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4)
        by_market.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4 + len(by_media) + 3)
        by_date.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4 + len(by_media) + len(by_market) + 6)

        workbook = writer.book
        ws = writer.sheets[sheet_name]
        ws['A1'] = 'Dashboard Summary'
        ws['A5'] = 'By Media'
        ws[f'A{5 + len(by_media) + 3}'] = 'By Market'
        ws[f'A{5 + len(by_media) + len(by_market) + 6}'] = 'By Date'
        title_fill = PatternFill(fill_type='solid', fgColor='1F4E78')
        header_fill = PatternFill(fill_type='solid', fgColor='D9EAF7')
        white_font = Font(color='FFFFFF', bold=True)
        bold_font = Font(bold=True)
        for cell in ['A1', 'A5', f'A{5 + len(by_media) + 3}', f'A{5 + len(by_media) + len(by_market) + 6}']:
            ws[cell].fill = title_fill
            ws[cell].font = white_font
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                if cell.row in [2, 6, 6 + len(by_media) + 3, 6 + len(by_media) + len(by_market) + 6]:
                    cell.fill = header_fill
                    cell.font = bold_font
                cell.alignment = Alignment(vertical='center')

        # Media sheets
        for media_name, sub_df in dashboard_df.groupby('Media', dropna=False):
            safe_name = str(media_name if pd.notna(media_name) else 'Unknown').replace('/', '_')[:31]
            temp = sub_df[TARGET_COLUMNS + ['source_file', 'source_sheet']].copy()
            temp.to_excel(writer, sheet_name=safe_name, index=False)
            media_ws = writer.sheets[safe_name]
            for cell in media_ws[1]:
                cell.fill = header_fill
                cell.font = bold_font

    output.seek(0)
    return output.getvalue()


def format_num(v, pct=False):
    if pd.isna(v):
        return '-'
    if pct:
        return f"{v:.2%}"
    return f"{v:,.2f}"


st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
.hero {
    background: linear-gradient(135deg, #0F4C81 0%, #2A9FD6 100%);
    padding: 24px 28px; border-radius: 18px; color: white; margin-bottom: 1rem;
}
.metric-card {
    background: #ffffff; border: 1px solid #e9eef5; border-radius: 16px; padding: 14px 16px;
    box-shadow: 0 2px 10px rgba(15, 76, 129, 0.05);
}
.small {font-size: 12px; color: #6b7280;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='hero'>
  <h1 style='margin:0'>Raw Data Dashboard Generator</h1>
  <p style='margin:8px 0 0 0'>上传多个不同媒体的 raw data Excel，自动按目标字段标准化，生成 Dashboard，并导出 Excel（Sheet1 为 Dashboard，后续 Sheet 按 Media 拆分明细）。</p>
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
            raw_parts.extend(read_uploaded_file(uf))
        except Exception as e:
            st.error(f"读取文件失败：{uf.name}，错误：{e}")

    if raw_parts:
        raw_df = pd.concat(raw_parts, ignore_index=True)
        clean_df = clean_data(raw_df)
        summary, by_media, by_market, by_date = build_dashboard_tables(clean_df)

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1: st.metric("Total Cost", format_num(summary.loc[0, 'Total Cost']))
        with c2: st.metric("Total IMP", format_num(summary.loc[0, 'Total IMP']))
        with c3: st.metric("Total CLICK", format_num(summary.loc[0, 'Total CLICK']))
        with c4: st.metric("Overall CTR", format_num(summary.loc[0, 'Overall CTR'], pct=True))
        with c5: st.metric("Revenue", format_num(summary.loc[0, 'Total Revenue']))
        with c6: st.metric("Orders", format_num(summary.loc[0, 'Total Orders']))

        tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "By Media", "By Market", "Preview Data"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Spend by Media")
                st.bar_chart(by_media.set_index('Media')[['Cost']])
            with col2:
                date_plot = by_date.copy()
                date_plot = date_plot[date_plot['Date'].notna() & (date_plot['Date'] != 'NaT')]
                if not date_plot.empty:
                    st.subheader("Trend by Date")
                    st.line_chart(date_plot.set_index('Date')[['Cost', 'IMP', 'CLICK']])
                else:
                    st.info("没有可用于日期趋势图的 Date 数据。")

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
            selected_media = st.selectbox("选择 Media 预览对应 sheet 内容", options=sorted(by_media['Media'].dropna().astype(str).unique().tolist()))
            if selected_media:
                temp = clean_df[clean_df['Media'].astype(str) == selected_media].copy()
                temp['Date'] = temp['Date'].dt.strftime('%Y-%m-%d')
                st.dataframe(temp[TARGET_COLUMNS + ['source_file', 'source_sheet']], use_container_width=True)

        with tab3:
            st.subheader("By Market")
            st.dataframe(by_market, use_container_width=True)

        with tab4:
            preview = clean_df.copy()
            preview['Date'] = preview['Date'].dt.strftime('%Y-%m-%d')
            st.dataframe(preview[TARGET_COLUMNS + ['source_file', 'source_sheet']], use_container_width=True)

        excel_bytes = export_workbook(clean_df)
        st.download_button(
            label="下载输出 Excel（Sheet1 Dashboard + Media Sheets）",
            data=excel_bytes,
            file_name="tracking_dashboard_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("字段映射说明", expanded=False):
            st.json(COLUMN_MAP)
else:
    st.info("请先上传文件后再生成 Dashboard。")
