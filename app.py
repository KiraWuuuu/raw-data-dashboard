import io
import re
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


st.set_page_config(page_title="Prada Tracking Builder", layout="wide")
st.title("Prada Tracking Builder")


# =========================
# 工具函数
# =========================
def norm_text(v):
    return str(v).strip().lower()


def safe_get_col(df, target):
    for c in df.columns:
        if norm_text(c) == norm_text(target):
            return df[c]
    return pd.Series([np.nan] * len(df))


def to_number(v):
    if pd.isna(v):
        return np.nan
    try:
        v = str(v).replace(",", "").replace("%", "")
        return float(v)
    except:
        return np.nan


def safe_div(a, b):
    if pd.isna(a) or pd.isna(b) or b == 0:
        return np.nan
    return a / b


# =========================
# Excel格式函数（修复语法版本）
# =========================
def format_excel(ws, df):

    accounting = '#,##0.00'
    integer = '#,##0'
    decimal2 = '0.00'
    percent = '0.00%'

    format_map = {
        "Cost": accounting,
        "Revenue": accounting,

        "IMP": integer,
        "CLICK": integer,
        "ENG": integer,
        "Like": integer,
        "Forward": integer,
        "Comment": integer,
        "Orders": integer,

        "CPM": decimal2,
        "CPC": decimal2,
        "CTR": percent
    }

    headers = [cell.value for cell in ws[1]]

    for col_idx, col_name in enumerate(headers, start=1):
        if col_name in format_map:
            fmt = format_map[col_name]
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).number_format = fmt


# =========================
# 数据处理
# =========================
def process_wechat(df):
    out = pd.DataFrame()

    out["Date"] = safe_get_col(df, "日期")
    out["Campaign"] = safe_get_col(df, "广告名称")

    out["Cost"] = safe_get_col(df, "花费").apply(to_number)
    out["IMP"] = safe_get_col(df, "曝光次数").apply(to_number)
    out["CLICK"] = safe_get_col(df, "点击次数").apply(to_number)
    out["CTR"] = safe_get_col(df, "点击率").apply(to_number)

    out["Like"] = safe_get_col(df, "点赞次数").apply(to_number)
    out["Forward"] = safe_get_col(df, "分享次数").apply(to_number)
    out["Comment"] = safe_get_col(df, "评论次数").apply(to_number)

    out["Revenue"] = safe_get_col(df, "下单金额").apply(to_number)
    out["Orders"] = safe_get_col(df, "下单次数").apply(to_number)

    out["ENG"] = out[["Like","Forward","Comment"]].fillna(0).sum(axis=1)

    out["Media"] = "WeChat"
    out["Position"] = "Moments"
    out["Market"] = "Auto"

    return out


def process_douyin(df):
    out = pd.DataFrame()

    out["Date"] = safe_get_col(df, "Date")
    out["Campaign"] = safe_get_col(df, "CampaignName")

    out["IMP"] = safe_get_col(df, "Impression").apply(to_number)
    out["CLICK"] = safe_get_col(df, "Click").apply(to_number)
    out["CTR"] = safe_get_col(df, "CTR").apply(to_number)

    # ✅ 不分摊 cost（关键修复）
    out["Cost"] = np.nan

    out["Media"] = "Douyin"
    return out


def process_weibo(df):
    out = pd.DataFrame()

    out["Date"] = safe_get_col(df, "日期")
    out["IMP"] = safe_get_col(df, "PV").apply(to_number)
    out["CLICK"] = safe_get_col(df, "Click").apply(to_number)

    out["CTR"] = out["CLICK"] / out["IMP"]

    # ✅ 不分摊 cost
    out["Cost"] = np.nan

    out["Media"] = "Weibo"
    return out


# =========================
# 写Excel
# =========================
def recreate_sheet(wb, name):
    if name in wb.sheetnames:
        wb.remove(wb[name])
    return wb.create_sheet(name)


def write_table(ws, df):

    for i, col in enumerate(df.columns, 1):
        ws.cell(1, i, col)

    for r, row in enumerate(df.itertuples(index=False), 2):
        for c, val in enumerate(row, 1):
            ws.cell(r, c, val)

    format_excel(ws, df)


def build_excel(template_file, df):
    wb = load_workbook(template_file)

    ws = recreate_sheet(wb, "Campaign")

    write_table(ws, df)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# =========================
# 页面
# =========================
raw_files = st.file_uploader("Upload Raw Data", accept_multiple_files=True)
template = st.file_uploader("Upload Template")

if st.button("Generate Dashboard"):

    all_data = []

    for f in raw_files:
        df = pd.read_excel(f)

        if "wechat" in f.name.lower():
            std = process_wechat(df)
        elif "douyin" in f.name.lower():
            std = process_douyin(df)
        elif "weibo" in f.name.lower():
            std = process_weibo(df)
        else:
            continue

        all_data.append(std)

    final = pd.concat(all_data)

    final["CPM"] = final["Cost"] / final["IMP"] * 1000
    final["CPC"] = final["Cost"] / final["CLICK"]

    st.dataframe(final)

    excel = build_excel(template, final)

    st.download_button("Download Excel", excel, file_name="Tracking_Final.xlsx")
``
