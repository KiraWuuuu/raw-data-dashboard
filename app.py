import io
 typing import Dict, List, Optionalimport re

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Prada Tracking Builder", layout="wide")

# =========================
# 页面标题
# =========================
st.title("Prada Tracking Builder")
st.caption("上传 Preview Raw Data + Tracking Template + Mapping Config，自动生成输出 Excel。")

# =========================
# 目标标准字段
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
def norm_text(v: str) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip().lower().replace("_", " ").replace("-", " "))

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

def find_file_rule(file_map_df: pd.DataFrame, source_file: str) -> pd.Series:
    match = file_map_df[file_map_df["source_file"].astype(str) == str(source_file)]
    if match.empty:
        raise ValueError(f"在 mapping_config 的 file_mapping 中找不到文件规则：{source_file}")
    return match.iloc[0]

def map_market_by_keyword(text: str, market_map_df: pd.DataFrame) -> str:
    txt = str(text)
    for _, row in market_map_df.iterrows():
        keyword = str(row["keyword"])
        market_bucket = str(row["market_bucket"])
        if keyword and keyword in txt:
            return market_bucket
    return "Unknown"

def allocate_cost_by_imp(df: pd.DataFrame, total_cost: Optional[float]) -> pd.DataFrame:
    if total_cost is None or pd.isna(total_cost):
        return df
    imp_sum = df["IMP"].sum(skipna=True)
    if pd.isna(imp_sum) or imp_sum == 0:
        df["Cost"] = np.nan
        return df
    df["Cost"] = df["IMP"] / imp_sum * float(total_cost)
    return df

def add_kpi_fields(df: pd.DataFrame) -> pd.DataFrame:
    df["CTR"] = df["CTR"].where(df["CTR"].notna(), df.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1))
    df["CPM"] = df["CPM"].where(df["CPM"].notna(), df.apply(lambda x: safe_div(x["Cost"] * 1000, x["IMP"]), axis=1))
    df["CPC"] = df["CPC"].where(df["CPC"].notna(), df.apply(lambda x: safe_div(x["Cost"], x["CLICK"]), axis=1))
    return df

def format_date_col(df: pd.DataFrame, col: str = "Date") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

# =========================
# 读取 mapping 配置
# =========================
def load_mapping_config(config_file):
    sheets = pd.read_excel(config_file, sheet_name=None, engine="openpyxl")
    if "file_mapping" not in sheets:
        raise ValueError("mapping_config.xlsx 缺少 sheet：file_mapping")
    if "market_mapping" not in sheets:
        raise ValueError("mapping_config.xlsx 缺少 sheet：market_mapping")

    file_map = sheets["file_mapping"].copy()
    market_map = sheets["market_mapping"].copy()

    # 标准化列，避免大小写/空格问题
    file_map.columns = [str(c).strip() for c in file_map.columns]
    market_map.columns = [str(c).strip() for c in market_map.columns]

    return file_map, market_map

# =========================
# 读取 raw 文件
# =========================
def read_excel_first_sheet(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(data), sheet_name=first_sheet, engine="openpyxl")
    return df

# =========================
# 三类媒体解析器
# =========================
def process_wechat(df: pd.DataFrame, cfg: pd.Series, market_map: pd.DataFrame) -> pd.DataFrame:
    # 微信文件字段来自 raw：日期 / 广告名称 / 花费 / 曝光次数 / 点击次数 / 点击率 / 点赞次数 / 分享次数 / 评论次数 / 下单金额 / 下单次数
    out = pd.DataFrame()
    out["Date"] = df["日期"]
    out["Campaign"] = df["广告名称"]
    out["Cost"] = df["花费"].apply(to_number)
    out["IMP"] = df["曝光次数"].apply(to_number)
    out["CLICK"] = df["点击次数"].apply(to_number)
    out["CTR"] = df["点击率"].apply(to_number)
    out["Like"] = df["点赞次数"].apply(to_number)
    out["Forward"] = df["分享次数"].apply(to_number)
    out["Comment"] = df["评论次数"].apply(to_number)
    out["Revenue"] = df["下单金额"].apply(to_number)
    out["Orders"] = df["下单次数"].apply(to_number)

    out["ENG"] = out[["Like", "Forward", "Comment"]].fillna(0).sum(axis=1)
    out["Media"] = cfg["media_name"]
    out["Position"] = cfg["position_value"] if str(cfg["position_mode"]).lower() == "fixed" else df[str(cfg["position_value"])]

    # market_mode = keyword
    if str(cfg["market_mode"]).lower() == "keyword":
        out["Market"] = out["Campaign"].apply(lambda x: map_market_by_keyword(x, market_map))
    elif str(cfg["market_mode"]).lower() == "fixed":
        out["Market"] = cfg["market_value"]
    elif str(cfg["market_mode"]).lower() == "raw":
        out["Market"] = df[str(cfg["market_value"])]
    else:
        out["Market"] = "Unknown"

    out["Landing page"] = cfg["landing_page_default"]
    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out = format_date_col(out, "Date")
    out = add_kpi_fields(out)

    for col in STANDARD_COLUMNS:
        if col not in out.columns:
            out[col] = np.nan

    return out[STANDARD_COLUMNS]

def process_douyin(df: pd.DataFrame, cfg: pd.Series) -> pd.DataFrame:
    # 抖音文件字段来自 raw：Region / Type / SPID / Website / Channel / Ad Placement / Campaign ID / CampaignName / Date / Impression / Click / CTR
    # 为避免 KeyError，对列名做 strip
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    out = pd.DataFrame()
    out["Date"] = df["Date"]
    out["Campaign"] = df["CampaignName"]
    out["IMP"] = df["Impression"].apply(to_number)
    out["CLICK"] = df["Click"].apply(to_number)
    out["CTR"] = df["CTR"].apply(to_number)

    # market_mode = raw / market_value = Region
    if str(cfg["market_mode"]).lower() == "raw":
        raw_market_col = str(cfg["market_value"]).strip()
        out["Market"] = df[raw_market_col] if raw_market_col in df.columns else "Unknown"
    elif str(cfg["market_mode"]).lower() == "fixed":
        out["Market"] = cfg["market_value"]
    else:
        out["Market"] = "Unknown"

    # position_mode
    if str(cfg["position_mode"]).lower() == "raw":
        raw_pos_col = str(cfg["position_value"]).strip()
        out["Position"] = df[raw_pos_col] if raw_pos_col in df.columns else "Unknown"
    else:
        out["Position"] = cfg["position_value"]

    out["Media"] = cfg["media_name"]
    out["Landing page"] = cfg["landing_page_default"]

    # Cost：如果 raw 没有，就按 IMP 分摊
    out["Cost"] = np.nan
    if str(cfg["cost_mode"]).lower() == "allocate_total":
        out = allocate_cost_by_imp(out, cfg["cost_total"])
    elif str(cfg["cost_mode"]).lower() == "raw":
        if "Cost" in df.columns:
            out["Cost"] = df["Cost"].apply(to_number)

    out["Like"] = np.nan
    out["Forward"] = np.nan
    out["Comment"] = np.nan
    out["ENG"] = np.nan
    out["Revenue"] = np.nan
    out["Orders"] = np.nan

    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out = format_date_col(out, "Date")
    out = add_kpi_fields(out)

    for col in STANDARD_COLUMNS:
        if col not in out.columns:
            out[col] = np.nan

    return out[STANDARD_COLUMNS]

def process_weibo(df: pd.DataFrame, cfg: pd.Series) -> pd.DataFrame:
    # 微博文件字段来自 raw：日期 / 点位 / PV / Click
    out = pd.DataFrame()
    out["Date"] = df["日期"]
    out["Position"] = df["点位"]
    out["IMP"] = df["PV"].apply(to_number)
    out["CLICK"] = df["Click"].apply(to_number)
    out["CTR"] = out.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1)

    out["Media"] = cfg["media_name"]

    # 微博 market 建议 fixed=National
    if str(cfg["market_mode"]).lower() == "fixed":
        out["Market"] = cfg["market_value"]
    else:
        out["Market"] = "Unknown"

    out["Campaign"] = np.nan
    out["Landing page"] = cfg["landing_page_default"]

    out["Cost"] = np.nan
    if str(cfg["cost_mode"]).lower() == "allocate_total":
        out = allocate_cost_by_imp(out, cfg["cost_total"])

    out["Like"] = np.nan
    out["Forward"] = np.nan
    out["Comment"] = np.nan
    out["ENG"] = np.nan
    out["Revenue"] = np.nan
    out["Orders"] = np.nan

    out["source_file"] = cfg["source_file"]
    out["target_sheet"] = cfg["target_sheet"]

    out = format_date_col(out, "Date")
    out = add_kpi_fields(out)

    for col in STANDARD_COLUMNS:
        if col not in out.columns:
            out[col] = np.nan

    return out[STANDARD_COLUMNS]

# =========================
# 汇总页生成
# =========================
def build_campaign_sheet_df(all_std: pd.DataFrame) -> pd.DataFrame:
    """
    这里做“模板复刻版”的第一步：
    先把标准表按 Campaign 页的核心字段聚合成扁平表。
    如果你后续要完全复刻模板布局（按块写入），可以在这个函数基础上继续扩展。
    """
    df = all_std.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    group_cols = ["Media", "Position", "Market", "Date"]
    value_cols = ["Cost", "IMP", "CLICK", "Revenue", "Orders", "ENG", "Like", "Forward", "Comment"]

    out = df.groupby(group_cols, dropna=False)[value_cols].sum(min_count=1).reset_index()
    out["CTR"] = out.apply(lambda x: safe_div(x["CLICK"], x["IMP"]), axis=1)
    out["CPM"] = out.apply(lambda x: safe_div(x["Cost"] * 1000, x["IMP"]), axis=1)
    out["CPC"] = out.apply(lambda x: safe_div(x["Cost"], x["CLICK"]), axis=1)

    # 列顺序按目标 tracking dashboard
    out = out[[
        "Media", "Position", "Market", "Date",
        "Cost", "IMP", "CLICK", "CPM", "CPC", "CTR",
        "ENG", "Like", "Forward", "Comment", "Revenue", "Orders"
    ]]

    return out

# =========================
# 写入模板
# =========================
def write_dataframe_to_sheet(ws, df: pd.DataFrame):
    # 清空旧内容（简单版：把已有单元格值清掉）
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    # 写表头
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col_name)
        ws.cell(row=1, column=c_idx).font = Font(bold=True)
        ws.cell(row=1, column=c_idx).fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    # 写数据
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    # 简单列宽
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 50)

def build_output_workbook(template_file, all_std: pd.DataFrame) -> bytes:
    data = template_file.getvalue()
    wb = load_workbook(io.BytesIO(data))

    # 1) 写 detail sheets
    for target_sheet, sub_df in all_std.groupby("target_sheet", dropna=False):
        if str(target_sheet) in wb.sheetnames:
            ws = wb[str(target_sheet)]
            export_df = sub_df.copy()
            export_df["Date"] = pd.to_datetime(export_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

            # 按目标 sheet 决定导出字段
            if str(target_sheet) in ["【WeChat Moments】", "【WeChat Banner】"]:
                detail_df = export_df[[
                    "Market", "Date", "Campaign", "Cost", "IMP", "CLICK", "CTR",
                    "Like", "Forward", "Comment", "Revenue", "Orders"
                ]].copy()
                detail_df.columns = [
                    "Market", "日期", "广告名称", "花费", "曝光次数", "点击次数", "点击率",
                    "点赞次数", "分享次数", "评论次数", "下单金额", "下单次数"
                ]
            else:
                # Douyin / Weibo 这类 ByDay 结构先输出标准字段版本
                detail_df = export_df[[
                    "Market", "Position", "Campaign", "Date", "IMP", "CLICK", "CTR"
                ]].copy()
                detail_df.columns = [
                    "Region", "Ad Placement", "CampaignName", "Date", "Impression", "Click", "CTR"
                ]

            write_dataframe_to_sheet(ws, detail_df)

    # 2) 写 Campaign 汇总页（如果存在）
    if "Campaign" in wb.sheetnames:
        ws_campaign = wb["Campaign"]
        campaign_df = build_campaign_sheet_df(all_std)
        write_dataframe_to_sheet(ws_campaign, campaign_df)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# =========================
# 页面输入区
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
        st.error("请先上传 Preview Raw Data 文件。")
        st.stop()
    if template_file is None:
        st.error("请先上传 Template 文件。")
        st.stop()
    if config_file is None:
        st.error("请先上传 Mapping Config 文件。")
        st.stop()

    try:
        file_map_df, market_map_df = load_mapping_config(config_file)
    except Exception as e:
        st.error(f"读取 mapping_config 失败：{e}")
        st.stop()

    all_std_parts = []

    for f in raw_files:
        try:
            cfg = find_file_rule(file_map_df, f.name)
            df_raw = read_excel_first_sheet(f)

            lower_name = f.name.lower()

            if "wechat" in lower_name or "微信" in lower_name:
                std_df = process_wechat(df_raw, cfg, market_map_df)

            elif "douyin feeds" in lower_name or ("douyin" in lower_name and "feeds" in lower_name):
                std_df = process_douyin(df_raw, cfg)

            elif "douyin opening" in lower_name or ("douyin" in lower_name and "opening" in lower_name):
                std_df = process_douyin(df_raw, cfg)

            elif "weibo" in lower_name or "微博" in lower_name:
                std_df = process_weibo(df_raw, cfg)

            else:
                st.warning(f"文件 {f.name} 暂时没有匹配到专属解析规则，已跳过。")
                continue

            all_std_parts.append(std_df)

        except Exception as e:
            st.error(f"处理文件 {f.name} 失败：{e}")
            st.stop()

    if not all_std_parts:
        st.error("没有可生成的数据。")
        st.stop()

    all_std = pd.concat(all_std_parts, ignore_index=True)

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

