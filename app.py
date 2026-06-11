import streamlit as st
import pandas as pd
import numpy as np
import io

st.title("Prada Tracking Builder")

# 上传区
raw_files = st.file_uploader("Upload Raw Data", accept_multiple_files=True)
template_file = st.file_uploader("Upload Template", type=["xlsx"])
config_file = st.file_uploader("Upload Mapping Config", type=["xlsx"])


def load_config(file):
    cfg = pd.read_excel(file, sheet_name=None)
    file_map = cfg["file_mapping"]
    market_map = cfg["market_mapping"]
    return file_map, market_map


def process_wechat(df, cfg, market_map):
    df = df.copy()

    df["Date"] = df["日期"]
    df["Campaign"] = df["广告名称"]
    df["Cost"] = df["花费"]
    df["IMP"] = df["曝光次数"]
    df["CLICK"] = df["点击次数"]
    df["CTR"] = df["点击率"]

    df["Like"] = df["点赞次数"]
    df["Forward"] = df["分享次数"]
    df["Comment"] = df["评论次数"]

    df["Revenue"] = df["下单金额"]
    df["Orders"] = df["下单次数"]

    df["ENG"] = df["Like"] + df["Forward"] + df["Comment"]

    df["Media"] = cfg["media_name"]
    df["Position"] = cfg["position_value"]

    # Market mapping
    def map_market(text):
        for _, row in market_map.iterrows():
            if row["keyword"] in str(text):
                return row["market_bucket"]
        return "Unknown"

    df["Market"] = df["Campaign"].apply(map_market)

    return df


def process_douyin(df, cfg):
    df = df.copy()

    df["Date"] = df["Date"]
    df["Campaign"] = df["CampaignName"]
    df["IMP"] = df["Impression"]
    df["CLICK"] = df["Click"]
    df["CTR"] = df["CTR"]

    df["Market"] = df["Region"]

    df["Position"] = df[cfg["position_value"]]
    df["Media"] = cfg["target_sheet"]

    return df


def process_weibo(df, cfg):
    df = df.copy()

    df["Date"] = df["日期"]
    df["IMP"] = df["PV"]
    df["CLICK"] = df["Click"]
    df["Position"] = df["点位"]

    df["Media"] = "Weibo"
    df["Market"] = cfg["market_value"]

    return df


def allocate_cost(df, total):
    if total is None or pd.isna(total):
        return df

    total_imp = df["IMP"].sum()
    df["Cost"] = df["IMP"] / total_imp * total
    return df


if st.button("Generate Dashboard"):

    file_map, market_map = load_config(config_file)

    all_data = []

    for f in raw_files:
        df = pd.read_excel(f)

        cfg = file_map[file_map["source_file"] == f.name].iloc[0]

        if "wechat" in f.name:
            df = process_wechat(df, cfg, market_map)

        elif "douyin feeds" in f.name:
            df = process_douyin(df, cfg)

        elif "douyin opening" in f.name:
            df = process_douyin(df, cfg)

        elif "weibo" in f.name:
            df = process_weibo(df, cfg)

        if cfg["cost_mode"] == "allocate_total":
            df = allocate_cost(df, cfg["cost_total"])

        all_data.append(df)

    final_df = pd.concat(all_data)

    final_df["CPM"] = final_df["Cost"] / final_df["IMP"] * 1000
    final_df["CPC"] = final_df["Cost"] / final_df["CLICK"]

    st.dataframe(final_df)

    output = io.BytesIO()
    final_df.to_excel(output, index=False)

    st.download_button("Download Excel", output.getvalue(), file_name="output.xlsx")
