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
    # ===== 中文字段（微信）=====
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

    # 1) 文件名判断
    if "wechat" in name or "微信" in name:
        return "WeChat"
    if "douyin" in name or "抖音" in name:
        return "Douyin"
    if "bili" in name or "b站" in name or "bilibili" in name:
        return "Bilibili"

    cols = [norm_text(c) for c in columns]

