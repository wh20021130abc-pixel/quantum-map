# -*- coding: utf-8 -*-
"""
量子科技产业地图 - Streamlit 应用（v6 企业介绍页版）
"""

import os
import numpy as np
import pandas as pd
import streamlit as st

# ==================== 路径 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = BASE_DIR

COMPANY_FILE = os.path.join(DATA_DIR, "量子科技-企业基本信息汇总.xlsx")
KEYWORD_FILE = os.path.join(DATA_DIR, "量子科技企业关键词识别-名称、简介、经营范围、行业.xlsx")
PATENT_FILE = os.path.join(DATA_DIR, "量子科技企业-专利汇总.xlsx")
PATENT_KW_FILE = os.path.join(DATA_DIR, "专利关键词.xlsx")

# ==================== 配色 ====================
PRIMARY = "#2563EB"
PRIMARY_DARK = "#1D4ED8"
PRIMARY_SOFT = "#EFF6FF"
TEXT = "#1A2332"
TEXT_SUB = "#6B7A8F"
TEXT_LIGHT = "#94A3B8"
BORDER = "#E5EAF2"
BG = "#F7F9FC"

# ==================== 页面配置 ====================
st.set_page_config(page_title="量子科技产业地图", page_icon="", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown("".join([
    "<style>",
    '#MainMenu {visibility: hidden;}',
    'header[data-testid="stHeader"] {display: none;}',
    'footer {visibility: hidden;}',
    'html, body, [class*="css"], .stApp {',
    '    font-family: "Microsoft YaHei", -apple-system, sans-serif !important;',
    "}",
    f".stApp {{ background: {BG}; }}",
    ".block-container { padding-top: 1.8rem; padding-bottom: 2rem; max-width: 100%; }",

    # 页面标题
    f".page-title {{ font-size: 22px; font-weight: 800; color: {TEXT};",
    "    margin: 0; letter-spacing: 0.3px; }",
    f".page-sub {{ font-size: 13px; color: {TEXT_SUB}; margin-top: 6px; }}",
    ".page-header { margin-bottom: 26px; }",

    # 分组标题
    f".group-title {{ font-size: 12px; color: {TEXT_SUB}; font-weight: 600;",
    "    letter-spacing: 0.5px; margin: 20px 0 10px 0; }",

    # 领域按钮
    'div[data-testid="stButton"] > button {',
    "    width: 100%; text-align: left; padding: 10px 14px;",
    "    border-radius: 8px;",
    f"    border: 1px solid {BORDER};",
    "    background: white; font-size: 13.5px;",
    f"    color: {TEXT}; font-weight: 500;",
    "    margin-bottom: 6px; transition: all 0.15s; box-shadow: none;",
    '    font-family: "Microsoft YaHei", sans-serif !important;',
    "}",
    'div[data-testid="stButton"] > button:hover {',
    f"    background: {PRIMARY_SOFT}; border-color: {PRIMARY}; color: {PRIMARY_DARK};",
    "}",
    'div[data-testid="stButton"] > button[kind="primary"] {',
    f"    background: {PRIMARY}; color: white; border-color: {PRIMARY};",
    "}",
    'div[data-testid="stButton"] > button[kind="primary"]:hover {',
    f"    background: {PRIMARY_DARK}; border-color: {PRIMARY_DARK};",
    "}",

    # 卡片头部
    ".card-name { font-size: 17px; font-weight: 700; margin-bottom: 10px;",
    f"    color: {TEXT}; line-height: 1.4; }}".replace("}}", "}"),
    ".badge-high { background: #DCFCE7; color: #166534; padding: 2px 10px;",
    "    border-radius: 4px; font-size: 12px; font-weight: 600;",
    "    margin-left: 8px; vertical-align: middle; }",
    ".badge-low { background: #FEF3C7; color: #92400E; padding: 2px 10px;",
    "    border-radius: 4px; font-size: 12px; font-weight: 600;",
    "    margin-left: 8px; vertical-align: middle; }",

    # meta pill
    ".meta-pill { display: inline-block; background: #F1F5F9; color: #475569;",
    "    padding: 3px 10px; border-radius: 6px; font-size: 12.5px;",
    "    margin-right: 6px; margin-bottom: 4px; font-weight: 500; }",
    f".meta-pill-blue {{ background: {PRIMARY_SOFT}; color: {PRIMARY_DARK}; }}",

    # 赛道标签
    f".tag {{ display: inline-block; background: {PRIMARY_SOFT}; color: {PRIMARY_DARK};",
    "    padding: 3px 10px; border-radius: 5px; font-size: 12.5px;",
    "    margin-right: 6px; margin-top: 5px; margin-bottom: 5px; font-weight: 500; }",

    # popover 分数按钮
    'div[data-testid="stPopover"] > button {',
    f"    background: {PRIMARY} !important; border: none !important;",
    "    color: white !important; font-weight: 700 !important;",
    "    font-size: 15px !important; border-radius: 8px !important;",
    "    padding: 8px 16px !important; width: 100% !important;",
    '    font-family: "Microsoft YaHei", sans-serif !important;',
    "}",
    'div[data-testid="stPopover"] > button:hover {',
    f"    background: {PRIMARY_DARK} !important;",
    "}",

    # 弹层明细
    ".detail-row { display: flex; justify-content: space-between;",
    "    padding: 9px 0; border-bottom: 1px solid #F1F5F9; font-size: 13px; }",
    f".detail-label {{ color: {TEXT_SUB}; }}",
    f".detail-val {{ color: {TEXT}; font-weight: 700; }}",
    ".domain-block { padding: 10px 0; }",
    ".domain-head { display: flex; justify-content: space-between;",
    "    align-items: center; margin-bottom: 6px; }",
    f".domain-name {{ font-size: 12.5px; color: {TEXT_SUB}; }}",
    f".domain-score {{ font-size: 16px; font-weight: 800; color: {PRIMARY}; }}",
    ".domain-bar-bg { height: 6px; background: #F1F5F9; border-radius: 3px; overflow: hidden; }",
    f".domain-bar {{ height: 100%; background: {PRIMARY}; border-radius: 3px; }}",

    # expander 微调
    'div[data-testid="stExpander"] details {',
    f"    border: 1px solid {BORDER}; border-radius: 8px; background: white;",
    "}",
    'div[data-testid="stExpander"] summary {',
    f"    font-size: 13.5px; color: {PRIMARY}; font-weight: 600;",
    "    padding: 10px 16px;",
    "}",

    # 企业介绍页 - 信息网格
    ".profile-block { padding: 4px 0 12px 0; }",
    ".profile-label { font-size: 12.5px; color: #94A3B8; margin-bottom: 4px; }",
    ".profile-value { font-size: 14px; color: #1A2332; font-weight: 600; line-height: 1.5; }",
    ".profile-block-large { padding: 12px 0; }",

    # 国标行业层级链
    ".hier-chain { display: flex; flex-wrap: wrap; align-items: center;",
    "    gap: 6px; margin-top: 6px; }",
    ".hier-chip { background: #F8FAFC; border: 1px solid #E5EAF2;",
    "    color: #475569; padding: 4px 10px; border-radius: 6px;",
    "    font-size: 12.5px; }",
    ".hier-chip-active { background: #EFF6FF; border-color: #BFDBFE;",
    "    color: #1D4ED8; }",
    ".hier-arrow { color: #CBD5E1; font-size: 12px; }",

    # 专利列表
    ".patent-item { padding: 10px 0; border-bottom: 1px solid #F1F5F9; }",
    ".patent-item:last-child { border-bottom: none; }",
    ".patent-title { font-size: 13.5px; color: #1A2332; font-weight: 500;",
    "    line-height: 1.5; margin-bottom: 4px; }",
    ".patent-meta { font-size: 12px; color: #94A3B8; }",
    ".patent-code { color: #2563EB; font-weight: 600; margin-right: 10px; }",
    "</style>",
]), unsafe_allow_html=True)

# ==================== 关键词加载 ====================
BLACKLIST = {"直接关键词", "间接关键词", "量子计算_关键词", "量子通信_关键词",
             "量子精密测量_关键词", "词频", "命中企业数"}

def _clean(series):
    out, seen = [], set()
    for v in series:
        if pd.isna(v): continue
        s = str(v).strip()
        if not s or s in BLACKLIST or s in seen: continue
        seen.add(s); out.append(s)
    return out

@st.cache_data
def load_company_keywords(path):
    sheets = {"名称": "名称", "简介": "简介", "经营范围": "经营范围", "市场化": "市场化行业"}
    result = {}
    for key, sheet in sheets.items():
        df = pd.read_excel(path, sheet_name=sheet, header=1)
        result[key] = {
            "direct": _clean(df.iloc[:, 0]),
            "indirect": _clean(df.iloc[:, 3]),
            "量子计算": _clean(df.iloc[:, 6]),
            "量子通信": _clean(df.iloc[:, 9]),
            "量子精密测量": _clean(df.iloc[:, 12]),
        }
    return result

@st.cache_data
def load_patent_keywords(path):
    df = pd.read_excel(path, sheet_name="专利", header=0)
    return {
        "direct": _clean(df.iloc[:, 0]),
        "量子计算": _clean(df.iloc[:, 3]),
        "量子通信": _clean(df.iloc[:, 6]),
        "量子精密测量": _clean(df.iloc[:, 9]),
    }

# ==================== 打分 ====================
FIELD_WEIGHTS = {"名称": 3, "简介": 2, "经营范围": 1, "市场化": 2}
FIELD_COLS = {
    "名称": ["企业名称", "曾用名", "企业简称"],
    "简介": ["公司简介"],
    "经营范围": ["经营范围"],
    "市场化": ["来觅赛道(一级市场)", "企查查行业中类", "企查查行业小类",
             "Wind行业分类-一级", "Wind行业分类-二级",
             "Wind行业分类-三级", "Wind行业分类-四级"],
}
DOMAINS = ["量子计算", "量子通信", "量子精密测量"]

def _join(row, cols):
    parts = []
    for c in cols:
        v = row.get(c, "")
        if pd.notna(v) and str(v).strip():
            parts.append(str(v))
    return " ".join(parts)

def score_company(row, keywords):
    direct = indirect = 0
    domains = {d: 0 for d in DOMAINS}
    for field, w in FIELD_WEIGHTS.items():
        text = _join(row, FIELD_COLS[field])
        kw = keywords[field]
        direct += len([k for k in kw["direct"] if k in text]) * 3 * w
        indirect += len([k for k in kw["indirect"] if k in text]) * 1 * w
        for dom in DOMAINS:
            domains[dom] += len([k for k in kw[dom] if k in text]) * 2 * w
    return direct, indirect, domains

# ==================== 专利索引 + 明细 ====================
@st.cache_data
def build_patent_index(patent_path):
    df = pd.read_excel(patent_path)
    kw = load_patent_keywords(PATENT_KW_FILE)

    app_col = None
    for c in df.columns:
        if "申请" in str(c) and "人" in str(c):
            app_col = c
            break
    if app_col is None:
        app_col = "标题"

    code_col = None
    for c in df.columns:
        if "公开" in str(c) or "公告" in str(c):
            code_col = c
            break

    text_cols = [c for c in ["标题", "摘要", "技术领域"] if c in df.columns]
    df["_text"] = df[text_cols].fillna("").agg(" ".join, axis=1)

    direct_set = set(kw["direct"])
    qdf = df[df["_text"].apply(lambda t: any(k in t for k in direct_set))]

    index = {}
    for _, row in qdf.iterrows():
        c = str(row.get(app_col, "") or "").strip()
        if not c:
            continue
        text = row["_text"]
        title = str(row.get("标题", "") or "").strip()
        code = str(row.get(code_col, "") or "").strip() if code_col else ""
        year = ""
        if code:
            m = pd.Series([code]).str.extract(r"(\d{4})")[0].iloc[0]
            year = m or ""

        if c not in index:
            index[c] = {"count": 0, **{d: 0 for d in DOMAINS}, "patents": []}
        index[c]["count"] += 1
        for dom in DOMAINS:
            if any(k in text for k in kw[dom]):
                index[c][dom] += 1
        # 保留最多 20 条专利明细
        if len(index[c]["patents"]) < 20:
            index[c]["patents"].append({
                "title": title,
                "code": code,
                "year": year,
            })
    return index

def compress(raw, low=30, high=95):
    raw = pd.Series(raw)
    log_v = np.log1p(raw)
    mn, mx = log_v.min(), log_v.max()
    if mx > mn:
        return low + (log_v - mn) / (mx - mn) * (high - low)
    return pd.Series([(low + high) / 2] * len(raw), index=raw.index)

# ==================== 汇总 ====================
@st.cache_data
def compute_all():
    df = pd.read_excel(COMPANY_FILE)
    kw = load_company_keywords(KEYWORD_FILE)
    pat = build_patent_index(PATENT_FILE)

    rows = []
    for idx, row in df.iterrows():
        d, i, doms = score_company(row, kw)
        cname = str(row.get("企业名称", "") or "").strip()
        p = pat.get(cname, {"count": 0, **{d: 0 for d in DOMAINS}, "patents": []})
        p_score = float(np.sqrt(p["count"])) * 10 if p["count"] > 0 else 0.0

        industry = str(row.get("企查查行业中类", "") or "").strip()
        if not industry or industry in ("-", "nan"):
            industry = str(row.get("国民经济行业-中类", "") or "").strip()
        if industry in ("-", "nan", ""):
            industry = str(row.get("国民经济行业-门类", "") or "").strip()

        rows.append({
            "idx": idx,
            "企业名称": cname,
            "省份": str(row.get("所属省份", "") or ""),
            "城市": str(row.get("所属城市", "") or "").replace("市", ""),
            "行业": industry,
            "来觅赛道": str(row.get("来觅赛道(一级市场)", "") or ""),
            "简介": str(row.get("公司简介", "") or ""),
            "经营范围": str(row.get("经营范围", "") or ""),
            # 国标四层
            "门类": str(row.get("国民经济行业-门类", "") or ""),
            "大类": str(row.get("国民经济行业-大类", "") or ""),
            "中类": str(row.get("国民经济行业-中类", "") or ""),
            "小类": str(row.get("国民经济行业-小类", "") or ""),
            "直接分": d,
            "间接分": i,
            "专利数": p["count"],
            "专利侧分": p_score,
            "专利列表": p["patents"],
            "量子计算_raw": doms["量子计算"] + p["量子计算"] * 3,
            "量子通信_raw": doms["量子通信"] + p["量子通信"] * 3,
            "量子精密测量_raw": doms["量子精密测量"] + p["量子精密测量"] * 3,
        })
    res = pd.DataFrame(rows)

    raw_total = res["直接分"] + res["间接分"] * 0.5 + res["专利侧分"] * 1.5
    res["综合分"] = compress(raw_total).round(1)

    for dom in DOMAINS:
        res[dom] = compress(res[f"{dom}_raw"]).round(1)

    res["主领域"] = res[DOMAINS].idxmax(axis=1)

    def verdict(r):
        if r["直接分"] > 0:
            return "高相关"
        if r["间接分"] >= 3:
            return "低相关"
        return "不相关"

    res["相关度"] = res.apply(verdict, axis=1)
    return res

# ==================== 省份排序 ====================
PROVINCE_PY = {
    "北京": "B", "天津": "T", "河北": "H", "山西": "S", "内蒙古": "N",
    "辽宁": "L", "吉林": "J", "黑龙江": "H", "上海": "S", "江苏": "J",
    "浙江": "Z", "安徽": "A", "福建": "F", "江西": "J", "山东": "S",
    "河南": "H", "湖北": "H", "湖南": "H", "广东": "G", "广西": "G",
    "海南": "H", "重庆": "C", "四川": "S", "贵州": "G", "云南": "Y",
    "西藏": "X", "陕西": "S", "甘肃": "G", "青海": "Q", "宁夏": "N",
    "新疆": "X", "台湾": "T", "香港": "X", "澳门": "A",
}

def sort_provinces(provinces):
    def key(p):
        base = p.replace("省", "").replace("市", "").replace("自治区", "").replace("特别行政区", "")
        return (PROVINCE_PY.get(base, "Z"), base)
    return sorted(provinces, key=key)

# ==================== 打分明细 popover ====================
def _detail_html(row):
    h = '<div>'
    h += f'<div class="detail-row"><span class="detail-label">企业侧分</span><span class="detail-val">{row["直接分"]:.0f}</span></div>'
    h += f'<div class="detail-row"><span class="detail-label">间接分</span><span class="detail-val">{row["间接分"]:.0f}</span></div>'
    h += f'<div class="detail-row"><span class="detail-label">量子专利</span><span class="detail-val">{row["专利数"]} 条</span></div>'
    h += '<div style="margin-top:16px;margin-bottom:10px;font-size:13px;font-weight:700;color:#1A2332;">三领域得分</div>'

    for name, val in [("量子计算", row["量子计算"]),
                      ("量子通信", row["量子通信"]),
                      ("量子精密测量", row["量子精密测量"])]:
        h += '<div class="domain-block">'
        h += '<div class="domain-head">'
        h += f'<span class="domain-name">{name}</span>'
        h += f'<span class="domain-score">{val:.1f}</span>'
        h += '</div>'
        h += '<div class="domain-bar-bg">'
        h += f'<div class="domain-bar" style="width:{val}%;"></div>'
        h += '</div>'
        h += '</div>'

    h += '<div style="margin-top:14px;padding-top:10px;border-top:1px solid #F1F5F9;font-size:11.5px;color:#94A3B8;line-height:1.6;">'
    h += '综合分 = log 压缩后（企业侧分 + 间接分×0.5 + 专利侧分×1.5）映射到 30-95'
    h += '</div>'
    h += '</div>'
    return h

# ==================== 企业介绍页 ====================
def render_profile(row):
    """企业详情介绍页"""
    # ---- 基本信息网格 ----
    st.markdown('<div class="profile-block">', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="profile-label">所属地区</div>'
            f'<div class="profile-value">{row["省份"]} · {row["城市"]}</div>',
            unsafe_allow_html=True)
    with c2:
        st.markdown(
            f'<div class="profile-label">行业</div>'
            f'<div class="profile-value">{row["行业"] or "—"}</div>',
            unsafe_allow_html=True)
    with c3:
        st.markdown(
            f'<div class="profile-label">量子专利</div>'
            f'<div class="profile-value">{row["专利数"]} 条</div>',
            unsafe_allow_html=True)
    with c4:
        st.markdown(
            f'<div class="profile-label">主领域</div>'
            f'<div class="profile-value">{row["主领域"]}</div>',
            unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ---- 国标行业层级（可点击展开） ----
    with st.expander("国标行业分类"):
        hier = [
            ("门类", row["门类"]),
            ("大类", row["大类"]),
            ("中类", row["中类"]),
            ("小类", row["小类"]),
        ]
        chain = ""
        for i, (label, val) in enumerate(hier):
            if val and val != "nan" and val != "—":
                chip_cls = "hier-chip hier-chip-active" if i == len(hier) - 1 else "hier-chip"
                chain += f'<span class="{chip_cls}">{label}：{val}</span>'
                if i < len(hier) - 1:
                    chain += '<span class="hier-arrow">›</span>'
        if chain:
            st.markdown(f'<div class="hier-chain">{chain}</div>',
                        unsafe_allow_html=True)
        else:
            st.write("（暂无国标分类）")

    # ---- 简介 ----
    if row["简介"]:
        with st.expander("公司简介", expanded=True):
            st.markdown(
                f'<div style="font-size:13.5px;line-height:1.8;color:#334155;">'
                f'{row["简介"]}</div>',
                unsafe_allow_html=True)

    # ---- 经营范围 ----
    if row["经营范围"]:
        with st.expander("经营范围"):
            st.markdown(
                f'<div style="font-size:13px;line-height:1.8;color:#475569;">'
                f'{row["经营范围"]}</div>',
                unsafe_allow_html=True)

    # ---- 量子相关核心专利 ----
    patents = row["专利列表"]
    if patents:
        with st.expander(f"量子相关核心专利（{row['专利数']} 条，展示前 {len(patents)} 条）",
                         expanded=True):
            for p in patents:
                meta = ""
                if p["code"]:
                    meta += f'<span class="patent-code">{p["code"]}</span>'
                if p["year"]:
                    meta += f'<span>{p["year"]}</span>'
                st.markdown(
                    f'<div class="patent-item">'
                    f'<div class="patent-title">{p["title"]}</div>'
                    f'<div class="patent-meta">{meta}</div>'
                    f'</div>',
                    unsafe_allow_html=True)
    elif row["专利数"] > 0:
        st.caption(f"共 {row['专利数']} 条量子相关专利")

# ==================== 企业卡片 ====================
def render_company_card(row):
    with st.container(border=True):
        c1, c2 = st.columns([5, 1])

        with c1:
            badge_cls = "badge-high" if row["相关度"] == "高相关" else "badge-low"
            st.markdown(
                f'<div class="card-name">{row["企业名称"]}'
                f'<span class="{badge_cls}">{row["相关度"]}</span></div>',
                unsafe_allow_html=True)

            pills = []
            if row["城市"]:
                pills.append(f'<span class="meta-pill">{row["城市"]}</span>')
            if row["行业"]:
                pills.append(f'<span class="meta-pill">{row["行业"]}</span>')
            if row["专利数"] > 0:
                pills.append(f'<span class="meta-pill meta-pill-blue">量子专利 {row["专利数"]} 条</span>')
            st.markdown(
                f'<div>{"".join(pills)}</div>',
                unsafe_allow_html=True)

            if row["来觅赛道"]:
                tags = ""
                for t in str(row["来觅赛道"]).split(",")[:8]:
                    t = t.strip()
                    if t:
                        tags += f'<span class="tag">{t}</span>'
                st.markdown(tags, unsafe_allow_html=True)

        with c2:
            with st.popover(f"{row['综合分']:.1f}", use_container_width=True):
                st.markdown(_detail_html(row), unsafe_allow_html=True)

        # ---- 企业介绍页 ----
        with st.expander("查看企业详情"):
            render_profile(row)

# ==================== 主界面 ====================
def main():
    if "current_domain" not in st.session_state:
        st.session_state.current_domain = None

    with st.spinner("加载数据中..."):
        df_all = compute_all()

    df = df_all[df_all["相关度"].isin(["高相关", "低相关"])].reset_index(drop=True)

    total = len(df)
    qc = int((df["量子计算_raw"] > 0).sum())
    qcom = int((df["量子通信_raw"] > 0).sum())
    qpm = int((df["量子精密测量_raw"] > 0).sum())

    st.markdown(
        '<div class="page-header">'
        '<div class="page-title">量子科技产业地图</div>'
        '<div class="page-sub">企业信息 · 关键词 · 专利数据综合评分</div>'
        '</div>',
        unsafe_allow_html=True)

    left, right = st.columns([1, 3.2], gap="medium")

    with left:
        st.markdown('<div class="group-title">筛选</div>', unsafe_allow_html=True)

        provinces = sort_provinces(
            [p for p in df["省份"].dropna().unique() if p and p != "nan"])
        province = st.selectbox("所属省份", ["全部"] + provinces,
                                 label_visibility="collapsed")

        verdict_sel = st.multiselect("相关度", ["高相关", "低相关"],
                                      default=["高相关", "低相关"],
                                      label_visibility="collapsed")

        st.markdown('<div class="group-title">产业领域</div>',
                    unsafe_allow_html=True)

        current = st.session_state.current_domain

        is_sel = current == "量子科技"
        if st.button(f"量子科技（{total}）", key="btn_all",
                     use_container_width=True,
                     type="primary" if is_sel else "secondary"):
            st.session_state.current_domain = None if is_sel else "量子科技"
            st.rerun()

        c_indent, c_btns = st.columns([0.1, 0.9])
        with c_btns:
            for dom, cnt in [("量子计算", qc), ("量子通信", qcom),
                             ("量子精密测量", qpm)]:
                is_sel = current == dom
                if st.button(f"{dom}（{cnt}）", key=f"btn_{dom}",
                             use_container_width=True,
                             type="primary" if is_sel else "secondary"):
                    st.session_state.current_domain = None if is_sel else dom
                    st.rerun()

    with right:
        current = st.session_state.current_domain

        if current is None:
            st.info("从左侧选择一个领域，查看企业列表")
        else:
            view = df.copy()
            if current != "量子科技":
                view = view[view[f"{current}_raw"] > 0]
            if province != "全部":
                view = view[view["省份"] == province]
            view = view[view["相关度"].isin(verdict_sel)]
            view = view.sort_values("综合分", ascending=False).reset_index(drop=True)

            header = f"{current} · 共 {len(view)} 家"
            if province != "全部":
                header += f" · {province}"
            st.markdown(
                f'<div style="font-size:15px;font-weight:700;color:#1A2332;'
                f'margin-bottom:14px;">{header}</div>',
                unsafe_allow_html=True)

            if view.empty:
                st.info("当前筛选条件下没有匹配企业")
            else:
                for _, row in view.iterrows():
                    render_company_card(row)

if __name__ == "__main__":
    main()
