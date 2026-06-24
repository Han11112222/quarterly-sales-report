import io
import json
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib as mpl
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from github import Github  # 🟢 GitHub 연동용

# ─────────────────────────────────────────────────────────
# 기본 설정
# ─────────────────────────────────────────────────────────
def set_korean_font():
    ttf = Path(__file__).parent / "NanumGothic-Regular.ttf"
    if ttf.exists():
        try:
            mpl.font_manager.fontManager.addfont(str(ttf))
            mpl.rcParams["font.family"] = "NanumGothic"
            mpl.rcParams["axes.unicode_minus"] = False
        except Exception:
            pass

set_korean_font()
st.set_page_config(page_title="도시가스 판매량 분석 보고서", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"
DEFAULT_CSV = "가정용외_202601.csv"
BO_GUP_FILE = "보급률 현황.xlsx"

# ─────────────────────────────────────────────────────────
# 🟢 DB 및 유틸
# ─────────────────────────────────────────────────────────
COMMENT_DB_FILE = "report_comments_db.json"
REPO_NAME = "Han11112222/quarterly-sales-report"

def load_comments_db():
    if os.path.exists(COMMENT_DB_FILE):
        try:
            with open(COMMENT_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {}
    return {}

def save_comments_db(db_data):
    with open(COMMENT_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db_data, f, ensure_ascii=False, indent=4)
    try:
        if "GITHUB_TOKEN" in st.secrets:
            token = st.secrets["GITHUB_TOKEN"]
            g = Github(token)
            repo = g.get_repo(REPO_NAME)
            content_string = json.dumps(db_data, ensure_ascii=False, indent=4)
            try:
                contents = repo.get_contents(COMMENT_DB_FILE)
                repo.update_file(contents.path, "Update", content_string, contents.sha)
            except:
                repo.create_file(COMMENT_DB_FILE, "Create", content_string)
    except: pass

def render_comment_section(title, db_key, curr_db, comments_db, height, placeholder, widget_key):
    st.markdown(f"**{title}**")
    saved_text = curr_db.get(db_key, None)
    if saved_text is not None:
        formatted_text = saved_text.replace('\n', '<br>')
        st.markdown(f"""<div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #1f77b4;">{formatted_text}</div>""", unsafe_allow_html=True)
        with st.expander("🔒 수정/삭제"):
            pw = st.text_input("PW", type="password", key=f"pw_{widget_key}")
            if pw == "1234":
                new_text = st.text_area("수정", value=saved_text, height=height, key=f"edit_ta_{widget_key}")
                c1, c2 = st.columns(2)
                if c1.button("💾 저장", key=f"edit_save_{widget_key}"):
                    curr_db[db_key] = new_text; save_comments_db(comments_db); st.rerun()
                if c2.button("🗑️ 삭제", key=f"del_{widget_key}"):
                    curr_db.pop(db_key, None); save_comments_db(comments_db); st.rerun()
    else:
        input_text = st.text_area("내용", height=height, placeholder=placeholder, key=f"ta_{widget_key}")
        if st.button("💾 저장", key=f"save_{widget_key}"):
            curr_db[db_key] = input_text; save_comments_db(comments_db); st.rerun()

def center_style(styler):
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([
        dict(selector="th", props=[("text-align", "center"), ("background-color", "#1e3a8a"), ("color", "#ffffff")]),
        dict(selector="tbody tr th", props=[("background-color", "#1e3a8a"), ("color", "#ffffff")])
    ])
    return styler

def highlight_subtotal(s):
    is_subtotal = s.astype(str).str.contains('💡 소계|💡 총계|💡 합계')
    return ['background-color: #1e3a8a; color: #ffffff; font-weight: bold;' if is_subtotal.any() else '' for _ in s]

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

USE_COL_TO_GROUP = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "일반용": "영업용", "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용", "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "열전용설비용": "열전용설비용",
}

def keyword_group(col: str) -> Optional[str]:
    c = str(col)
    if "열병합" in c: return "열병합"
    if "연료전지" in c: return "연료전지"
    if "수송용" in c: return "수송용"
    if "열전용" in c: return "열전용설비용"
    if c in ["산업용"]: return "산업용"
    if c in ["일반용"]: return "영업용"
    if any(k in c for k in ["취사용", "난방용", "자가열"]): return "가정용"
    if any(k in c for k in ["업무", "냉방", "주한미군"]): return "업무용"
    return None

def make_long(plan_df: pd.DataFrame, actual_df: pd.DataFrame) -> pd.DataFrame:
    plan_df = _clean_base(plan_df); actual_df = _clean_base(actual_df)
    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col) or keyword_group(col)
            if group is None: continue
            base = df[["연", "월"]].copy()
            base["그룹"] = group; base["용도"] = col; base["계획/실적"] = label; base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)
    if not records: return pd.DataFrame(columns=["연", "월", "그룹", "용도", "계획/실적", "값"])
    long_df = pd.concat(records, ignore_index=True).dropna(subset=["연", "월"])
    long_df["연"] = long_df["연"].astype(int); long_df["월"] = long_df["월"].astype(int)
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    out = {}
    for name in ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]:
        if name in xls.sheet_names: out[name] = xls.parse(name)
    return out

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets): long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets): long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

def render_metric_card(icon: str, title: str, main: str, sub: str = "", color: str = "#1f77b4"):
    st.markdown(f"""<div style="background-color:#ffffff; border-radius:22px; padding:24px; box-shadow:0 4px 18px rgba(0,0,0,0.06); height:100%;">
        <div style="font-size:30px;">{icon}</div><div style="font-size:16px;">{title}</div><div style="font-size:24px; color:{color}; font-weight:700;">{main}</div></div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
# (중략 - 사이드바 등 기존 로직 유지)
# ...

# 🟢 대구시장 보고용 탭 메인
if app_mode == "for 대구시장 보고용":
    st.markdown("### 🏢 대구시장 보고용 요약 대시보드")
    
    # 1. 2021~2025 연도별 판매량 추이
    st.markdown("#### 1. 연도별 판매량 추이 (2021~2025)")
    all_years = sorted(df_long_rpt["연"].unique().tolist()) if not df_long_rpt.empty else [2021, 2022, 2023, 2024, 2025]
    sel_years = st.multiselect("연도 선택", options=all_years, default=[y for y in [2021, 2022, 2023, 2024, 2025] if y in all_years], key="stack_yrs_sel")
    
    df_stack = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["연"].isin(sel_years))].copy()
    if not df_stack.empty:
        # 그룹 세분화 (가정/산업/수송/업무/영업/기타)
        df_stack["그룹"] = df_stack["그룹"].apply(lambda g: g if g in ["가정용", "산업용", "수송용", "업무용", "영업용"] else "기타")
        stack_grp = df_stack.groupby(["연", "그룹"], as_index=False)["값"].sum()
        yearly_totals = stack_grp.groupby("연")["값"].transform("sum")
        stack_grp["비율(%)"] = (stack_grp["값"] / yearly_totals * 100).round(1)
        stack_grp["텍스트"] = stack_grp.apply(lambda x: f"{x['값']:,.0f}<br>({x['비율(%)']}%)" if (x['값'] > 0 and x['그룹'] != "기타") else "", axis=1)

        fig_stack = px.bar(stack_grp, x="연", y="값", color="그룹", text="텍스트")
        fig_stack.update_layout(barmode="stack", title="연도별 그룹 판매량")
        fig_stack.update_traces(textposition='inside', insidetextanchor='middle', textfont=dict(size=14))
        for year in sel_years:
            total = stack_grp[stack_grp["연"] == year]["값"].sum()
            fig_stack.add_annotation(x=year, y=total, text=f"<b>총 {total:,.0f}</b>", showarrow=False, yshift=15)
        st.plotly_chart(fig_stack, use_container_width=True)
        
        st.markdown(f"**📊 상세 표 ({unit_str})**")
        piv = stack_grp.pivot(index="연", columns="그룹", values="값").fillna(0)
        piv["합계"] = piv.sum(axis=1)
        st.dataframe(center_style(piv.style.format("{:,.0f}")), use_container_width=True)

    # 2. 산업용 용도 구성비
    st.markdown("#### 2. 산업용 용도 구성비 (2025년 기준)")
    if not df_csv_tab.empty:
        df_ind = df_csv_tab[(df_csv_tab["상품명"].str.contains("산업용", na=False)) & (df_csv_tab["연_csv"] == 2025)].copy()
        if not df_ind.empty:
            df_ind["단순업종"] = df_ind["업종"].apply(lambda x: x if x in ["섬유업종", "펄프업종", "1차금속", "식료품"] else "기타")
            ind_grp = df_ind.groupby("단순업종", as_index=False)[val_col].sum()
            total_ind = ind_grp[val_col].sum()
            
            c1, c2, c3 = st.columns([1, 2, 1])
            with c2:
                fig = px.pie(ind_grp, values=val_col, names="단순업종", hole=0.5)
                fig.add_annotation(text=f"<b>총량<br>{total_ind:,.0f}</b>", showarrow=False, font_size=16)
                fig.update_traces(texttemplate="%{label}<br>%{value:,.0f}<br>(%{percent})", textfont=dict(size=14))
                st.plotly_chart(fig, use_container_width=True)
            
            # 기타를 식료품 하단으로(총계 제외하고 가장 아래)
            df_main = ind_grp[ind_grp["단순업종"] != "기타"].sort_values(by=val_col, ascending=False)
            df_others = ind_grp[ind_grp["단순업종"] == "기타"]
            df_final = pd.concat([df_main, df_others]).reset_index(drop=True)
            df_final["비율(%)"] = (df_final[val_col] / total_ind * 100)
            df_final = pd.concat([df_final, pd.DataFrame([{"단순업종": "💡 총계", val_col: total_ind, "비율(%)": 100.0}])])
            st.dataframe(center_style(df_final.style.format({val_col: "{:,.0f}", "비율(%)": "{:.1f}"}).apply(highlight_subtotal, axis=1)), use_container_width=True)

    # 3. 보급률
    st.markdown("#### 3. 도시가스 보급률 현황")
    c1, c2 = st.columns(2)
    with c1: render_metric_card("📊", "전체 보급률", "96.8%")
    with c2: render_metric_card("🏙️", "대구시 보급률", "97.5%")
    
    try:
        df_rate = pd.read_excel(BO_GUP_FILE, header=3).iloc[:, [0, 3]].dropna()
        df_rate.columns = ["구군명", "보급률"]
        df_rate["보급률"] = pd.to_numeric(df_rate["보급률"].astype(str).str.replace('%', ''), errors='coerce')
        st.bar_chart(df_rate.set_index("구군명")["보급률"])
    except: st.info("보급률 데이터를 불러올 수 없습니다.")
    st.markdown("※ 보급률 = 가정용 청구전수 / 주민등록 세대수")
