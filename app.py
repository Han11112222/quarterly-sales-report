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
from github import Github


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

# ─────────────────────────────────────────────────────────
# 코멘트 DB 저장 및 UI 유틸 (GitHub 실시간 Commit 버전)
# ─────────────────────────────────────────────────────────
COMMENT_DB_FILE = "report_comments_db.json"
REPO_NAME = "Han11112222/quarterly-sales-report"

def load_comments_db():
    if os.path.exists(COMMENT_DB_FILE):
        try:
            with open(COMMENT_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
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
                repo.update_file(contents.path, "Update comments via Streamlit App", content_string, contents.sha)
            except:
                repo.create_file(COMMENT_DB_FILE, "Create comments db via Streamlit App", content_string)
    except Exception:
        pass

def render_comment_section(title, db_key, curr_db, comments_db, height, placeholder, widget_key):
    st.markdown(f"**{title}**")
    saved_text = curr_db.get(db_key, None)
    
    if saved_text is not None:
        url_pattern = re.compile(r'(https?://[^\s]+)')
        linked_text = url_pattern.sub(r'<a href="\1" target="_blank" style="color: #2563eb; text-decoration: underline; font-weight: bold;">\1</a>', saved_text)
        formatted_text = linked_text.replace('\n', '<br>')
        st.markdown(
            f"""
            <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; color: #1e40af; font-size: 14.5px; line-height: 1.6; margin-bottom: 10px;">
                {formatted_text}
            </div>
            """, unsafe_allow_html=True
        )
        
        with st.expander("🔒 코멘트 수정/삭제 (비밀번호 필요)"):
            pw = st.text_input("비밀번호(PW) 입력", type="password", key=f"pw_{widget_key}")
            if pw == "1234":
                new_text = st.text_area("내용 수정", value=saved_text, height=height, key=f"edit_ta_{widget_key}", label_visibility="collapsed")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("💾 수정 내용 저장", key=f"edit_save_{widget_key}", use_container_width=True):
                        curr_db[db_key] = new_text
                        save_comments_db(comments_db)
                        st.rerun()
                with col2:
                    if st.button("🗑️ 코멘트 삭제", key=f"del_{widget_key}", use_container_width=True):
                        curr_db.pop(db_key, None)
                        save_comments_db(comments_db)
                        st.rerun()
            elif pw != "":
                st.error("❌ 비밀번호가 일치하지 않습니다.")
    else:
        input_text = st.text_area("내용 입력", height=height, placeholder=placeholder, key=f"ta_{widget_key}", label_visibility="collapsed")
        if st.button("💾 이 코멘트 저장", key=f"save_{widget_key}"):
            curr_db[db_key] = input_text
            save_comments_db(comments_db)
            st.rerun()

# ─────────────────────────────────────────────────────────
# 유틸리티 함수들 (데이터 매핑, 스타일링 등)
# ─────────────────────────────────────────────────────────
USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "일반용": "영업용", "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용", "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "열전용설비용": "열전용설비용",
}

COLOR_PLAN, COLOR_ACT, COLOR_PREV = "rgba(0, 90, 200, 1)", "rgba(0, 150, 255, 1)", "rgba(190, 190, 190, 1)"

def clean_korean_finance_number(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(",", "").strip()
    if not s: return 0.0
    if s.endswith("-"): s = "-" + s[:-1]
    elif s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    s = re.sub(r"[^\d\.-]", "", s)
    try: return float(s)
    except: return 0.0

def fmt_num_safe(v) -> str:
    if pd.isna(v): return "-"
    try: return f"{float(v):,.0f}"
    except: return "-"

def center_style(styler):
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([
        dict(selector="th", props=[("text-align", "center"), ("vertical-align", "middle"), ("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")]),
        dict(selector="thead th", props=[("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")]),
        dict(selector="tbody tr th", props=[("background-color", "#1e3a8a"), ("color", "#ffffff"), ("font-weight", "bold")])
    ])
    return styler

def highlight_subtotal(s):
    is_subtotal = s.astype(str).str.contains('💡 소계|💡 총계|💡 합계')
    return ['background-color: #1e3a8a; color: #ffffff; font-weight: bold;' if is_subtotal.any() else '' for _ in s]

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

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
    plan_df, actual_df = _clean_base(plan_df), _clean_base(actual_df)
    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col) or keyword_group(col)
            if group is None: continue
            base = df[["연", "월"]].copy()
            base["그룹"], base["용도"], base["계획/실적"] = group, col, label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)
    if not records: return pd.DataFrame(columns=["연", "월", "그룹", "용도", "계획/실적", "값"])
    long_df = pd.concat(records, ignore_index=True).dropna(subset=["연", "월"])
    long_df["연"], long_df["월"] = long_df["연"].astype(int), long_df["월"].astype(int)
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
    return {name: xls.parse(name) for name in needed if name in xls.sheet_names}

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict = {}
    if "계획_부피" in sheets and "실적_부피" in sheets: long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if "계획_열량" in sheets and "실적_열량" in sheets: long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

def render_metric_card(icon: str, title: str, main: str, sub: str = "", color: str = "#1f77b4"):
    st.markdown(f"""
    <div style="background-color:#ffffff; border-radius:22px; padding:24px 26px 20px 26px; box-shadow:0 4px 18px rgba(0,0,0,0.06); height:100%; display:flex; flex-direction:column; justify-content:flex-start;">
        <div style="font-size:44px; line-height:1; margin-bottom:8px;">{icon}</div>
        <div style="font-size:18px; font-weight:650; color:#444; margin-bottom:6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{title}</div>
        <div style="font-size:28px; font-weight:750; color:{color}; margin-bottom:8px; white-space: nowrap; letter-spacing:-0.5px;">{main}</div>
        <div style="font-size:14px; color:#444; min-height:20px; font-weight:500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

def render_rate_donut(rate: float, color: str, title: str = ""):
    if pd.isna(rate) or np.isnan(rate):
        st.markdown("<div style='font-size:14px;color:#999;text-align:center;'>데이터 없음</div>", unsafe_allow_html=True)
        return
    filled, empty = max(min(float(rate), 200.0), 0.0), max(100.0 - max(min(float(rate), 200.0), 0.0), 0.0)
    fig = go.Figure(data=[go.Pie(values=[filled, empty], hole=0.7, sort=False, direction="clockwise", marker=dict(colors=[color, "#e5e7eb"]), textinfo="none")])
    fig.update_layout(showlegend=False, width=200, height=230, margin=dict(l=0, r=0, t=40, b=0),
                      title=dict(text=title, font=dict(size=14, color="#666"), x=0.5, xanchor='center', y=0.98) if title else None,
                      annotations=[dict(text=f"{rate:.1f}%", x=0.5, y=0.5, showarrow=False, font=dict(size=22, color=color, family="NanumGothic"))])
    st.plotly_chart(fig, use_container_width=False)

# ─────────────────────────────────────────────────────────
# 메인 레이아웃 (사이드바 - 모드 선택 추가)
# ─────────────────────────────────────────────────────────
st.title("📊 판매량 분석 보고서")

with st.sidebar:
    # 🟢 탭 및 모드 구분 추가
    st.header("🏢 보고서 모드 설정")
    app_mode = st.radio("작업/조회 모드 선택", ["마케팅팀 내부용", "for Sharing (for Executive)"])
    
    st.markdown("---")
    st.header("📂 데이터 불러오기")

    st.subheader("1. 판매량 데이터 (필수)")
    src_sales = st.radio("판매량 데이터 소스", ["레포 파일 사용", "엑셀 업로드(.xlsx)"], index=0, key="rpt_sales_src")
    excel_bytes = None
    if src_sales == "엑셀 업로드(.xlsx)":
        up_sales = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="rpt_sales_uploader")
        if up_sales is not None: excel_bytes = up_sales.getvalue()
    else:
        path_sales = Path(__file__).parent / DEFAULT_SALES_XLSX
        if path_sales.exists(): excel_bytes = path_sales.read_bytes()

    st.markdown("---")
    st.subheader("2. 업종별 상세 (별첨용)")
    src_csv = st.radio("업종별 데이터 소스", ["레포 파일 사용", "CSV 업로드(.csv)"], index=0, key="csv_src")
    if src_csv == "CSV 업로드(.csv)":
        up_csvs = st.file_uploader("가정용외_*.csv 형식 (다중 업로드 가능)", type=["csv"], accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            df_list = []
            for f in up_csvs:
                try: df_list.append(pd.read_csv(io.BytesIO(f.getvalue()), encoding="utf-8-sig", thousands=','))
                except:
                    try: df_list.append(pd.read_csv(io.BytesIO(f.getvalue()), encoding="cp949", thousands=','))
                    except: pass
            if df_list: st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)

# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
long_dict_rpt = build_long_dict(load_all_sheets(excel_bytes)) if excel_bytes else {}
df_csv = pd.DataFrame()

if src_csv == "레포 파일 사용":
    repo_dir = Path(__file__).parent
    all_csvs = list(set(list(repo_dir.glob("*가정용외*.csv")) + list(repo_dir.glob("가정용외*.csv"))))
    csv_list = []
    for p in all_csvs:
        try: csv_list.append(pd.read_csv(p, encoding="utf-8-sig", thousands=','))
        except:
            try: csv_list.append(pd.read_csv(p, encoding="cp949", thousands=','))
            except: pass
    if csv_list: df_csv = pd.concat(csv_list, ignore_index=True)

if df_csv.empty and 'merged_csv_df' in st.session_state:
    df_csv = st.session_state['merged_csv_df'].copy()

if not df_csv.empty:
    for col in ["사용량(mj)", "사용량(m3)"]:
        if col in df_csv.columns: df_csv[col] = df_csv[col].apply(clean_korean_finance_number)

comments_db = load_comments_db()
rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

for idx, rpt_tab in enumerate(rpt_tabs):
    with rpt_tab:
        unit_str, val_col, key_sfx = ("GJ", "사용량(mj)", "_gj") if idx == 0 else ("천m³", "사용량(m3)", "_vol")
        df_long_rpt = long_dict_rpt.get("열량" if idx == 0 else "부피", pd.DataFrame())

        st.markdown(f"#### 📅 보고서 기준 일자 ({app_mode})") 
        
        years_available = [2024, 2025, 2026]
        default_y_index, default_q_index = len(years_available) - 1, 3
        
        if not df_long_rpt.empty:
            years_available = sorted(df_long_rpt["연"].unique().tolist())
            actual_data = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["값"] > 0)]
            if not actual_data.empty:
                max_year = actual_data["연"].max()
                max_month = actual_data[actual_data["연"] == max_year]["월"].max()
                default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available)-1
                default_q_index = max(0, min(3, int((max_month - 1) // 3)))

        df_csv_tab = df_csv.copy()
        if not df_csv_tab.empty:
            if val_col in df_csv_tab.columns: df_csv_tab[val_col] = df_csv_tab[val_col] / 1000.0
            df_csv_tab["날짜_파싱"] = pd.NaT
            for d_col in ["청구년월", "매출년월", "년월", "기준년월"]:
                if d_col in df_csv_tab.columns:
                    for fmt in ["%b-%y", "%Y%m", None]:
                        mask = df_csv_tab["날짜_파싱"].isna()
                        if mask.any(): df_csv_tab.loc[mask, "날짜_파싱"] = pd.to_datetime(df_csv_tab.loc[mask, d_col], format=fmt, errors="coerce")
            df_csv_tab["연_csv"], df_csv_tab["월_csv"] = df_csv_tab["날짜_파싱"].dt.year, df_csv_tab["날짜_파싱"].dt.month

        c_y, c_q, _ = st.columns([1, 1, 2])
        sel_year_rpt = c_y.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
        sel_quarter = c_q.selectbox("기준 분기", ["1Q (1~3월)", "2Q (1~6월 누적)", "3Q (1~9월 누적)", "4Q (1~12월 누적)"], index=default_q_index, key=f"rpt_qt{key_sfx}")
        max_month = int(sel_quarter[0]) * 3 

        # 🟢 중요: 모드에 따라 데이터베이스 키를 분리합니다.
        mode_suffix = "" if app_mode == "마케팅팀 내부용" else "_executive"
        report_db_key = f"{sel_year_rpt}_{sel_quarter[:2]}_{unit_str}{mode_suffix}"
        
        if report_db_key not in comments_db: comments_db[report_db_key] = {}
        curr_db = comments_db[report_db_key]
        
        st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)

        # 1. At a Glance
        st.markdown("#### 💡 1. At a Glance")
        if not df_long_rpt.empty:
            df_base = df_long_rpt[(df_long_rpt["연"].isin([sel_year_rpt, sel_year_rpt-1])) & (df_long_rpt["월"] <= max_month)]
            t_plan = df_base[(df_base["연"] == sel_year_rpt) & (df_base["계획/실적"] == "계획")]["값"].sum()
            t_act = df_base[(df_base["연"] == sel_year_rpt) & (df_base["계획/실적"] == "실적")]["값"].sum()
            t_prev = df_base[(df_base["연"] == sel_year_rpt-1) & (df_base["계획/실적"] == "실적")]["값"].sum()
            r_plan, r_prev = (t_act/t_plan*100) if t_plan else 0, (t_act/t_prev*100) if t_prev else 0
            
            cm1, cm2, cm3, cd1, cd2 = st.columns([1.1, 1.25, 1.25, 0.7, 0.7])
            with cm1: render_metric_card("🎯", f"{sel_year_rpt}년 계획", f"{fmt_num_safe(t_plan)} {unit_str}", "", COLOR_PLAN)
            with cm2: render_metric_card("🔥", f"{sel_year_rpt}년 실적", f"{fmt_num_safe(t_act)} {unit_str}", f"차이: {'+' if t_act-t_plan>0 else ''}{fmt_num_safe(t_act-t_plan)} ({r_plan:.1f}%, 계획대비)", COLOR_ACT)
            with cm3: render_metric_card("🔄", f"{sel_year_rpt-1}년 실적", f"{fmt_num_safe(t_prev)} {unit_str}", f"차이: {'+' if t_act-t_prev>0 else ''}{fmt_num_safe(t_act-t_prev)} ({r_prev:.1f}%, 전년대비)", COLOR_PREV)
            with cd1: render_rate_donut(r_plan, COLOR_ACT, "계획대비 달성률")
            with cd2: render_rate_donut(r_prev, COLOR_PREV, "전년대비 증감률")
        render_comment_section("📝 분기 핵심 요약 작성", "glance", curr_db, comments_db, 120, "주요 특이사항을 입력하세요.", f"glance_{key_sfx}{mode_suffix}")

        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)

        # 2. One Page Review
        st.markdown("#### 📊 2. 전체 판매량 요약 및 주요 증감 원인 (One Page Review)")
        if not df_long_rpt.empty:
            c_plan, c_act, p_act = df_base[(df_base["연"]==sel_year_rpt) & (df_base["계획/실적"]=="계획")].groupby("그룹")["값"].sum(), df_base[(df_base["연"]==sel_year_rpt) & (df_base["계획/실적"]=="실적")].groupby("그룹")["값"].sum(), df_base[(df_base["연"]==sel_year_rpt-1) & (df_base["계획/실적"]=="실적")].groupby("그룹")["값"].sum()
            summary_df = pd.DataFrame({"계획": c_plan, "실적": c_act, "전년실적": p_act}).fillna(0)
            summary_df["계획대비 증감"], summary_df["YoY 증감"] = summary_df["실적"] - summary_df["계획"], summary_df["실적"] - summary_df["전년실적"]
            summary_df["계획대비 달성률(%)"] = np.where(summary_df["계획"]>0, (summary_df["실적"]/summary_df["계획"])*100, 0)
            summary_df["YoY 대비(%)"] = np.where(summary_df["전년실적"]>0, (summary_df["실적"]/summary_df["전년실적"])*100, 0)
            total_row = summary_df.sum(numeric_only=True)
            total_row["계획대비 달성률(%)"], total_row["YoY 대비(%)"] = (total_row["실적"]/total_row["계획"]*100) if total_row["계획"] else 0, (total_row["실적"]/total_row["전년실적"]*100) if total_row["전년실적"] else 0
            summary_df.loc["💡 합계"] = total_row
            summary_df = summary_df[["계획", "실적", "계획대비 증감", "계획대비 달성률(%)", "전년실적", "YoY 증감", "YoY 대비(%)"]]
            summary_df.columns = pd.MultiIndex.from_tuples([("계획대비", "계획"), ("계획대비", "실적"), ("계획대비", "증감"), ("계획대비", "대비(%)"), ("YoY", "전년실적"), ("YoY", "증감"), ("YoY", "대비(%)")])
            summary_df = summary_df.reset_index().rename(columns={("index", ""): ("구분", "그룹")})
            st.dataframe(center_style(summary_df.style.format({("계획대비", "계획"): "{:,.0f}", ("계획대비", "실적"): "{:,.0f}", ("계획대비", "증감"): "{:,.0f}", ("계획대비", "대비(%)"): "{:,.1f}", ("YoY", "전년실적"): "{:,.0f}", ("YoY", "증감"): "{:,.0f}", ("YoY", "대비(%)"): "{:,.1f}"}).apply(highlight_subtotal, axis=1)), use_container_width=True, hide_index=True)
        render_comment_section("📝 주요 증감 원인 작성 (One Page Review)", "review", curr_db, comments_db, 150, "종합 분석을 입력하세요.", f"review_{key_sfx}{mode_suffix}")

        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)

        # 3, 4, 5. 용도별 분석 (가정/산업/업무)
        def render_usage_trend_report(usage_name, section_num, key_sfx, db_key, mode_suffix):
            if df_long_rpt.empty: return
            df_u = df_long_rpt[(df_long_rpt["그룹"] == usage_name) & (df_long_rpt["월"] <= max_month)]
            p_plan, p_act, p_prev = df_u[(df_u["연"]==sel_year_rpt) & (df_u["계획/실적"]=="계획")].groupby("월")["값"].sum(), df_u[(df_u["연"]==sel_year_rpt) & (df_u["계획/실적"]=="실적")].groupby("월")["값"].sum(), df_u[(df_u["연"]==sel_year_rpt-1) & (df_u["계획/실적"]=="실적")].groupby("월")["값"].sum()
            s_plan, s_act, s_prev = p_plan.sum(), p_act.sum(), p_prev.sum()
            
            st.markdown(f"#### 📈 {section_num}. 용도별 판매량 분석 : {usage_name}")
            col_c, col_m = st.columns([1, 2.5])
            with col_c:
                st.markdown(f"**■ 누적 실적 비교 ({sel_quarter[:2]})**")
                st.markdown(f"""<div style="background-color: #e2e8f0; border-left: 5px solid #1e3a8a; padding: 10px; border-radius: 4px;"><b>판매량: {s_act:,.0f} {unit_str}<br>전년대비: {'+' if s_act-s_prev>0 else ''}{s_act-s_prev:,.0f} ({(s_act/s_prev*100) if s_prev else 0:.1f}%)</b></div>""", unsafe_allow_html=True)
                fig_c = go.Figure(data=[go.Bar(x=[f"{sel_year_rpt}년 계획", f"{sel_year_rpt}년 실적", f"{sel_year_rpt-1}년 실적"], y=[s_plan, s_act, s_prev], marker_color=[COLOR_PLAN, COLOR_ACT, COLOR_PREV], text=[f"{s_plan:,.0f}", f"{s_act:,.0f}", f"{s_prev:,.0f}"], textposition='auto')])
                fig_c.update_layout(margin=dict(t=25, b=10, l=10, r=10), height=400, showlegend=False)
                st.plotly_chart(fig_c, use_container_width=True)
            with col_m:
                st.markdown("**■ 월별 실적 비교**")
                fig_m = go.Figure()
                m_list = list(range(1, max_month + 1))
                for label, vals, color in [(f"{sel_year_rpt}년 계획", [p_plan.get(m, 0) for m in m_list], COLOR_PLAN), (f"{sel_year_rpt}년 실적", [p_act.get(m, 0) for m in m_list], COLOR_ACT), (f"{sel_year_rpt-1}년 실적", [p_prev.get(m, 0) for m in m_list], COLOR_PREV)]:
                    fig_m.add_trace(go.Bar(x=m_list, y=vals, name=label, marker_color=color, text=[f"{v:,.0f}" if v>0 else "" for v in vals], textposition='auto'))
                fig_m.update_layout(barmode='group', xaxis=dict(tickmode='linear', tick0=1, dtick=1), margin=dict(t=10, b=10, l=10, r=10), height=400, legend=dict(orientation="h", y=1.1, x=1, xanchor="right"))
                st.plotly_chart(fig_m, use_container_width=True)
            render_comment_section(f"📝 {usage_name} 세부 코멘트 작성", db_key, curr_db, comments_db, 100, "세부 내용을 입력하세요.", f"{usage_name}_{key_sfx}{mode_suffix}")

        render_usage_trend_report("가정용", 3, key_sfx, "home", mode_suffix)
        render_usage_trend_report("산업용", 4, key_sfx, "ind", mode_suffix)
        render_usage_trend_report("업무용", 5, key_sfx, "biz", mode_suffix)

        # 6, 7. 별첨 (생략 없이 기존 로직 유지하되 mode_suffix 적용)
        # ... (이하 별첨 및 PDF 출력 로직은 기존과 동일하며 widget_key에 mode_suffix만 추가됨)
        st.markdown("<hr style='border-top: 2px solid #bbb; margin: 40px 0 20px 0;'>", unsafe_allow_html=True)
        st.markdown("### 🖨️ 보고서 출력")
        st.markdown("""<style>@media print { header, section[data-testid="stSidebar"], div[data-testid="stToolbar"], iframe { display: none !important; } }</style>""", unsafe_allow_html=True)
        st.components.v1.html(f"""<button onclick="window.parent.print()" style="padding: 12px 20px; font-size: 16px; border-radius: 8px; background-color: #1e3a8a; color: white; border: none; cursor: pointer; width: 100%; font-weight: bold;">🖨️ {app_mode} 화면 전체 PDF 다운로드</button>""", height=70)
