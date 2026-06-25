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
from github import Github  # 🟢 GitHub 연동을 위한 라이브러리 추가


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
# 🟢 코멘트 DB 저장 및 UI 유틸 (PW: 1234) - GitHub 실시간 Commit 버전
# ─────────────────────────────────────────────────────────
COMMENT_DB_FILE = "report_comments_db.json"
REPO_NAME = "Han11112222/quarterly-sales-report"  # 🟢 확인된 레포지토리 이름 적용

def load_comments_db():
    if os.path.exists(COMMENT_DB_FILE):
        try:
            with open(COMMENT_DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_comments_db(db_data):
    """
    로컬에 먼저 json을 저장한 뒤, 
    스트림릿 Secrets에 저장된 토큰을 이용해 깃허브 원본 파일도 업데이트합니다.
    """
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
    except Exception as e:
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


# 엑셀 헤더 → 분석 그룹 매핑
USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "일반용": "영업용",
    "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용",
    "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "열전용설비용": "열전용설비용",
}

COLOR_PLAN = "rgba(0, 90, 200, 1)"
COLOR_ACT = "rgba(0, 150, 255, 1)"
COLOR_PREV = "rgba(190, 190, 190, 1)"


# ─────────────────────────────────────────────────────────
# 공통 유틸
# ─────────────────────────────────────────────────────────
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
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)

    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col)
            if group is None: group = keyword_group(col)
            if group is None: continue

            base = df[["연", "월"]].copy()
            base["그룹"] = group
            base["용도"] = col
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)

    if not records: return pd.DataFrame(columns=["연", "월", "그룹", "용도", "계획/실적", "값"])
    long_df = pd.concat(records, ignore_index=True).dropna(subset=["연", "월"])
    long_df["연"] = long_df["연"].astype(int)
    long_df["월"] = long_df["월"].astype(int)
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
    out: Dict[str, pd.DataFrame] = {}
    for name in needed:
        if name in xls.sheet_names:
            out[name] = xls.parse(name)
    return out

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict: Dict[str, pd.DataFrame] = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets):
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets):
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict


def render_metric_card(icon: str, title: str, main: str, sub: str = "", color: str = "#1f77b4"):
    html = f"""
    <div style="background-color:#ffffff; border-radius:22px; padding:24px 26px 20px 26px; box-shadow:0 4px 18px rgba(0,0,0,0.06); height:100%; display:flex; flex-direction:column; justify-content:flex-start;">
        <div style="font-size:44px; line-height:1; margin-bottom:8px;">{icon}</div>
        <div style="font-size:18px; font-weight:650; color:#444; margin-bottom:6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{title}</div>
        <div style="font-size:28px; font-weight:750; color:{color}; margin-bottom:8px; white-space: nowrap; letter-spacing:-0.5px;">{main}</div>
        <div style="font-size:14px; color:#444; min-height:20px; font-weight:500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{sub}</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def render_rate_donut(rate: float, color: str, title: str = ""):
    if pd.isna(rate) or np.isnan(rate):
        st.markdown("<div style='font-size:14px;color:#999;text-align:center;'>데이터 없음</div>", unsafe_allow_html=True)
        return

    filled = max(min(float(rate), 200.0), 0.0)
    empty = max(100.0 - filled, 0.0)

    fig = go.Figure(
        data=[go.Pie(values=[filled, empty], hole=0.7, sort=False, direction="clockwise", marker=dict(colors=[color, "#e5e7eb"]), textinfo="none")]
    )
    fig.update_layout(
        showlegend=False, width=200, height=230, margin=dict(l=0, r=0, t=40, b=0),
        title=dict(text=title, font=dict(size=14, color="#666"), x=0.5, xanchor='center', y=0.98) if title else None,
        annotations=[dict(text=f"{rate:.1f}%", x=0.5, y=0.5, showarrow=False, font=dict(size=22, color=color, family="NanumGothic"))],
    )
    st.plotly_chart(fig, use_container_width=False)


# ─────────────────────────────────────────────────────────
# 메인 레이아웃 (사이드바)
# ─────────────────────────────────────────────────────────
st.title("📊 판매량 분석 보고서")

with st.sidebar:
    st.header("🏢 보고서 모드 설정")
    app_mode = st.radio("조회 모드 선택", ["for Executive", "for Sharing", "for 대구시장 보고용"])
    st.markdown("---")

    st.header("📂 데이터 불러오기")
    st.subheader("1. 판매량 데이터 (필수)")
    src_sales = st.radio("판매량 데이터 소스", ["레포 파일 사용", "엑셀 업로드(.xlsx)"], index=0, key="rpt_sales_src")
    excel_bytes = None
    rpt_base_info = ""
    if src_sales == "엑셀 업로드(.xlsx)":
        up_sales = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="rpt_sales_uploader")
        if up_sales is not None:
            excel_bytes = up_sales.getvalue()
            rpt_base_info = f"소스: 업로드 파일 — {up_sales.name}"
    else:
        path_sales = Path(__file__).parent / DEFAULT_SALES_XLSX
        if path_sales.exists():
            excel_bytes = path_sales.read_bytes()
            rpt_base_info = f"소스: 레포 파일 — {DEFAULT_SALES_XLSX}"
        else:
            rpt_base_info = f"레포 경로에 {DEFAULT_SALES_XLSX} 파일이 없습니다."
    st.caption(rpt_base_info)
    st.markdown("---")

    st.subheader("2. 업종별 상세 (별첨용)")
    src_csv = st.radio("업종별 데이터 소스", ["레포 파일 사용", "CSV 업로드(.csv)"], index=0, key="csv_src")
    csv_bytes = None
    csv_info = ""
    if src_csv == "CSV 업로드(.csv)":
        up_csvs = st.file_uploader("가정용외_*.csv 형식 (다중 업로드 가능)", type=["csv"], accept_multiple_files=True, key="csv_uploader")
        if up_csvs:
            df_list = []
            for f in up_csvs:
                try: df_list.append(pd.read_csv(io.BytesIO(f.getvalue()), encoding="utf-8-sig", thousands=','))
                except:
                    try: df_list.append(pd.read_csv(io.BytesIO(f.getvalue()), encoding="cp949", thousands=','))
                    except: pass
            if df_list:
                st.session_state['merged_csv_df'] = pd.concat(df_list, ignore_index=True)
            csv_info = f"소스: 업로드 파일 {len(up_csvs)}개 병합 완료"
        else:
            if 'merged_csv_df' in st.session_state: del st.session_state['merged_csv_df']
    else:
        path_csv = Path(__file__).parent / DEFAULT_CSV
        if path_csv.exists():
            csv_info = f"소스: 레포 파일 전체 자동 스캔 중..."
        else:
            csv_info = f"레포 경로에 {DEFAULT_CSV} 파일이 없습니다."
    st.caption(csv_info)


# ─────────────────────────────────────────────────────────
# 본문 로직
# ─────────────────────────────────────────────────────────
if app_mode == "for Sharing":
    st.info("🔒 'for Sharing' 모드입니다. 내용을 확인하려면 비밀번호를 입력해주세요.")
    share_pw = st.text_input("접근 비밀번호 (PW)", type="password")
    if share_pw != "1234":
        if share_pw != "": st.error("❌ 비밀번호가 일치하지 않습니다.")
        st.stop()
    else:
        st.success("🔓 인증되었습니다. 공유용 화면을 표시합니다.")


long_dict_rpt: Dict[str, pd.DataFrame] = {}
if 'excel_bytes' in locals() and excel_bytes is not None:
    sheets_rpt = load_all_sheets(excel_bytes)
    long_dict_rpt = build_long_dict(sheets_rpt)
    
df_csv = pd.DataFrame()
if src_csv == "레포 파일 사용":
    repo_dir = Path(__file__).parent
    all_csvs = list(repo_dir.glob("*가정용외*.csv")) + list(repo_dir.glob("가정용외*.csv"))
    all_csvs = list(set(all_csvs)) 
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
    if "사용량(mj)" in df_csv.columns: df_csv["사용량(mj)"] = df_csv["사용량(mj)"].apply(clean_korean_finance_number)
    if "사용량(m3)" in df_csv.columns: df_csv["사용량(m3)"] = df_csv["사용량(m3)"].apply(clean_korean_finance_number)
        
comments_db = load_comments_db()
        
rpt_tabs = st.tabs(["열량 기준 (GJ)", "부피 기준 (천m³)"])

for idx, rpt_tab in enumerate(rpt_tabs):
    with rpt_tab:
        if idx == 0:
            df_long_rpt = long_dict_rpt.get("열량", pd.DataFrame())
            unit_str = "GJ"
            val_col = "사용량(mj)"
            key_sfx = "_gj"
        else:
            df_long_rpt = long_dict_rpt.get("부피", pd.DataFrame())
            unit_str = "천m³"
            val_col = "사용량(m3)"
            key_sfx = "_vol"

        st.markdown(f"#### 📅 보고서 기준 일자 ({app_mode})") 
        
        years_available = [2024, 2025, 2026]
        default_y_index = len(years_available) - 1
        default_q_index = 3 
        
        if not df_long_rpt.empty:
            years_available = sorted(df_long_rpt["연"].unique().tolist())
            actual_data = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["값"] > 0)]
            
            if not actual_data.empty:
                max_year = actual_data["연"].max()
                max_month = actual_data[actual_data["연"] == max_year]["월"].max()
                default_y_index = years_available.index(max_year) if max_year in years_available else len(years_available) - 1
                default_q_index = int((max_month - 1) // 3) 
                
                if default_q_index < 0: default_q_index = 0
                if default_q_index > 3: default_q_index = 3
                
        df_csv_tab = df_csv.copy()
        
        if not df_csv_tab.empty:
            if unit_str == "GJ" and "사용량(mj)" in df_csv_tab.columns:
                df_csv_tab["사용량(mj)"] = df_csv_tab["사용량(mj)"] / 1000.0
            elif unit_str == "천m³" and "사용량(m3)" in df_csv_tab.columns:
                df_csv_tab["사용량(m3)"] = df_csv_tab["사용량(m3)"] / 1000.0
                
            df_csv_tab["날짜_파싱"] = pd.NaT
            date_col = None
            for c in ["청구년월", "매출년월", "년월", "기준년월"]:
                if c in df_csv_tab.columns:
                    date_col = c
                    break
                
            if date_col:
                mask1 = df_csv_tab["날짜_파싱"].isna()
                df_csv_tab.loc[mask1, "날짜_파싱"] = pd.to_datetime(df_csv_tab.loc[mask1, date_col], format="%b-%y", errors="coerce")
                mask2 = df_csv_tab["날짜_파싱"].isna()
                if mask2.any(): df_csv_tab.loc[mask2, "날짜_파싱"] = pd.to_datetime(df_csv_tab.loc[mask2, date_col], format="%Y%m", errors="coerce")
                mask3 = df_csv_tab["날짜_파싱"].isna()
                if mask3.any(): df_csv_tab.loc[mask3, "날짜_파싱"] = pd.to_datetime(df_csv_tab.loc[mask3, date_col], errors="coerce")

            df_csv_tab["연_csv"] = df_csv_tab["날짜_파싱"].dt.year
            df_csv_tab["월_csv"] = df_csv_tab["날짜_파싱"].dt.month
        
        # 🟢 보고서 모드에 따른 상단 UI (연도/분기) 렌더링 분기
        if app_mode != "for 대구시장 보고용":
            c_y, c_q, c_empty = st.columns([1, 1, 2])
            with c_y: sel_year_rpt = st.selectbox("기준 연도", years_available, index=default_y_index, key=f"rpt_yr{key_sfx}")
            with c_q: sel_quarter = st.selectbox("기준 분기", ["1Q (1~3월)", "2Q (1~6월 누적)", "3Q (1~9월 누적)", "4Q (1~12월 누적)"], index=default_q_index, key=f"rpt_qt{key_sfx}")
            max_month = int(sel_quarter[0]) * 3 
        else:
            # 대구시장 보고용 모드에서는 선택창을 숨기고 최신 데이터 기준으로 자동 설정
            sel_year_rpt = years_available[-1] if years_available else 2025
            sel_quarter = "4Q (1~12월 누적)"
            max_month = 12
        
        mode_suffix = "_sharing" if app_mode == "for Sharing" else ("_mayor" if app_mode == "for 대구시장 보고용" else "_executive")
        report_db_key = f"{sel_year_rpt}_{sel_quarter[:2]}_{unit_str}{mode_suffix}"
        
        if report_db_key not in comments_db: comments_db[report_db_key] = {}
        curr_db = comments_db[report_db_key]
        st.markdown("<hr style='margin: 10px 0 30px 0;'>", unsafe_allow_html=True)


    # ─────────────────────────────────────────────────────────
        # 🟢 대구시장 보고용 특화 대시보드 렌더링
        # ─────────────────────────────────────────────────────────
        if app_mode == "for 대구시장 보고용":
            st.markdown(f"### 🏢 대구시장 보고용 요약 대시보드")
            
            # 1. 2021~2025년 연도별 판매량 추이 (스택 그래프)
            st.markdown("#### 1. 연도별 판매량 추이 (2021~2025)")
            df_stack = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["연"].isin([2021, 2022, 2023, 2024, 2025]))]
            
            if not df_stack.empty:
                stack_grp = df_stack.groupby(["연", "그룹"], as_index=False)["값"].sum()
                yearly_totals = stack_grp.groupby("연")["값"].transform("sum")
                stack_grp["비율(%)"] = (stack_grp["값"] / yearly_totals * 100).round(1)
                
                # 값과 비율을 함께 표시하는 텍스트 컬럼 생성
                stack_grp["텍스트"] = stack_grp.apply(lambda x: f"{x['값']:,.0f}<br>({x['비율(%)']}%)" if x['값'] > 0 else "", axis=1)

                fig_stack = px.bar(stack_grp, x="연", y="값", color="그룹", title=f"2021~2025 그룹별 판매량 ({unit_str})", text="텍스트")
                fig_stack.update_layout(xaxis_title="연도", yaxis_title=f"판매량 ({unit_str})", barmode="stack", margin=dict(t=40, b=20, l=20, r=20))
                fig_stack.update_traces(textposition='inside', insidetextanchor='middle')
                st.plotly_chart(fig_stack, use_container_width=True)
                
                # 🟢 스택 그래프 상세 데이터 표
                st.markdown(f"**📊 연도별 그룹 판매량 상세 표 ({unit_str})**")
                stack_pivot = stack_grp.pivot(index="연", columns="그룹", values="값").fillna(0)
                stack_pivot["합계"] = stack_pivot.sum(axis=1)
                stack_pivot = stack_pivot.reset_index().rename(columns={"연": "연도"})
                
                format_dict = {col: "{:,.0f}" for col in stack_pivot.columns if col != "연도"}
                st.dataframe(center_style(stack_pivot.style.format(format_dict)), use_container_width=True, hide_index=True)

                # 요약 박스 (1번 그래프 하단)
                total_2025 = stack_grp[stack_grp["연"] == 2025]["값"].sum() if 2025 in stack_grp["연"].values else stack_grp[stack_grp["연"] == stack_grp["연"].max()]["값"].sum()
                last_year = 2025 if 2025 in stack_grp["연"].values else stack_grp["연"].max()
                st.markdown(f"""
                <div style="background-color: #f8f9fa; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; margin-bottom: 40px; color: #1e40af; font-size: 15px;">
                    <strong>💡 [최근 동향 요약]</strong> {last_year}년 총 판매량은 <strong>{total_2025:,.0f} {unit_str}</strong> 입니다.
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("2021~2025년 실적 데이터가 충분하지 않습니다.")

            # 2. 연도별 산업용 세부 업종 추이
            st.markdown("#### 2. 연도별 산업용 세부 업종 추이")
            if not df_csv_tab.empty and val_col in df_csv_tab.columns:
                csv_products = df_csv_tab["상품명"].astype(str).str.replace(r"\s+", "", regex=True)
                df_ind = df_csv_tab[(csv_products == "산업용")].copy()
                
                if "업종분류" in df_ind.columns and "업종" not in df_ind.columns:
                    df_ind["업종"] = df_ind["업종분류"]
                    
                if "업종" in df_ind.columns and not df_ind.empty:
                    # 주요 업종 단순화 매핑 함수
                    def map_industry_name(name):
                        name = str(name)
                        if "섬유" in name: return "섬유업종"
                        if "펄프" in name or "종이" in name: return "펄프업종"
                        if "1차" in name and "금속" in name: return "1차금속"
                        if "식료품" in name: return "식료품"
                        return "기타"
                        
                    # 전체 데이터에 매핑 적용하여 연도별 추이 분석 가능하도록 설정
                    df_ind["단순업종"] = df_ind["업종"].apply(map_industry_name)
                    
                    # 4번 하단 리스트업(상위 30개)의 독립적 작동을 위해 2025년 실적 데이터 복사본 유지
                    df_ind_2025 = df_ind[df_ind["연_csv"] == 2025].copy()
                    total_ind_val = df_ind_2025[val_col].sum() if not df_ind_2025.empty else 0
                    
                    if not df_ind.empty:
                        # 연도 및 단순업종별 그룹화
                        ind_stack_grp = df_ind.groupby(["연_csv", "단순업종"], as_index=False)[val_col].sum()
                        yearly_ind_totals = ind_stack_grp.groupby("연_csv")[val_col].transform("sum")
                        ind_stack_grp["비율(%)"] = (ind_stack_grp[val_col] / yearly_ind_totals * 100).round(1)
                        
                        # 차트 내부에 노출할 텍스트 포맷팅
                        ind_stack_grp["텍스트"] = ind_stack_grp.apply(lambda x: f"{x[val_col]:,.0f}<br>({x['비율(%)']}%)" if x[val_col] > 0 else "", axis=1)

                        # 파이/트리맵 레이아웃을 걷어내고 넓은 화면으로 스택 그래프 출력
                        fig_ind_stack = px.bar(
                            ind_stack_grp, x="연_csv", y=val_col, color="단순업종",
                            title=f"연도별 산업용 세부 업종 판매량 추이", text="텍스트"
                        )
                        fig_ind_stack.update_layout(xaxis_title="연도", yaxis_title=f"판매량 ({unit_str})", barmode="stack", margin=dict(t=40, b=20, l=20, r=20))
                        fig_ind_stack.update_traces(textposition='inside', insidetextanchor='middle')
                        st.plotly_chart(fig_ind_stack, use_container_width=True)
                        
                        # 🟢 상세 데이터 표 (연도별 구성비를 보여주는 피벗 테이블 구조)
                        st.markdown(f"**📊 연도별 산업용 구성비 상세 표 ({unit_str})**")
                        ind_table = ind_stack_grp.pivot(index="연_csv", columns="단순업종", values=val_col).fillna(0)
                        
                        # 컬럼 정렬 순서 보장 (사용자 요청 분류 순)
                        desired_cols = [c for c in ["섬유업종", "펄프업종", "1차금속", "식료품", "기타"] if c in ind_table.columns]
                        remaining_cols = [c for c in ind_table.columns if c not in desired_cols]
                        ind_table = ind_table[desired_cols + remaining_cols]
                        
                        ind_table["💡 총계"] = ind_table.sum(axis=1)
                        ind_table = ind_table.reset_index().rename(columns={"연_csv": "연도"})
                        
                        format_dict_ind = {col: "{:,.0f}" for col in ind_table.columns if col != "연도"}
                        st.dataframe(center_style(ind_table.style.format(format_dict_ind)), use_container_width=True, hide_index=True)

                        # 최신 연도 데이터를 기준으로 동적 요약 박스 구성
                        latest_year = ind_stack_grp["연_csv"].max()
                        latest_data = ind_stack_grp[ind_stack_grp["연_csv"] == latest_year]
                        total_latest_val = latest_data[val_col].sum()
                        top4_val = latest_data[latest_data["단순업종"] != "기타"][val_col].sum()
                        top4_ratio = (top4_val / total_latest_val * 100) if total_latest_val > 0 else 0
                        
                        st.markdown(f"""
                        <div style="background-color: #f8f9fa; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; margin-bottom: 40px; color: #1e40af; font-size: 15px;">
                            <strong>💡 [산업용 구성 요약]</strong> {latest_year}년 산업용 전체 판매량은 <strong>{total_latest_val:,.0f} {unit_str}</strong>이며, 주요 4대 업종(섬유, 펄프, 1차금속, 식료품)이 전체의 <strong>{top4_ratio:.1f}%</strong> ({top4_val:,.0f} {unit_str})를 점유하고 있습니다.
                        </div>
                        """, unsafe_allow_html=True)
                        
                    else:
                        st.info("산업용 실적 데이터가 존재하지 않습니다.")
                else:
                    st.info("CSV 내 산업용 업종 분류 데이터가 부족합니다.")
            else:
                st.warning("CSV 세부 데이터가 업로드되지 않았습니다. 좌측에서 CSV를 추가해주세요.")
            
            # 3. 도시가스 보급률 현황
            st.markdown("#### 3. 도시가스 보급률 현황")
            
            # 지표 카드 렌더링
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1: render_metric_card("📊", "전체 보급률", "96.8%", "", "#1f77b4")
            with col2: render_metric_card("🏙️", "대구시", "97.5%", "", "#2ca02c")
            with col3: render_metric_card("🏘️", "경산시", "101.3%", "", "#ff7f0e")
            with col4: render_metric_card("⛰️", "고령군", "38.0%", "", "#d62728")
            with col5: render_metric_card("🏞️", "칠곡군", "데이터참조", "※ 파일 연동", "#9467bd")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 🟢 key_sfx를 고유하게 붙여 StreamlitDuplicateElementId 에러 해결
            show_gu_rate = st.toggle("🔍 대구시내 구청별 보급률 상세 보기 (전체 96.8%)", key=f"toggle_gu_rate_{key_sfx}")
            if show_gu_rate:
                # 깃허브에 업로드된 보급률 현황 파일 로드 시도
                try:
                    repo_dir = Path(__file__).parent
                    rate_files = list(repo_dir.glob("*보급률*.csv")) + list(repo_dir.glob("*보급률*.xlsx"))
                    if rate_files:
                        if str(rate_files[0]).endswith('.csv'):
                            df_rate = pd.read_csv(rate_files[0], encoding='utf-8-sig')
                        else:
                            df_rate = pd.read_excel(rate_files[0])
                        st.dataframe(center_style(df_rate.style), use_container_width=True)
                    else:
                        st.info("💡 GitHub 레포지토리에 '보급률 현황' 파일이 인식되면 구청별 상세 내역이 표출됩니다.")
                except Exception as e:
                    st.info("💡 GitHub 레포지토리에 '보급률 현황' 파일이 인식되면 구청별 상세 내역이 표출됩니다.")
            st.markdown("<br><br>", unsafe_allow_html=True)


            # 4. 전체 산업체 상위 30개 업체 리스트 (2025년 기준)
            st.markdown("#### 4. 산업용 상위 30개 업체 리스트 (2025년 기준)")
            if "고객명" in df_ind_2025.columns:
                cust_grp = df_ind_2025.groupby(["고객명", "단순업종"], as_index=False)[val_col].sum().sort_values(val_col, ascending=False)
                top30 = cust_grp.head(30).reset_index(drop=True)
                
                top30_sum = top30[val_col].sum()
                top30_ratio_overall = (top30_sum / total_ind_val * 100) if total_ind_val > 0 else 0
                
                # 소계 추가
                subtotal_row = pd.DataFrame([{"고객명": "💡 소계 (Top 30)", "단순업종": f"전체 산업용 대비 {top30_ratio_overall:.1f}%", val_col: top30_sum}])
                top30_show = pd.concat([top30, subtotal_row], ignore_index=True)
                
                ranks = list(range(1, len(top30) + 1)) + ["-"]
                top30_show.insert(0, "순위", ranks)
                top30_show = top30_show.rename(columns={val_col: f"2025년 판매량 ({unit_str})", "단순업종": "업종"})
                
                st.dataframe(
                    center_style(top30_show.style.format({f"2025년 판매량 ({unit_str})": "{:,.0f}"}).apply(highlight_subtotal, axis=1)), 
                    use_container_width=True, hide_index=True
                )
                
                # 요약 박스 (4번 표 하단)
                st.markdown(f"""
                <div style="background-color: #f8f9fa; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; margin-bottom: 30px; color: #1e40af; font-size: 15px;">
                    <strong>💡 [상위 업체 요약]</strong> 2025년 상위 30개 업체의 총 판매량은 <strong>{top30_sum:,.0f} {unit_str}</strong>으로, 전체 산업용 판매량의 <strong>{top30_ratio_overall:.1f}%</strong>를 차지하고 있습니다.
                </div>
                """, unsafe_allow_html=True)
                
            else:
                st.info("고객명 데이터가 없거나 2025년 실적이 없습니다.")
            
            # 대구시장 보고용 화면 렌더링이 끝나면 아래의 기본 보고서 화면은 스킵
            continue
