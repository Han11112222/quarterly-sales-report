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
BO_GUP_FILE = "보급률 현황.xlsx"  # 🟢 보급률 현황 파일 정의

# ─────────────────────────────────────────────────────────
# 🟢 코멘트 DB 저장 및 UI 유틸
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
                repo.update_file(contents.path, "Update comments", content_string, contents.sha)
            except:
                repo.create_file(COMMENT_DB_FILE, "Create comments", content_string)
    except: pass

def render_comment_section(title, db_key, curr_db, comments_db, height, placeholder, widget_key):
    st.markdown(f"**{title}**")
    saved_text = curr_db.get(db_key, None)
    if saved_text is not None:
        url_pattern = re.compile(r'(https?://[^\s]+)')
        linked_text = url_pattern.sub(r'<a href="\1" target="_blank" style="color: #2563eb; text-decoration: underline; font-weight: bold;">\1</a>', saved_text)
        formatted_text = linked_text.replace('\n', '<br>')
        st.markdown(f"""<div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-left: 4px solid #1f77b4; padding: 15px; border-radius: 4px; color: #1e40af; font-size: 14.5px; line-height: 1.6; margin-bottom: 10px;">{formatted_text}</div>""", unsafe_allow_html=True)
        with st.expander("🔒 코멘트 수정/삭제 (비밀번호 1234)"):
            pw = st.text_input("비밀번호(PW) 입력", type="password", key=f"pw_{widget_key}")
            if pw == "1234":
                new_text = st.text_area("내용 수정", value=saved_text, height=height, key=f"edit_ta_{widget_key}", label_visibility="collapsed")
                col1, col2 = st.columns(2)
                if col1.button("💾 저장", key=f"edit_save_{widget_key}"):
                    curr_db[db_key] = new_text; save_comments_db(comments_db); st.rerun()
                if col2.button("🗑️ 삭제", key=f"del_{widget_key}"):
                    curr_db.pop(db_key, None); save_comments_db(comments_db); st.rerun()
            elif pw != "": st.error("❌ 비밀번호 오류")
    else:
        input_text = st.text_area("내용 입력", height=height, placeholder=placeholder, key=f"ta_{widget_key}", label_visibility="collapsed")
        if st.button("💾 이 코멘트 저장", key=f"save_{widget_key}"):
            curr_db[db_key] = input_text; save_comments_db(comments_db); st.rerun()

# [중략: 기존 유틸 함수(center_style, clean 등) 동일]
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

# [중략: 기존 로직 함수(make_long, load_all_sheets 등) 동일]
# ...

# ─────────────────────────────────────────────────────────
# 메인 레이아웃 및 본문 로직
# ─────────────────────────────────────────────────────────
# (이전과 동일한 초기화 로직 생략, 실제 적용 시 이전 코드 그대로 유지)

# 🟢 대구시장 보고용 대시보드 내부 부분 수정 (전체 코드 덮어쓰기 권장)
# (상단 생략...)

        if app_mode == "for 대구시장 보고용":
            st.markdown(f"### 🏢 대구시장 보고용 요약 대시보드")
            
            # 1. 연도별 판매량 추이
            st.markdown("#### 1. 연도별 판매량 추이 (2021~2025)")
            df_stack = df_long_rpt[(df_long_rpt["계획/실적"] == "실적") & (df_long_rpt["연"].isin([2021, 2022, 2023, 2024, 2025]))].copy()
            
            if not df_stack.empty:
                # 🟢 그룹핑 재정의
                def remap_stack_group(g):
                    if g in ["가정용", "산업용", "수송용", "업무용", "영업용"]: return g
                    return "기타"
                df_stack["그룹"] = df_stack["그룹"].apply(remap_stack_group)
                
                stack_grp = df_stack.groupby(["연", "그룹"], as_index=False)["값"].sum()
                yearly_totals = stack_grp.groupby("연")["값"].transform("sum")
                stack_grp["비율(%)"] = (stack_grp["값"] / yearly_totals * 100).round(1)
                stack_grp["텍스트"] = stack_grp.apply(lambda x: f"{x['값']:,.0f}<br>({x['비율(%)']}%)" if x['값'] > 0 else "", axis=1)

                fig_stack = px.bar(stack_grp, x="연", y="값", color="그룹", title="2021~2025 그룹별 판매량", text="텍스트")
                fig_stack.update_layout(barmode="stack")
                st.plotly_chart(fig_stack, use_container_width=True)
                
                # 표
                st.markdown("**📊 연도별 그룹 판매량 상세 표**")
                pivot_df = stack_grp.pivot(index="연", columns="그룹", values="값").fillna(0)
                pivot_df["합계"] = pivot_df.sum(axis=1)
                st.dataframe(center_style(pivot_df.style.format("{:,.0f}")), use_container_width=True)

            # 2. 산업용 구성비
            st.markdown("#### 2. 산업용 용도 구성비 (2025년 기준)")
            if not df_csv_tab.empty:
                df_ind = df_csv_tab[(df_csv_tab["상품명"].str.contains("산업용", na=False)) & (df_csv_tab["연_csv"] == 2025)].copy()
                if "업종" in df_ind.columns and not df_ind.empty:
                    df_ind["단순업종"] = df_ind["업종"].apply(lambda x: x if x in ["섬유업종", "펄프업종", "1차금속", "식료품"] else "기타")
                    ind_grp = df_ind.groupby("단순업종", as_index=False)[val_col].sum()
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        fig = px.pie(ind_grp, values=val_col, names="단순업종", hole=0.4)
                        st.plotly_chart(fig, use_container_width=True)
                    with c2:
                        fig_tree = px.treemap(ind_grp, path=["단순업종"], values=val_col)
                        st.plotly_chart(fig_tree, use_container_width=True)

                    # 표 (기타 하단 정렬)
                    st.markdown("**📊 산업용 구성비 상세 표**")
                    ind_grp = ind_grp.sort_values(by=val_col, ascending=False)
                    # 기타를 가장 아래로
                    other_row = ind_grp[ind_grp["단순업종"] == "기타"]
                    main_rows = ind_grp[ind_grp["단순업종"] != "기타"]
                    ind_table = pd.concat([main_rows, other_row])
                    st.dataframe(center_style(ind_table.style.format({val_col: "{:,.0f}"})), use_container_width=True)

            # 3. 보급률 현황
            st.markdown("#### 3. 도시가스 보급률 현황")
            c1, c2 = st.columns(2)
            with c1: render_metric_card("📊", "전체 보급률", "96.8%")
            with c2: render_metric_card("🏙️", "대구시", "97.5%")
            
            # 🟢 상세 그래프 (토글 키를 모드별 고유키로 분리)
            show_detail = st.toggle("🔍 대구시 구군별 상세 보기", key=f"mayor_gu_toggle_{key_sfx}")
            if show_detail:
                try:
                    file_path = Path(__file__).parent / BO_GUP_FILE
                    if file_path.exists():
                        df_rate = pd.read_excel(file_path)
                        st.bar_chart(df_rate.set_index(df_rate.columns[0]))
                    else: st.warning(f"파일 {BO_GUP_FILE}을 찾을 수 없습니다.")
                except: st.error("파일 읽기 오류")
            st.markdown("※ 보급률 = 가정용 청구전수 / 주민등록 세대수")

# [나머지 로직은 이전 코드와 동일하게 유지...]
