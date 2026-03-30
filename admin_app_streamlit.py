import streamlit as st
import pandas as pd
import boto3
import json
import urllib.parse
from datetime import datetime

# ==========================================
# 0. UI 설정 및 모던 SaaS 디자인 CSS 적용
# ==========================================
st.set_page_config(page_title="관리자 대시보드", layout="wide")

st.markdown("""
    <style>
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/variable/pretendardvariable.css");
    html, body, [class*="css"], .stMarkdown, .stText, button, input, select, .stDataFrame {
        font-family: 'Pretendard Variable', Pretendard, -apple-system, sans-serif !important;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display:none;}
    header {background-color: transparent !important;}

    /* 모던 카드 UI */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        box-shadow: rgba(0, 0, 0, 0.03) 0px 4px 10px !important;
        border: 1px solid rgba(226, 232, 240, 0.8) !important;
        background-color: #ffffff !important;
        padding: 8px;
        transition: all 0.2s ease;
    }
    
    /* 주요 버튼 */
    .stButton > button[kind="primary"] {
        background-color: #4f46e5;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #3730a3;
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
    }
    
    /* 일반 버튼 (영수증 보기 등) */
    .stButton > button[kind="secondary"] {
        border-radius: 6px;
        font-weight: 500;
        font-size: 13px !important;
        border: 1px solid rgba(148, 163, 184, 0.3);
        background-color: transparent;
        padding: 2px 10px;
    }
    .stButton > button[kind="secondary"]:hover {
        background-color: rgba(148, 163, 184, 0.1);
    }
    
    h1 { font-weight: 700 !important; letter-spacing: -1px; margin-bottom: 0px !important;}
    h3 { font-weight: 600 !important; letter-spacing: -0.5px; }
    
    /* 입력 폼 라운딩 */
    div[data-baseweb="input"], div[data-baseweb="select"] {
        border-radius: 8px !important;
    }
    
    /* 데이터프레임 헤더 스타일링 */
    th { background-color: #f8fafc !important; color: #475569 !important; font-weight: 600 !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. 설정 및 AWS S3 연동
# ==========================================
try:
    AWS_ACCESS_KEY_ID = st.secrets["AWS_ACCESS_KEY_ID"]
    AWS_SECRET_ACCESS_KEY = st.secrets["AWS_SECRET_ACCESS_KEY"]
    S3_BUCKET_NAME = st.secrets["S3_BUCKET_NAME"]
    AWS_REGION = st.secrets["AWS_REGION"]
except:
    st.error("설정 파일(secrets.toml)이 없습니다. 로컬 테스트 시 설정을 확인하세요.")
    st.stop()

FIXED_CATEGORIES = ["야근식대", "야근교통비", "외근교통비", "프로젝트비용", "기타"]

s3_client = boto3.client(
    's3', 
    aws_access_key_id=AWS_ACCESS_KEY_ID, 
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY, 
    region_name=AWS_REGION
)

@st.cache_data(ttl=60) # 성능 향상을 위해 60초 캐싱
def get_all_s3_data(year_month_path):
    all_data = []
    prefix = f"data/{year_month_path}/"
    try:
        paginator = s3_client.get_paginator('list_objects_v2')
        for page in paginator.paginate(Bucket=S3_BUCKET_NAME, Prefix=prefix):
            if 'Contents' in page:
                for obj in page['Contents']:
                    if obj['Key'].endswith('.json'):
                        content = s3_client.get_object(Bucket=S3_BUCKET_NAME, Key=obj['Key'])['Body'].read().decode('utf-8')
                        all_data.extend(json.loads(content))
        
        df = pd.DataFrame(all_data)
        # 구버전 데이터 호환성을 위해 누락된 컬럼 안전하게 추가
        if not df.empty:
            if '비고' not in df.columns: df['비고'] = ""
            if '배달비_증빙URL' not in df.columns: df['배달비_증빙URL'] = ""
        return df
    except Exception as e:
        st.sidebar.error(f"S3 데이터 로드 오류: {e}")
        return pd.DataFrame()

# Pre-signed URL 생성 헬퍼 함수
def get_presigned_url(full_url):
    if not full_url or str(full_url).strip() in ["", "N/A", "nan"]: 
        return None
    try:
        pure_url = urllib.parse.unquote(full_url)
        key_start = pure_url.find("images/")
        if key_start != -1:
            s3_key = pure_url[key_start:]
            return s3_client.generate_presigned_url(
                'get_object',
                Params={'Bucket': S3_BUCKET_NAME, 'Key': s3_key},
                ExpiresIn=600
            )
    except: pass
    return None

# ==========================================
# 2. 메인 화면 구성
# ==========================================
st.title("경비 정산 관리자 대시보드")
st.markdown("<p style='color: #64748b; font-size: 15px; margin-bottom: 2rem;'>임직원이 제출한 경비 내역과 증빙 자료를 검토하고 집계합니다.</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h3 style='margin-bottom: 0.5rem;'>조회 설정</h3>", unsafe_allow_html=True)
    target_date = st.date_input("조회 월 선택", value=datetime.today()) 
    year_month = target_date.strftime('%Y/%m')
    
    with st.spinner("데이터를 불러오는 중..."):
        raw_df = get_all_s3_data(year_month)

if not raw_df.empty:
    team_list = ["전체"] + sorted(raw_df['팀명'].dropna().unique().tolist())
    sel_team = st.sidebar.selectbox("팀 선택", team_list)
    display_df = raw_df if sel_team == "전체" else raw_df[raw_df['팀명'] == sel_team]

    # --- 1. 요약 집계 표 (Pivot) ---
    st.markdown(f"<h3 style='margin-bottom: 1rem;'>{year_month} [{sel_team}] 집계 현황</h3>", unsafe_allow_html=True)
    
    with st.container(border=True):
        pivot = display_df.pivot_table(index=['팀명', '이름'], columns='항목', values='금액', aggfunc='sum', fill_value=0)
        for cat in FIXED_CATEGORIES:
            if cat not in pivot.columns: pivot[cat] = 0
        pivot = pivot[FIXED_CATEGORIES]
        pivot['합계'] = pivot.sum(axis=1)
        
        total_row = pivot.sum().to_frame().T
        total_row.index = pd.MultiIndex.from_tuples([('전체', '합계')])
        pivot = pd.concat([pivot, total_row])
        
        # 데이터프레임 출력
        st.dataframe(pivot.style.format("{:,.0f}원"), use_container_width=True)

    st.markdown("<hr style='margin: 2rem 0; border-top: 1px solid rgba(148, 163, 184, 0.2);'>", unsafe_allow_html=True)

    # --- 2. 상세 내역 및 증빙 검토 ---
    st.markdown("<h3 style='margin-bottom: 1rem;'>상세 내역 및 증빙 검토</h3>", unsafe_allow_html=True)
    
    c1, c2 = st.columns([1, 3])
    sel_user = c1.selectbox("조회 대상자 선택", sorted(display_df['이름'].dropna().unique()), label_visibility="collapsed")
    user_detail = display_df[display_df['이름'] == sel_user]
    
    # 수행일자(프로젝트 기간 등) 정보 표시
    user_proj_dates = user_detail['수행일자'].unique()
    proj_info = user_proj_dates[0] if len(user_proj_dates) > 0 else "정보 없음"
    c2.markdown(f"<div style='padding-top:8px; color:#64748b; font-size:14px;'>📌 <b>프로젝트/수행 기간:</b> {proj_info}</div>", unsafe_allow_html=True)
    
    st.write("")
    
    # 컬럼 헤더
    h = st.columns([1, 1.2, 1.8, 2.5, 1, 1])
    headers = ["항목", "결제일자", "사용처", "비고 (동석자/기타)", "금액", "증빙 자료"]
    for i, name in enumerate(headers):
        h[i].markdown(f"<div style='font-size:13px; font-weight:600; color:#475569; padding-bottom:8px; border-bottom:2px solid #e2e8f0; margin-bottom:8px;'>{name}</div>", unsafe_allow_html=True)

    # 상세 데이터 행 (카드 스타일 적용)
    for idx, row in user_detail.iterrows():
        with st.container(border=True):
            r = st.columns([1, 1.2, 1.8, 2.5, 1, 1])
            
            # 1~5 텍스트 데이터
            r[0].markdown(f"<div style='font-size:14px; font-weight:500; margin-top:6px;'>{row.get('항목', '-')}</div>", unsafe_allow_html=True)
            r[1].markdown(f"<div style='font-size:14px; color:#475569; margin-top:6px;'>{row.get('결제일자', '-')}</div>", unsafe_allow_html=True)
            r[2].markdown(f"<div style='font-size:14px; margin-top:6px;'>{row.get('사용처', '-')}</div>", unsafe_allow_html=True)
            
            note_text = row.get('비고', '')
            r[3].markdown(f"<div style='font-size:13px; color:#64748b; margin-top:6px;'>{note_text if note_text else '-'}</div>", unsafe_allow_html=True)
            
            r[4].markdown(f"<div style='font-size:15px; font-weight:600; color:#0f172a; margin-top:6px;'>{row.get('금액', 0):,} 원</div>", unsafe_allow_html=True)
            
            # 6. 증빙 자료 버튼 (메인 + 배달)
            with r[5]:
                btn_cols = st.columns(2)
                
                # 메인 영수증
                main_url = get_presigned_url(row.get('증빙URL'))
                if main_url:
                    with btn_cols[0]:
                        with st.popover("영수증"):
                            st.image(main_url, width=400)
                            
                # 배달비 영수증
                del_url = get_presigned_url(row.get('배달비_증빙URL'))
                if del_url:
                    with btn_cols[1]:
                        with st.popover("배달"):
                            st.image(del_url, width=400)

else:
    st.info("해당 월에 제출된 정산 데이터가 없습니다.", icon="📂")