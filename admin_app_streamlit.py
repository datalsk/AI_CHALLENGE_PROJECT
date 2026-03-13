import streamlit as st
import pandas as pd
import boto3
import json
import urllib.parse
from datetime import datetime

# ==========================================
# 1. 설정 (보안을 위해 secrets 사용)
# ==========================================
# GitHub에 올릴 때는 키를 직접 넣지 않고 streamlit의 secrets 기능을 사용합니다.
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
        return pd.DataFrame(all_data)
    except Exception as e:
        st.sidebar.error(f"S3 데이터 로드 오류: {e}")
        return pd.DataFrame()

# ==========================================
# 2. UI 구성
# ==========================================
st.set_page_config(page_title="관리자 대시보드", layout="wide")
st.title("📊 전사 경비 집계 대시보드")

with st.sidebar:
    st.header("🔍 조회 설정")
    target_date = st.date_input("조회 월 선택", value=datetime(2026, 3, 1)) 
    year_month = target_date.strftime('%Y/%m')
    raw_df = get_all_s3_data(year_month)

if not raw_df.empty:
    team_list = ["전체"] + sorted(raw_df['팀명'].unique().tolist())
    sel_team = st.selectbox("팀 선택", team_list)
    display_df = raw_df if sel_team == "전체" else raw_df[raw_df['팀명'] == sel_team]

    st.subheader(f"📅 {year_month} [{sel_team}] 집계 현황")
    pivot = display_df.pivot_table(index=['팀명', '이름'], columns='항목', values='금액', aggfunc='sum', fill_value=0)
    for cat in FIXED_CATEGORIES:
        if cat not in pivot.columns: pivot[cat] = 0
    pivot = pivot[FIXED_CATEGORIES]
    pivot['합계'] = pivot.sum(axis=1)
    
    total_row = pivot.sum().to_frame().T
    total_row.index = pd.MultiIndex.from_tuples([('전체', '합계')])
    pivot = pd.concat([pivot, total_row])
    st.dataframe(pivot.style.format("{:,.0f}원"), use_container_width=True)

    st.divider()
    st.subheader("🔎 상세 내역 및 증빙 확인")
    sel_user = st.selectbox("조회 대상자 선택", sorted(display_df['이름'].unique()))
    user_detail = display_df[display_df['이름'] == sel_user]
    
    h = st.columns([1, 1.5, 2, 1, 1, 0.5])
    for i, name in enumerate(["항목", "결제일자", "사용처", "금액", "수행일자", "증빙"]):
        h[i].write(f"**{name}**")

    for idx, row in user_detail.iterrows():
        r = st.columns([1, 1.5, 2, 1, 1, 0.5])
        r[0].write(row['항목'])
        r[1].write(row['결제일자'])
        r[2].write(row['사용처'])
        r[3].write(f"{row['금액']:,}원")
        r[4].write(row['수행일자'])
        
        with r[5]:
            if row['증빙URL'] != "N/A":
                with st.popover("🖼️"):
                    try:
                        full_url = row['증빙URL']
                        pure_url = urllib.parse.unquote(full_url)
                        key_start = pure_url.find("images/")
                        if key_start != -1:
                            s3_key = pure_url[key_start:]
                            presigned_url = s3_client.generate_presigned_url(
                                'get_object',
                                Params={'Bucket': S3_BUCKET_NAME, 'Key': s3_key},
                                ExpiresIn=600
                            )
                            st.image(presigned_url, use_container_width=True)
                        else:
                            st.error("이미지 경로 탐색 실패")
                    except Exception as e:
                        st.error("이미지 로드 오류")
            else:
                st.write("❌")
else:
    st.warning("데이터가 없습니다.")