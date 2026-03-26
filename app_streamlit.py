import streamlit as st
import pandas as pd
import boto3
import base64
import requests
import json
import io
from datetime import datetime
from PIL import Image

# ==========================================
# 0. UI 숨기기 설정 (메뉴 및 워터마크 제거)
# ==========================================
st.set_page_config(page_title="AI 경비 제출 시스템", layout="wide")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    /* 배포 버튼 등 툴바 숨기기 */
    .stAppDeployButton {display:none;}
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. 설정 (보안을 위해 st.secrets 사용)
# ==========================================
try:
    OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
    AWS_ACCESS_KEY = st.secrets["AWS_ACCESS_KEY_ID"]
    AWS_SECRET_KEY = st.secrets["AWS_SECRET_ACCESS_KEY"]
    S3_BUCKET = st.secrets["S3_BUCKET_NAME"]
    AWS_REGION = st.secrets["AWS_REGION"]
except KeyError as e:
    st.error(f"비밀 키 설정 오류: {e} 항목이 secrets.toml에 없습니다.")
    st.stop()

# S3 클라이언트 초기화
s3_client = boto3.client(
    's3',
    aws_access_key_id=AWS_ACCESS_KEY,
    aws_secret_access_key=AWS_SECRET_KEY,
    region_name=AWS_REGION
)

# ==========================================
# 2. 유틸리티 함수
# ==========================================

def analyze_receipt(uploaded_file):
    base64_image = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {OPENAI_API_KEY}"}
    
    prompt = """
    영수증에서 다음 정보를 추출해 JSON으로 응답해줘:
    1. '결제 날짜(YYYY-MM-DD)'
    2. '사용처': 상호명이 안 보이면 메뉴나 주소로 상호명을 추론해서 적고 뒤에 (추론)을 붙여줘.
    3. '합계 금액': 숫자 이외의 기호는 제거하고 순수 숫자만 추출해.
    정보를 찾을 수 없으면 '미확인'으로 표시해.
    """
    
    payload = {
        "model": "gpt-4o-mini",
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{uploaded_file.type};base64,{base64_image}"}}
            ]
        }],
        "response_format": { "type": "json_object" }
    }
    try:
        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
        return json.loads(response.json()['choices'][0]['message']['content'])
    except:
        return {"결제 날짜": "에러", "사용처": "분석실패", "합계 금액": 0}

def process_and_save_to_s3(user_name, team_name, day_status, expense_items):
    now = datetime.now()
    date_path = now.strftime('%Y/%m')
    timestamp = now.strftime('%Y%m%d_%H%M%S')
    summary_list = []

    for idx, item in enumerate(expense_items):
        img_url = "N/A"
        if item.get('image_display'):
            clean_name = user_name.strip().replace(" ", "_")
            img_filename = f"{clean_name}_{timestamp}_{idx}.png"
            img_key = f"images/{date_path}/{team_name}/{img_filename}"
            
            img_byte_arr = io.BytesIO()
            item['image_display'].save(img_byte_arr, format='PNG')
            
            s3_client.put_object(
                Bucket=S3_BUCKET, Key=img_key, 
                Body=img_byte_arr.getvalue(), ContentType='image/png'
            )
            img_url = f"https://{S3_BUCKET}.s3.{AWS_REGION}.amazonaws.com/{img_key}"

        summary_list.append({
            "이름": user_name, "팀명": team_name, "항목": item['종류'],
            "금액": item['금액'], "결제일자": item['결제 날짜'],
            "사용처": item['사용처'], "수행일자": day_status, "증빙URL": img_url
        })

    data_key = f"data/{date_path}/{team_name}/{user_name.strip()}_{timestamp}.json"
    s3_client.put_object(
        Bucket=S3_BUCKET, Key=data_key,
        Body=json.dumps(summary_list, ensure_ascii=False).encode('utf-8')
    )
    return True

# ==========================================
# 3. Streamlit UI
# ==========================================

if 'expense_items' not in st.session_state: st.session_state.expense_items = []
if 'selected_cat' not in st.session_state: st.session_state.selected_cat = "야근식대"

with st.sidebar:
    st.header("👤 제출자 정보")
    user_name = st.text_input("성함", placeholder="홍길동")
    team_name = st.selectbox("소속 팀", ["영업1팀", "영업2팀", "개발팀", "인사팀", "마케팅팀", "기타"])
    day_status = "20일 이상" if st.checkbox("월 20일 이상 수행") else ("10일 이상" if st.checkbox("월 10일 이상") else "해당없음")

st.title("📑 AI 경비 제출 시스템")

# 항목 선택 버튼 (콜백 대신 간단한 로직 유지)
categories = ["야근식대", "야근교통비", "외근교통비", "프로젝트비용", "기타"]
cols = st.columns(5)
for i, cat in enumerate(categories):
    if cols[i].button(f"📁 {cat}", use_container_width=True, type="primary" if st.session_state.selected_cat == cat else "secondary"):
        st.session_state.selected_cat = cat
        st.rerun()

st.divider()
c1, c2 = st.columns([4, 1])
with c1:
    uploaded_files = st.file_uploader(f"[{st.session_state.selected_cat}] 영수증 업로드", accept_multiple_files=True)
with c2:
    st.write("")
    if st.button("➕ 직접 추가"):
        st.session_state.expense_items.append({
            "종류": st.session_state.selected_cat, "결제 날짜": datetime.now().strftime("%Y-%m-%d"),
            "사용처": "", "금액": 0, "image_display": None
        })

if uploaded_files and st.button(f"✨ {len(uploaded_files)}건 AI 분석 시작", type="primary"):
    with st.spinner("AI가 분석 중입니다..."):
        for f in uploaded_files:
            res = analyze_receipt(f)
            img = Image.open(f)
            img.thumbnail((500, 500))
            st.session_state.expense_items.append({
                "종류": st.session_state.selected_cat, "결제 날짜": res.get("결제 날짜"),
                "사용처": res.get("사용처"), "금액": res.get("합계 금액"), "image_display": img
            })
    st.rerun()

if st.session_state.expense_items:
    st.divider()
    for idx, item in enumerate(st.session_state.expense_items):
        r = st.columns([1, 1.5, 2, 1, 0.5, 0.5])
        item['종류'] = r[0].selectbox(f"cat_{idx}", categories, index=categories.index(item['종류']), label_visibility="collapsed")
        item['결제 날짜'] = r[1].text_input(f"dt_{idx}", item['결제 날짜'], label_visibility="collapsed")
        item['사용처'] = r[2].text_input(f"vn_{idx}", item['사용처'], label_visibility="collapsed")
        item['금액'] = r[3].number_input(f"am_{idx}", value=int(item['금액']), label_visibility="collapsed")
        with r[4]:
            if item['image_display']:
                with st.popover("🖼️"): st.image(item['image_display'])
        if r[5].button("🗑️", key=f"del_{idx}"):
            st.session_state.expense_items.pop(idx)
            st.rerun()

    # [수정된 제출 로직]
    if st.button("🚀 서버로 최종 제출", type="primary", use_container_width=True):
        if not user_name: 
            st.error("성함을 입력해주세요.")
        else:
            with st.spinner("서버로 전송 중..."):
                success = process_and_save_to_s3(user_name, team_name, day_status, st.session_state.expense_items)
            
            if success:
                st.balloons()
                st.success(f"🎉 {user_name}님의 내역이 성공적으로 제출되었습니다!")
                # 상태 초기화
                st.session_state.expense_items = []
                # 잠시 대기 후 화면 갱신 (사용자가 성공 메시지를 볼 시간 확보)
                import time
                time.sleep(2)
                st.rerun()