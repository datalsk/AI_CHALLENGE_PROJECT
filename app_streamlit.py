import streamlit as st
import pandas as pd
import boto3
import base64
import requests
import json
import io
import time
import calendar
from datetime import datetime
from PIL import Image

# ==========================================
# 0. UI 설정 및 상태 관리
# ==========================================
st.set_page_config(page_title="AI 경비 제출 시스템", layout="wide")
st.markdown("<style>#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;} .stAppDeployButton {display:none;}</style>", unsafe_allow_html=True)

if 'expense_items' not in st.session_state: st.session_state.expense_items = []
if 'selected_cat' not in st.session_state: st.session_state.selected_cat = "야근식대"
if 'file_cat_map' not in st.session_state: st.session_state.file_cat_map = {}
if 'submitted' not in st.session_state: st.session_state.submitted = False
if 'uploader_key' not in st.session_state: st.session_state.uploader_key = 0

def change_category(cat_name):
    st.session_state.selected_cat = cat_name

# ==========================================
# 1. 유틸리티 함수
# ==========================================
def safe_int(value):
    try:
        if isinstance(value, str):
            clean_val = "".join(filter(lambda x: x.isdigit() or x == '-', value))
            return abs(int(clean_val)) if clean_val else 0
        return abs(int(value)) if value is not None else 0
    except: return 0

def analyze_receipt(uploaded_file):
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except KeyError:
        st.error("Secrets에 OPENAI_API_KEY가 없습니다.")
        return {"결제 날짜": "에러", "사용처": "키 없음", "합계 금액": 0}

    base64_image = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    prompt = "영수증 이미지에서 '결제 날짜(YYYY-MM-DD)', '사용처', '합계 금액'을 추출해 JSON 응답해줘. 음수 금액은 무시하고 최종 합계만 가져와."
    payload = {"model": "gpt-4o-mini", "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{uploaded_file.type};base64,{base64_image}"}}]}], "response_format": { "type": "json_object" }}
    
    try:
        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=45)
        res_data = json.loads(response.json()['choices'][0]['message']['content'])
        res_data["is_uncertain"] = (res_data.get("사용처") == "미확인" or safe_int(res_data.get("합계 금액")) == 0)
        return res_data
    except: 
        return {"결제 날짜": datetime.now().strftime("%Y-%m-%d"), "사용처": "분석 실패", "합계 금액": 0, "is_uncertain": True}

def save_to_s3(user_name, team_name, day_status, expense_items):
    now = datetime.now()
    date_path = now.strftime('%Y/%m')
    timestamp = now.strftime('%Y%m%d_%H%M%S')
    summary_list = []
    
    s3_bucket = st.secrets["S3_BUCKET_NAME"]
    aws_region = st.secrets["AWS_REGION"]
    s3_client = boto3.client('s3', aws_access_key_id=st.secrets["AWS_ACCESS_KEY_ID"], aws_secret_access_key=st.secrets["AWS_SECRET_ACCESS_KEY"], region_name=aws_region)

    for idx, item in enumerate(expense_items):
        # [핵심 수정] UI에서 계산된 '실제 반영 금액'을 가져옵니다.
        final_amt = item.get('_effective_cost', 0)
        
        # 한도 초과로 전액 깎여버린(0원) 프로젝트 영수증은 S3 저장에서 깔끔하게 스킵(삭제)
        if final_amt == 0 and item['종류'] == "프로젝트비용":
            continue

        img_url = "N/A"
        if item.get('image_display'):
            img_key = f"images/{date_path}/{team_name}/{user_name}_{timestamp}_{idx}.png"
            img_byte_arr = io.BytesIO()
            item['image_display'].save(img_byte_arr, format='PNG')
            s3_client.put_object(Bucket=s3_bucket, Key=img_key, Body=img_byte_arr.getvalue(), ContentType='image/png')
            img_url = f"https://{s3_bucket}.s3.{aws_region}.amazonaws.com/{img_key}"
            
        summary_list.append({
            "이름": user_name, "팀명": team_name, "항목": item['종류'], 
            "금액": final_amt, "결제일자": item['결제일자'], 
            "사용처": item['사용처'], "수행일자": day_status, 
            "비고": item.get('비고', ""), "증빙URL": img_url
        })
        
    s3_client.put_object(Bucket=s3_bucket, Key=f"data/{date_path}/{team_name}/{user_name}_{timestamp}.json", Body=json.dumps(summary_list, ensure_ascii=False).encode('utf-8'))
    return True

# ==========================================
# 2. 메인 UI 및 사이드바 로직
# ==========================================
st.title("📑 AI 경비 제출 시스템")

with st.sidebar:
    st.header("👤 제출자 정보")
    user_name = st.text_input("성함", placeholder="홍길동")
    team_name = st.selectbox("소속 팀", ["영업1팀", "영업2팀", "개발팀", "인사팀", "마케팅팀", "기타"])
    
    st.divider()
    project_type = st.radio("프로젝트 수행 여부", ["해당없음", "기간 선택"], horizontal=True)
    
    max_project_cost = 0
    day_status = "해당없음"

    if project_type == "기간 선택":
        today = datetime.today()
        first_day = today.replace(day=1)
        
        dates = st.date_input("프로젝트 기간", value=(first_day, today))
        
        if len(dates) == 2:
            start_date, end_date = dates
            working_days = (end_date - start_date).days + 1
            total_month_days = calendar.monthrange(start_date.year, start_date.month)[1]
            
            max_project_cost = int(200000 * (working_days / total_month_days))
            day_status = f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}"
            
            st.success(f"💡 이번 달 프로젝트 한도: **{max_project_cost:,}원**\n(근무 {working_days}일 / 해당월 {total_month_days}일)")
        else:
            st.warning("달력에서 종료일을 한 번 더 클릭해주세요.")

# 카테고리 버튼
categories = ["야근식대", "야근교통비", "외근교통비", "프로젝트비용", "기타"]
cols = st.columns(5)
for i, cat in enumerate(categories):
    cols[i].button(f"📁 {cat}", use_container_width=True, type="primary" if st.session_state.selected_cat == cat else "secondary", on_click=change_category, args=(cat,))

st.divider()

# 파일 업로더
uploaded_files = st.file_uploader("영수증 파일을 올려주세요.", accept_multiple_files=True, key=f"receipt_uploader_{st.session_state.uploader_key}")

if uploaded_files:
    current_files = [f.name for f in uploaded_files]
    keys_to_remove = [k for k in st.session_state.file_cat_map.keys() if k not in current_files]
    for k in keys_to_remove:
        st.session_state.file_cat_map.pop(k)
        
    for f in uploaded_files:
        if f.name not in st.session_state.file_cat_map:
            st.session_state.file_cat_map[f.name] = st.session_state.selected_cat
            
    st.caption("📍 분류 현황: " + " | ".join([f"📄 {f.name} → **[{st.session_state.file_cat_map[f.name]}]**" for f in uploaded_files]))
else:
    st.session_state.file_cat_map.clear()

# 분석 버튼
if uploaded_files and st.button(f"✨ {len(uploaded_files)}건 AI 분석 시작", type="primary", use_container_width=True):
    st.session_state.submitted = False 
    with st.spinner("AI 분석 중..."):
        for f in uploaded_files:
            assigned_cat = st.session_state.file_cat_map.get(f.name, st.session_state.selected_cat)
            res = analyze_receipt(f)
            img = Image.open(f)
            img.thumbnail((500, 500))
            st.session_state.expense_items.append({
                "종류": assigned_cat, "결제일자": str(res.get("결제 날짜")), 
                "사용처": str(res.get("사용처")), "인식금액": safe_int(res.get("합계 금액")), 
                "배달비": 0, "비고": "", "image_display": img, "is_uncertain": res.get("is_uncertain", False)
            })
    st.session_state.expense_items.sort(key=lambda x: (categories.index(x['종류']), x['결제일자']))
    st.session_state.file_cat_map = {} 
    st.session_state.uploader_key += 1 
    st.rerun()

# ==========================================
# 3. 리스트 표시, 자동 절사 및 제출 로직
# ==========================================
if st.session_state.expense_items:
    
    limit = max_project_cost if project_type == "기간 선택" else 0
    current_proj_total = 0
    
    # 1. 렌더링 전 한도 초과 여부 확인
    limit_exceeded = False
    for i in st.session_state.expense_items:
        if i['종류'] == "프로젝트비용":
            if current_proj_total + i['인식금액'] > limit:
                limit_exceeded = True
                break
            current_proj_total += i['인식금액']

    if limit_exceeded and not st.session_state.submitted:
        if limit == 0:
            st.error("🚨 **[오류]** 프로젝트 기간이 설정되지 않았습니다! (현재 모든 프로젝트 비용 영수증이 제외 처리됨)")
        else:
            st.warning(f"⚠️ 한도({limit:,}원) 초과! 넘치는 금액은 자동으로 절사 및 제출에서 제외됩니다.")

    st.subheader("📝 내역 확인")
    
    if st.session_state.submitted:
        st.success("✅ 위 내역이 서버로 성공적으로 제출되었습니다!")

    # 2. 리스트 렌더링 및 실시간 절사 계산
    current_proj_total = 0 # 렌더링용 합계 초기화
    
    for idx, item in enumerate(st.session_state.expense_items):
        # [비율 조정] 표시금액 공간을 넓게 확보
        r1 = st.columns([1.2, 1.3, 1.8, 1.2, 1.6, 0.4, 0.4])
        item['종류'] = r1[0].selectbox(f"cat_{idx}", categories, index=categories.index(item['종류']), label_visibility="collapsed", disabled=st.session_state.submitted)
        item['결제일자'] = r1[1].text_input(f"dt_{idx}", item['결제일자'], label_visibility="collapsed", disabled=st.session_state.submitted)
        item['사용처'] = r1[2].text_input(f"vn_{idx}", item['사용처'], label_visibility="collapsed", disabled=st.session_state.submitted)
        item['인식금액'] = r1[3].number_input(f"am_{idx}", value=safe_int(item['인식금액']), step=100, label_visibility="collapsed", disabled=st.session_state.submitted)
        
        # 기본 입력비용
        input_cost = item['인식금액']
        if item['종류'] == '야근식대':
            input_cost += item.get('배달비', 0)
            
        effective_cost = input_cost
        status_html = f"<div style='margin-top:8px;'>**{effective_cost:,}**</div>"
        
        # [핵심] 프로젝트 비용 실시간 절사 시각화
        if item['종류'] == "프로젝트비용":
            if limit == 0:
                effective_cost = 0
                status_html = f"<div style='margin-top:2px; line-height:1.2;'><del>{input_cost:,}</del><br/><span style='color:#ff4b4b; font-size:12px; font-weight:bold;'>기간미설정 (제외)</span></div>"
            elif current_proj_total >= limit:
                effective_cost = 0
                status_html = f"<div style='margin-top:2px; line-height:1.2;'><del>{input_cost:,}</del><br/><span style='color:#ff4b4b; font-size:12px; font-weight:bold;'>한도초과 (제외)</span></div>"
            elif current_proj_total + input_cost > limit:
                effective_cost = limit - current_proj_total
                current_proj_total = limit
                status_html = f"<div style='margin-top:2px; line-height:1.2;'>**{effective_cost:,}**<br/><span style='color:#ff9c2a; font-size:12px; font-weight:bold;'>절사됨 (원래 {input_cost:,})</span></div>"
            else:
                effective_cost = input_cost
                current_proj_total += effective_cost
                status_html = f"<div style='margin-top:8px;'>**{effective_cost:,}**</div>"
                
        item['_effective_cost'] = effective_cost # S3 저장을 위해 숨겨진 키에 실제 반영 금액 저장
        r1[4].markdown(status_html, unsafe_allow_html=True)
        
        with r1[5]:
            with st.popover("🖼️"): st.image(item['image_display'], use_container_width=True)
        if r1[6].button("🗑️", key=f"del_{idx}", disabled=st.session_state.submitted):
            st.session_state.expense_items.pop(idx)
            st.rerun()

        is_high_cost_meal = (item['종류'] == "야근식대" and input_cost >= 15000)
        if is_high_cost_meal:
            r2 = st.columns([1.2, 4.3, 1.5])
            item['비고'] = r2[1].text_input(f"note_{idx}", item['비고'], placeholder="함께 식사한 인원 정보를 입력하세요.", label_visibility="collapsed", disabled=st.session_state.submitted)
            item['배달비'] = r2[2].number_input(f"del_fee_{idx}", value=item['배달비'], step=500, label_visibility="collapsed", disabled=st.session_state.submitted)
            st.divider()
        else: st.divider()

    # 제출 버튼 로직
    if not st.session_state.submitted:
        if st.button("🚀 서버로 최종 제출", type="primary", use_container_width=True):
            if not user_name: st.error("제출자 성함을 입력해주세요.")
            elif project_type == "기간 선택" and max_project_cost == 0:
                st.error("달력에서 프로젝트 종료일을 마저 선택해주세요.")
            else:
                with st.spinner("서버로 전송 중... (제외된 항목은 전송되지 않습니다)"):
                    if save_to_s3(user_name, team_name, day_status, st.session_state.expense_items):
                        st.balloons()
                        st.session_state.submitted = True 
                        st.rerun()
    else:
        if st.button("🔄 새 영수증 작성하기 (목록 초기화)", use_container_width=True):
            st.session_state.expense_items = []
            st.session_state.submitted = False
            st.session_state.uploader_key += 1
            st.rerun()