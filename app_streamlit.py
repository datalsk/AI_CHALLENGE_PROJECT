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
# 0. UI 설정 및 모던 SaaS 디자인 CSS 적용
# ==========================================
st.set_page_config(page_title="경비 정산 시스템", layout="wide")

st.markdown("""
    <style>
    /* 1. 폰트: 현대적인 Pretendard 적용 */
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/variable/pretendardvariable.css");
    html, body, [class*="css"], .stMarkdown, .stText, button, input, select {
        font-family: 'Pretendard Variable', Pretendard, -apple-system, sans-serif !important;
    }
    
    /* 2. 불필요한 기본 UI 숨기기 (사이드바 버튼 복구) */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display:none;}
    /* header 전체를 숨기면 사이드바 열기 버튼이 날아가므로 배경만 투명하게 처리 */
    header {background-color: transparent !important;}

    /* 3. 모던 카드 UI (은은한 그림자와 둥근 모서리) */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        box-shadow: rgba(0, 0, 0, 0.04) 0px 4px 12px !important;
        border: 1px solid rgba(226, 232, 240, 0.8) !important;
        background-color: #ffffff !important;
        padding: 4px;
        transition: all 0.2s ease;
    }
    
    /* 4. 주요 버튼 (Primary) 세련된 딥블루/블랙 톤으로 변경 */
    .stButton > button[kind="primary"] {
        background-color: #1e293b;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        letter-spacing: -0.3px;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #0f172a;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* 5. 일반 카테고리 버튼 깔끔하게 */
    .stButton > button[kind="secondary"] {
        border-radius: 8px;
        font-weight: 500;
        border: 1px solid #cbd5e1;
        background-color: #f8fafc;
        color: #334155;
    }
    .stButton > button[kind="secondary"]:hover {
        border-color: #94a3b8;
        background-color: #f1f5f9;
    }
    
    /* 6. 타이틀 및 헤더 정렬 */
    h1 { font-weight: 700 !important; color: #0f172a; letter-spacing: -1px; margin-bottom: 0px !important;}
    h3 { font-weight: 600 !important; color: #1e293b; letter-spacing: -0.5px; }
    
    /* 7. 라디오, 셀렉트 박스, 텍스트 인풋 등 폼 요소 둥글게 */
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>div {
        border-radius: 6px;
        border: 1px solid #cbd5e1;
    }
    </style>
    """, unsafe_allow_html=True)

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
        final_amt = item.get('_effective_cost', 0)
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
st.title("경비 정산 시스템")
st.markdown("<p style='color: #64748b; font-size: 15px; margin-bottom: 2rem;'>영수증을 업로드하면 시스템이 자동으로 정보를 분류하고 입력합니다.</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h3 style='margin-bottom: 0.5rem;'>제출자 정보</h3>", unsafe_allow_html=True)
    user_name = st.text_input("이름", placeholder="이름을 입력하세요")
    team_name = st.selectbox("소속 팀", ["영업1팀", "영업2팀", "개발팀", "인사팀", "마케팅팀", "기타"])
    
    st.markdown("<hr style='margin: 1.5rem 0; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)
    
    st.markdown("<h3 style='margin-bottom: 0.5rem;'>프로젝트 설정</h3>", unsafe_allow_html=True)
    project_type = st.radio("프로젝트 수행 여부", ["해당없음", "기간 선택"], horizontal=True, label_visibility="collapsed")
    
    max_project_cost = 0
    day_status = "해당없음"

    if project_type == "기간 선택":
        today = datetime.today()
        first_day = today.replace(day=1)
        dates = st.date_input("프로젝트 수행 기간", value=(first_day, today))
        
        if len(dates) == 2:
            start_date, end_date = dates
            working_days = (end_date - start_date).days + 1
            total_month_days = calendar.monthrange(start_date.year, start_date.month)[1]
            max_project_cost = int(200000 * (working_days / total_month_days))
            day_status = f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}"
            
            st.info(f"이번 달 배정 한도: {max_project_cost:,}원\n(근무 {working_days}일 / 해당월 {total_month_days}일)", icon="ℹ️")
        else:
            st.warning("달력에서 종료일을 선택해주세요.", icon="⚠️")

st.write("") 

# 카테고리 버튼
categories = ["야근식대", "야근교통비", "외근교통비", "프로젝트비용", "기타"]
cols = st.columns(5)
for i, cat in enumerate(categories):
    cols[i].button(cat, use_container_width=True, type="primary" if st.session_state.selected_cat == cat else "secondary", on_click=change_category, args=(cat,))

st.write("") 

# 파일 업로더
uploaded_files = st.file_uploader("증빙 자료(영수증) 업로드", accept_multiple_files=True, key=f"receipt_uploader_{st.session_state.uploader_key}")

if uploaded_files:
    current_files = [f.name for f in uploaded_files]
    keys_to_remove = [k for k in st.session_state.file_cat_map.keys() if k not in current_files]
    for k in keys_to_remove:
        st.session_state.file_cat_map.pop(k)
        
    for f in uploaded_files:
        if f.name not in st.session_state.file_cat_map:
            st.session_state.file_cat_map[f.name] = st.session_state.selected_cat
            
    st.caption("분류 대기 항목: " + " | ".join([f"{f.name} ({st.session_state.file_cat_map[f.name]})" for f in uploaded_files]))
else:
    st.session_state.file_cat_map.clear()

# 분석 버튼
if uploaded_files and st.button(f"총 {len(uploaded_files)}건 영수증 자동 입력", type="primary", use_container_width=True):
    st.session_state.submitted = False 
    with st.spinner("데이터를 추출하고 있습니다..."):
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
    st.markdown("<hr style='margin: 2rem 0; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)
    limit = max_project_cost if project_type == "기간 선택" else 0
    current_proj_total = 0
    
    limit_exceeded = False
    for i in st.session_state.expense_items:
        if i['종류'] == "프로젝트비용":
            if current_proj_total + i['인식금액'] > limit:
                limit_exceeded = True
                break
            current_proj_total += i['인식금액']

    if limit_exceeded and not st.session_state.submitted:
        if limit == 0:
            st.error("프로젝트 기간이 설정되지 않아 관련 비용이 제외 처리되었습니다.", icon="🚨")
        else:
            st.warning(f"프로젝트 한도({limit:,}원) 초과분이 자동으로 절사되었습니다.", icon="⚠️")

    st.markdown("<h3 style='margin-bottom: 1rem;'>정산 내역 확인</h3>", unsafe_allow_html=True)
    
    if st.session_state.submitted:
        st.success("해당 내역이 시스템에 성공적으로 등록되었습니다.", icon="✅")

    current_proj_total = 0 
    
    for idx, item in enumerate(st.session_state.expense_items):
        with st.container(border=True):
            r1 = st.columns([1.2, 1.3, 1.8, 1.2, 1.6, 0.6, 0.5])
            item['종류'] = r1[0].selectbox(f"cat_{idx}", categories, index=categories.index(item['종류']), label_visibility="collapsed", disabled=st.session_state.submitted)
            item['결제일자'] = r1[1].text_input(f"dt_{idx}", item['결제일자'], label_visibility="collapsed", disabled=st.session_state.submitted)
            item['사용처'] = r1[2].text_input(f"vn_{idx}", item['사용처'], label_visibility="collapsed", disabled=st.session_state.submitted)
            item['인식금액'] = r1[3].number_input(f"am_{idx}", value=safe_int(item['인식금액']), step=100, label_visibility="collapsed", disabled=st.session_state.submitted)
            
            input_cost = item['인식금액']
            if item['종류'] == '야근식대':
                input_cost += item.get('배달비', 0)
                
            effective_cost = input_cost
            status_html = f"<div style='margin-top:8px; font-size: 16px; font-weight: 600; color: #0f172a;'>{effective_cost:,} 원</div>"
            
            if item['종류'] == "프로젝트비용":
                if limit == 0:
                    effective_cost = 0
                    status_html = f"<div style='margin-top:2px; line-height:1.2;'><del style='color: #94a3b8;'>{input_cost:,}</del><br/><span style='color:#ef4444; font-size:12px; font-weight:600;'>기간 미설정 (제외)</span></div>"
                elif current_proj_total >= limit:
                    effective_cost = 0
                    status_html = f"<div style='margin-top:2px; line-height:1.2;'><del style='color: #94a3b8;'>{input_cost:,}</del><br/><span style='color:#ef4444; font-size:12px; font-weight:600;'>한도 초과 (제외)</span></div>"
                elif current_proj_total + input_cost > limit:
                    effective_cost = limit - current_proj_total
                    current_proj_total = limit
                    status_html = f"<div style='margin-top:2px; line-height:1.2;'><span style='font-size: 16px; font-weight: 600; color: #0f172a;'>{effective_cost:,} 원</span><br/><span style='color:#f59e0b; font-size:12px; font-weight:600;'>절사됨 (입력 {input_cost:,})</span></div>"
                else:
                    effective_cost = input_cost
                    current_proj_total += effective_cost
                    status_html = f"<div style='margin-top:8px; font-size: 16px; font-weight: 600; color: #0f172a;'>{effective_cost:,} 원</div>"
                    
            item['_effective_cost'] = effective_cost
            r1[4].markdown(status_html, unsafe_allow_html=True)
            
            with r1[5]:
                with st.popover("영수증 보기"): st.image(item['image_display'], use_container_width=True)
            if r1[6].button("삭제", key=f"del_{idx}", disabled=st.session_state.submitted):
                st.session_state.expense_items.pop(idx)
                st.rerun()

            is_high_cost_meal = (item['종류'] == "야근식대" and input_cost >= 15000)
            if is_high_cost_meal:
                st.markdown("<hr style='margin: 0.5rem 0; border-top: 1px dashed #cbd5e1;'>", unsafe_allow_html=True)
                r2 = st.columns([1.2, 4.3, 1.5])
                item['비고'] = r2[1].text_input(f"note_{idx}", item['비고'], placeholder="함께 식사한 인원 등 비고 사항을 입력하세요", label_visibility="collapsed", disabled=st.session_state.submitted)
                item['배달비'] = r2[2].number_input(f"del_fee_{idx}", value=item['배달비'], step=500, label_visibility="collapsed", disabled=st.session_state.submitted)

    st.write("")
    if not st.session_state.submitted:
        if st.button("최종 제출하기", type="primary", use_container_width=True):
            if not user_name: st.error("제출자 이름을 확인해주세요.", icon="🚨")
            elif project_type == "기간 선택" and max_project_cost == 0:
                st.error("달력에서 프로젝트 종료일을 확인해주세요.", icon="🚨")
            else:
                with st.spinner("서버에 데이터를 등록하고 있습니다..."):
                    if save_to_s3(user_name, team_name, day_status, st.session_state.expense_items):
                        st.toast('정산 내역이 성공적으로 등록되었습니다.', icon='✔️')
                        st.session_state.submitted = True 
                        time.sleep(1) 
                        st.rerun()
    else:
        if st.button("새 정산 작성하기", use_container_width=True):
            st.session_state.expense_items = []
            st.session_state.submitted = False
            st.session_state.uploader_key += 1
            st.rerun()