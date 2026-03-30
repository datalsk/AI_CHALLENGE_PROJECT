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
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/variable/pretendardvariable.css");
    html, body, [class*="css"], .stMarkdown, .stText, button, input, select {
        font-family: 'Pretendard Variable', Pretendard, -apple-system, sans-serif !important;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display:none;}
    header {background-color: transparent !important;}

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        box-shadow: rgba(0, 0, 0, 0.04) 0px 4px 12px !important;
        border: 1px solid rgba(226, 232, 240, 0.1) !important;
        background-color: rgba(255, 255, 255, 0.03) !important;
        padding: 4px;
        transition: all 0.2s ease;
    }
    
    .stButton > button[kind="primary"] {
        background-color: #4f46e5;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        letter-spacing: -0.3px;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #3730a3;
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
    }
    
    .stButton > button[kind="secondary"] {
        border-radius: 8px;
        font-weight: 500;
        border: 1px solid rgba(148, 163, 184, 0.3);
        background-color: transparent;
    }
    .stButton > button[kind="secondary"]:hover {
        background-color: rgba(148, 163, 184, 0.1);
    }
    
    h1 { font-weight: 700 !important; letter-spacing: -1px; margin-bottom: 0px !important;}
    h3 { font-weight: 600 !important; letter-spacing: -0.5px; }
    
    div[data-baseweb="input"], div[data-baseweb="select"] {
        border-radius: 8px !important;
        border: none !important;
        background-color: rgba(148, 163, 184, 0.08) !important;
        box-shadow: inset 0 0 0 1px rgba(148, 163, 184, 0.2) !important;
        transition: all 0.2s ease;
    }
    div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within {
        box-shadow: inset 0 0 0 2px #4f46e5 !important;
        background-color: rgba(148, 163, 184, 0.12) !important;
    }
    div[data-baseweb="input"] > div > input, div[data-baseweb="select"] > div {
        background-color: transparent !important;
    }
    
    div[role="radiogroup"] { gap: 1rem; }
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

def analyze_receipt(uploaded_file, retries=1):
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except KeyError:
        st.error("Secrets에 OPENAI_API_KEY가 없습니다.")
        return {"결제 날짜": "에러", "사용처": "키 없음", "합계 금액": 0}

    base64_image = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    
    prompt = """
    영수증 이미지에서 다음 3가지 정보를 반드시 추출하여 JSON 형식으로만 응답해.
    1. "결제 날짜": YYYY-MM-DD 형식. 
    2. "사용처": 상호명 추출.
    3. "합계 금액": 최종 결제 금액 (숫자만).
    * 경고: 절대로 'None', 'null' 같은 문자열을 반환하지 마. 안 보이면 "미확인" 또는 0을 써.
    """
    
    payload = {
        "model": "gpt-4o-mini", 
        "temperature": 0.0, 
        "messages": [
            {"role": "system", "content": "너는 영수증 데이터를 기계처럼 정확하게 추출하는 시스템이야."},
            {"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{uploaded_file.type};base64,{base64_image}"}}]}
        ], 
        "response_format": { "type": "json_object" }
    }
    
    for attempt in range(retries + 1):
        try:
            response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=45)
            res_data = json.loads(response.json()['choices'][0]['message']['content'])
            
            date_str = str(res_data.get("결제 날짜", "")).strip().lower()
            shop_str = str(res_data.get("사용처", "")).strip().lower()
            
            if (date_str in ["none", "null", ""] or shop_str in ["none", "null", ""]) and attempt < retries:
                time.sleep(1)
                continue
            
            if date_str in ["none", "null", "", "미확인"]: res_data["결제 날짜"] = datetime.now().strftime("%Y-%m-%d")
            else: res_data["결제 날짜"] = str(res_data.get("결제 날짜"))
            
            if shop_str in ["none", "null", "", "미확인"]: res_data["사용처"] = "미확인"
            else: res_data["사용처"] = str(res_data.get("사용처"))
            
            res_data["is_uncertain"] = (res_data.get("사용처") == "미확인" or safe_int(res_data.get("합계 금액")) == 0)
            return res_data
            
        except:
            if attempt < retries:
                time.sleep(1)
                continue
            break
            
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
        del_img_url = "N/A"
        
        if item.get('image_display'):
            img_key = f"images/{date_path}/{team_name}/{user_name}_{timestamp}_{idx}.png"
            img_byte_arr = io.BytesIO()
            item['image_display'].save(img_byte_arr, format='PNG')
            s3_client.put_object(Bucket=s3_bucket, Key=img_key, Body=img_byte_arr.getvalue(), ContentType='image/png')
            img_url = f"https://{s3_bucket}.s3.{aws_region}.amazonaws.com/{img_key}"
            
        if item.get('배달비_이미지_display'):
            del_img_key = f"images/{date_path}/{team_name}/{user_name}_{timestamp}_{idx}_delivery.png"
            del_img_byte_arr = io.BytesIO()
            item['배달비_이미지_display'].save(del_img_byte_arr, format='PNG')
            s3_client.put_object(Bucket=s3_bucket, Key=del_img_key, Body=del_img_byte_arr.getvalue(), ContentType='image/png')
            del_img_url = f"https://{s3_bucket}.s3.{aws_region}.amazonaws.com/{del_img_key}"
            
        summary_list.append({
            "이름": user_name, "팀명": team_name, "항목": item['종류'], 
            "금액": final_amt, "결제일자": item['결제일자'], 
            "사용처": item['사용처'], "수행일자": day_status, 
            "비고": item.get('비고', ""), "증빙URL": img_url,
            "배달비_증빙URL": del_img_url if del_img_url != "N/A" else ""
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
    
    st.markdown("<hr style='margin: 1.5rem 0; border-top: 1px solid rgba(148, 163, 184, 0.2);'>", unsafe_allow_html=True)
    
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

categories = ["야근식대", "야근교통비", "외근교통비", "프로젝트비용", "기타"]
cols = st.columns(5)
for i, cat in enumerate(categories):
    cols[i].button(cat, use_container_width=True, type="primary" if st.session_state.selected_cat == cat else "secondary", on_click=change_category, args=(cat,))

st.write("") 

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

if uploaded_files and st.button(f"총 {len(uploaded_files)}건 영수증 자동 입력", type="primary", use_container_width=True):
    st.session_state.submitted = False 
    
    total_files = len(uploaded_files)
    progress_bar = st.progress(0, text="AI가 영수증 데이터를 추출하고 있습니다...")
    
    for i, f in enumerate(uploaded_files):
        assigned_cat = st.session_state.file_cat_map.get(f.name, st.session_state.selected_cat)
        res = analyze_receipt(f) 
        img = Image.open(f)
        img.thumbnail((500, 500))
        
        st.session_state.expense_items.append({
            "종류": assigned_cat, 
            "결제일자": res.get("결제 날짜"), 
            "사용처": res.get("사용처"), 
            "인식금액": safe_int(res.get("합계 금액")), 
            "배달비": 0, 
            "비고": "", 
            "image_display": img, 
            "배달비_이미지_display": None, 
            "is_uncertain": res.get("is_uncertain", False)
        })
        
        if i < total_files - 1:
            time.sleep(1) 
            
        progress_percentage = int(((i + 1) / total_files) * 100)
        progress_bar.progress(progress_percentage, text=f"총 {total_files}건 중 {i+1}건 완료...")
        
    st.session_state.expense_items.sort(key=lambda x: (categories.index(x['종류']), x['결제일자']))
    st.session_state.file_cat_map = {} 
    st.session_state.uploader_key += 1 
    time.sleep(0.5)
    progress_bar.empty()
    st.rerun()

# ==========================================
# 3. 리스트 표시, 자동 절사 및 제출 로직
# ==========================================
if st.session_state.expense_items:
    st.markdown("<hr style='margin: 2rem 0; border-top: 1px solid rgba(148, 163, 184, 0.2);'>", unsafe_allow_html=True)
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
            status_html = f"<div style='margin-top:8px; font-size: 16px; font-weight: 600;'>{effective_cost:,} 원</div>"
            
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
                    status_html = f"<div style='margin-top:2px; line-height:1.2;'><span style='font-size: 16px; font-weight: 600;'>{effective_cost:,} 원</span><br/><span style='color:#f59e0b; font-size:12px; font-weight:600;'>절사됨 (입력 {input_cost:,})</span></div>"
                else:
                    effective_cost = input_cost
                    current_proj_total += effective_cost
                    status_html = f"<div style='margin-top:8px; font-size: 16px; font-weight: 600;'>{effective_cost:,} 원</div>"
                    
            item['_effective_cost'] = effective_cost
            r1[4].markdown(status_html, unsafe_allow_html=True)
            
            with r1[5]:
                with st.popover("영수증"): st.image(item['image_display'], use_container_width=True)
            if r1[6].button("삭제", key=f"del_{idx}", disabled=st.session_state.submitted):
                st.session_state.expense_items.pop(idx)
                st.rerun()

            is_high_cost_meal = (item['종류'] == "야근식대" and input_cost >= 15000)
            if is_high_cost_meal:
                st.markdown("<hr style='margin: 0.5rem 0; border-top: 1px dashed rgba(148, 163, 184, 0.3);'>", unsafe_allow_html=True)
                
                reason_key = f"reason_{idx}"
                if reason_key not in st.session_state:
                    st.session_state[reason_key] = "동석자 입력"
                
                reason = st.radio("초과 사유 증빙 방식을 선택하세요", ["동석자 입력", "배달비 증빙"], horizontal=True, key=reason_key, disabled=st.session_state.submitted)
                
                if reason == "동석자 입력":
                    item['비고'] = st.text_input("동석자 정보", value=item.get('비고', ''), placeholder="함께 식사한 인원 정보를 입력하세요 (예: 홍길동, 김철수)", key=f"note_{idx}", disabled=st.session_state.submitted)
                    item['배달비'] = 0
                    item['배달비_이미지_display'] = None
                else:
                    # [수정] 배달비 입력, 파일 업로드, 그리고 팝업 미리보기 레이아웃 최적화
                    c2_1, c2_2, c2_3 = st.columns([1.5, 3, 1.2])
                    
                    item['배달비'] = c2_1.number_input("배달비 금액", value=item.get('배달비', 0), step=500, key=f"del_fee_{idx}", disabled=st.session_state.submitted)
                    
                    del_file = c2_2.file_uploader("배달비 영수증 첨부 (이미지 파일)", type=["png", "jpg", "jpeg"], key=f"del_file_{idx}", disabled=st.session_state.submitted)
                    
                    if del_file:
                        item['배달비_이미지_display'] = Image.open(del_file)
                        
                    with c2_3:
                        st.write("") # 버튼 위치를 업로더 박스 중앙과 맞추기 위한 여백
                        st.write("")
                        if item.get('배달비_이미지_display'):
                            with st.popover("미리보기"):
                                st.image(item['배달비_이미지_display'], use_container_width=True)
                                
                    item['비고'] = "배달비 증빙"

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