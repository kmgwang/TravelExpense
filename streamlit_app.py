import io
import os
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

# PDF 생성을 위한 ReportLab 모듈
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 페이지 기본 설정
st.set_page_config(
    page_title="화천기공 해외출장비 자동 산정 및 정산 프로그램",
    layout="wide"
)

# 숫자 입력 필드의 스핀박스 완전 제거 및 테이블/입력 필드 스타일 설정 CSS
st.markdown("""
<style>
/* 숫자 input 스핀박스 제거 (+, - 버튼 및 브라우저 기본 스핀박스 숨김) */
input[type=number]::-webkit-inner-spin-button, 
input[type=number]::-webkit-outer-spin-button { 
    -webkit-appearance: none; 
    margin: 0; 
}
input[type=number] {
    -moz-appearance: textfield;
}
/* 입력 필드 및 텍스트 가운데 정렬 */
.stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] {
    text-align: center !important;
}
</style>
""", unsafe_allow_html=True)

# 출장비 기준 규정 테이블
RULES_TABLE = {
    ("갑", "임원(부사장이상)"): {"daily": 135, "hotel": -1, "curr": "USD"},
    ("갑", "임원"): {"daily": 90, "hotel": 130, "curr": "USD"},
    ("갑", "1급"): {"daily": 70, "hotel": 100, "curr": "USD"},
    ("갑", "2급"): {"daily": 65, "hotel": 95, "curr": "USD"},
    ("갑", "3급이하"): {"daily": 60, "hotel": 90, "curr": "USD"},

    ("을", "임원(부사장이상)"): {"daily": 130, "hotel": -1, "curr": "USD"},
    ("을", "임원"): {"daily": 85, "hotel": 125, "curr": "USD"},
    ("을", "1급"): {"daily": 65, "hotel": 95, "curr": "USD"},
    ("을", "2급"): {"daily": 60, "hotel": 90, "curr": "USD"},
    ("을", "3급이하"): {"daily": 55, "hotel": 85, "curr": "USD"},

    ("병", "임원(부사장이상)"): {"daily": 130, "hotel": -1, "curr": "USD"},
    ("병", "임원"): {"daily": 80, "hotel": 110, "curr": "USD"},
    ("병", "1급"): {"daily": 60, "hotel": 90, "curr": "USD"},
    ("병", "2급"): {"daily": 55, "hotel": 85, "curr": "USD"},
    ("병", "3급이하"): {"daily": 55, "hotel": 80, "curr": "USD"},

    ("특", "임원(부사장이상)"): {"daily": 23000, "hotel": -1, "curr": "JPY"},
    ("특", "임원"): {"daily": 11000, "hotel": 17000, "curr": "JPY"},
    ("특", "1급"): {"daily": 8000, "hotel": 12000, "curr": "JPY"},
    ("특", "2급"): {"daily": 7000, "hotel": 11000, "curr": "JPY"},
    ("특", "3급이하"): {"daily": 7000, "hotel": 10000, "curr": "JPY"},
}

# 세션 상태 초기화
if "df_input" not in st.session_state:
    st.session_state.df_input = pd.DataFrame(columns=[
        "부서", "성명", "출장지", "지역구분", "직급", "직급구분", 
        "출발일", "도착일", "숙박일수", "출장일수", "적용환율", 
        "일당_KRW", "일당_지급처", "숙박비_KRW", "숙박비_지급처", 
        "교통비_KRW", "교통비_지급처", "기타경비_KRW", "기타_지급처"
    ])

if "custom_etc_items" not in st.session_state:
    st.session_state.custom_etc_items = [{"name": "여행자 보험", "amount": 0, "pay": "여행사"}]

DEPARTMENTS = [
    "임원", "경영지원본부", "경영지원실", "인사지원팀", "관리팀", 
    "재무전략실", "노동조합", "재무팀", "자금팀", "정보실", 
    "정보팀", "IBU", "성장전략실", "프로젝트팀", "구매전략본부", 
    "HTB 대만지사", "구매팀", "VI팀", "품질혁신본부", "QM팀", 
    "보전팀", "생산본부", "생산관리팀", "생산기술팀", "가공팀", 
    "F/S가공", "정밀가공", "가공지원", "UNIT팀", "UNIT준비", 
    "UNIT조립", "UNIT서비스", "생산1팀", "생산2팀", "서비스센터", 
    "서비스1팀", "서비스2팀", "서비스3팀", "서비스4팀", "기술개발연구소", 
    "MC개발팀", "TC개발팀", "5축개발팀", "UNIT개발팀", "제어개발팀", 
    "제어SW개발팀", "가공기술1팀", "가공기술2팀", "소재사업부문", "기타"
]

POSITION_MAPPING = {
    "회장": "임원(부사장이상)", "명예회장": "임원(부사장이상)", "사장": "임원(부사장이상)", "부사장": "임원(부사장이상)",
    "전무": "임원", "상무": "임원", "이사": "임원", "이사대우": "임원",
    "부장": "1급", "차장": "1급",
    "과장": "2급", "대리": "2급",
    "계장": "3급이하", "사원": "3급이하", "1급 기능장": "3급이하", "2급 기능장": "3급이하"
}

def get_region_by_location(loc):
    loc = loc.strip()
    if not loc:
        return ""
    group_gap = [
        "영국", "독일", "프랑스", "이탈리아", "스페인", "스위스", "네덜란드", "벨기에", "오스트리아", "포르투갈", "스웨덴", "노르웨이", "덴마크", "핀란드", "아일랜드", "그리스",
        "미국", "캐나다", "멕시코", "브라질", "아르헨티나", "칠레", "콜롬비아", "페루",
        "UAE", "아랍에미리트", "사우디", "카타르", "이스라엘", "쿠웨이트", "오만", "바레인", "요르단", "레바논",
        "이집트", "남아공", "남아프리카공화국", "나이지리아", "케냐", "모로코", "알제리", "튜니지아", "에티오피아", "가나",
        "호주", "뉴질랜드", "피지", "파푸아뉴기니",
        "폴란드", "체코", "루마니아", "우크라이나", "헝가리", "슬로바키아", "불가리아", "크로아티아", "세르비아", "리투아니아", "라트비아", "에스토니아",
        "러시아",
        "싱가포르", "홍콩", "대만"
    ]
    group_eul = ["중국"]
    group_byeong = [
        "베트남", "태국", "말레이시아", "인도네시아", "필리핀", "싱가포르", "미얀마", "캄보디아", "라오스", "브루나이",
        "인도", "파키스탄", "방글라데시", "스리랑카", "네パール", "부탄", "몰디브",
        "카자흐스탄", "우즈베키스탄", "투르크메니스탄", "키르기스스탄", "타지키스탄", "몽골"
    ]
    group_teuk = ["일본"]

    if any(k in loc for k in group_teuk):
        return "특"
    elif any(k in loc for k in group_eul):
        if not any(sub in loc for sub in ["홍콩", "대만"]):
            return "을"
        else:
            return "갑"
    elif any(k in loc for k in group_gap):
        return "갑"
    elif any(k in loc for k in group_byeong):
        return "병"
    return ""

def calculate_row_expenses(row, main_rate):
    region = str(row.get("지역구분", ""))
    rank = str(row.get("직급구분", ""))
    days = int(row.get("출장일수", 1))
    nights = int(row.get("숙박일수", 0))

    rule = RULES_TABLE.get((region, rank), {"daily": 60, "hotel": 90, "curr": "USD"})
    daily_std = rule["daily"]
    hotel_std = rule["hotel"]

    actual_rate = (main_rate / 100.0) if region == "특" else main_rate

    total_daily_foreign = daily_std * days
    daily_krw = int((total_daily_foreign * actual_rate) // 1000) * 1000

    if hotel_std == -1:
        hotel_krw = 0.0
    else:
        total_hotel_foreign = hotel_std * nights
        hotel_krw = int((total_hotel_foreign * actual_rate) // 1000) * 1000

    return daily_krw, hotel_krw

def calculate_expenses_df(df, main_rate):
    results = []
    for _, row in df.iterrows():
        days = int(row.get("출장일수", 1))
        nights = int(row.get("숙박일수", 0))
        
        d_calc, h_calc = calculate_row_expenses(row, main_rate)
        daily_krw = float(row.get("일당_KRW", d_calc)) if pd.notnull(row.get("일당_KRW")) else d_calc
        hotel_krw = float(row.get("숙박비_KRW", h_calc)) if pd.notnull(row.get("숙박비_KRW")) else h_calc
        
        trans_krw = float(row.get("교통비_KRW", 0))
        etc_krw = float(row.get("기타경비_KRW", 0))

        trans_pay = str(row.get("교통비_지급처", "여행사"))
        daily_pay = str(row.get("일당_지급처", "출장자"))
        hotel_pay = str(row.get("숙박비_지급처", "출장자"))
        etc_pay = str(row.get("기타_지급처", "여행사"))

        agency_pay = 0
        employee_pay = 0

        for amt, pay in [(daily_krw, daily_pay), (hotel_krw, hotel_pay), (trans_krw, trans_pay), (etc_krw, etc_pay)]:
            if "여행사" in pay:
                agency_pay += amt
            else:
                employee_pay += amt

        total_krw = agency_pay + employee_pay

        res_dict = row.to_dict()
        res_dict.update({
            "적용환율": main_rate,
            "일당_KRW": daily_krw,
            "숙박비_KRW": hotel_krw,
            "총출장비_KRW": total_krw,
            "여행사지급액_KRW": agency_pay,
            "출장자지급액_KRW": employee_pay,
            "출장기간": f"{nights}박 {days}일"
        })
        results.append(res_dict)
    return pd.DataFrame(results)

# 상단 공통 설정
st.title("✈️ 화천기공 해외출장비 자동 산정 및 정산 프로그램")

with st.container():
    col_r1, col_r2 = st.columns([2, 3])
    with col_r1:
        main_rate = st.number_input("💱 메인 적용 환율 (USD/JPY 기준)", value=1415.30, step=0.01, format="%.2f")
    with col_r2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"🔗 [서울외국환중개 환율조회 사이트 바로가기](http://www.smbs.biz/ExRate/TodayExRate.jsp)", unsafe_allow_html=True)

st.markdown("---")

tab1, tab2, tab3 = st.tabs([
    "1. 출장자 정보 직접 입력", 
    "2. 자금팀 연결용 이미지 파일 생성", 
    "3. 출장자 출장비 산정내역서 (PDF)"
])

# ----------------------------------------------------
# 탭 1: 출장자 정보 직접 입력
# ----------------------------------------------------
with tab1:
    st.subheader("📝 출장자 정보 등록 및 관리")
    
    # 엔터키 제출 방지용 on_submit 빈 함수 처리 적용
    with st.form("traveler_form", clear_on_submit=False, enter_to_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            dep = st.selectbox("부서", DEPARTMENTS)
            name = st.text_input("성명", placeholder="성명을 입력하세요")
        with c2:
            date_start = st.date_input("출발일", value=datetime.today())
            date_end = st.date_input("도착일", value=datetime.today() + timedelta(days=4))
            flight_night = st.checkbox("기내 박 적용 (숙박 1박 차감)")

        c3, c4 = st.columns(2)
        with c3:
            position = st.selectbox("직급", [""] + list(POSITION_MAPPING.keys()))
            auto_rank = POSITION_MAPPING.get(position, "")
            rank_options = ["", "3급이하", "2급", "1급", "임원", "임원(부사장이상)"]
            rank_idx = rank_options.index(auto_rank) if auto_rank in rank_options else 0
            rank = st.selectbox("직급구분", rank_options, index=rank_idx)
        with c4:
            location = st.text_input("출장지", placeholder="예: 미국 로스앤젤레스, 일본 도쿄")
            auto_reg = get_region_by_location(location)
            region_options = ["", "갑", "을", "병", "특"]
            region_idx = region_options.index(auto_reg) if auto_reg in region_options else 0
            region = st.selectbox("지역구분", region_options, index=region_idx)

        st.markdown("##### 💰 경비 항목 입력")
        
        temp_days = max(1, (date_end - date_start).days + 1)
        temp_nights = max(0, temp_days - 1)
        if flight_night:
            temp_nights = max(0, temp_nights - 1)
        
        dummy_row = {"지역구분": region if region else auto_reg, "직급구분": rank, "출장일수": temp_days, "숙박일수": temp_nights}
        calc_d_init, calc_h_init = calculate_row_expenses(dummy_row, main_rate)

        st.markdown("###### ⚙️ 기타 경비 항목 관리")
        etc_col_a, etc_col_b = st.columns([3, 1])
        with etc_col_a:
            new_etc_name = st.text_input("추가할 기타 항목명", placeholder="예: 비자발급비, 통신비 등", label_visibility="collapsed")
        with etc_col_b:
            if st.form_submit_button("기타 항목 추가"):
                if new_etc_name.strip():
                    st.session_state.custom_etc_items.append({"name": new_etc_name.strip(), "amount": 0, "pay": "여행사"})
                    st.rerun()

        # 하나의 통일된 표 구조 구현 (테두리, 폰트, 배경색 통일)
        table_html = """
        <style>
        .custom-expense-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            margin-bottom: 10px;
            font-size: 14px;
            text-align: center;
            font-family: inherit;
        }
        .custom-expense-table th, .custom-expense-table td {
            border: 1px solid #cbd5e1;
            padding: 8px;
            text-align: center;
            vertical-align: middle;
        }
        .custom-expense-table th {
            background-color: #f1f5f9;
            font-weight: bold;
            color: #1e293b;
        }
        .custom-expense-table tr:nth-child(even) {
            background-color: #f8fafc;
        }
        .custom-expense-table tr:nth-child(odd) {
            background-color: #ffffff;
        }
        .total-row {
            background-color: #fef08a !important;
            font-weight: bold;
        }
        </style>
        <table class="custom-expense-table">
            <thead>
                <tr>
                    <th style="width: 15%;">대분류</th>
                    <th style="width: 20%;">소분류</th>
                    <th style="width: 20%;">지급처</th>
                    <th style="width: 25%;">금액 (KRW)</th>
                    <th style="width: 20%;">비고</th>
                </tr>
            </thead>
            <tbody>
        """
        st.markdown(table_html, unsafe_allow_html=True)

        # 각 항목 입력 위젯 배치 (금액은 타이핑 입력 전용, step=None 처리)
        col_t1, col_t2, col_t3, col_t4, col_t5 = st.columns([1.5, 2.0, 2.0, 2.5, 2.0])
        with col_t1: st.markdown("교통비")
        with col_t2: st.markdown("항공권")
        with col_t3: air_pay = st.selectbox("항공권 지급처", ["여행사", "출장자"], index=0, key="air_pay_sel", label_visibility="collapsed")
        with col_t4: air_krw = st.number_input("항공권 금액", value=0, step=None, format="%d", key="air_val", label_visibility="collapsed")
        with col_t5: st.markdown("-")

        col_e1, col_e2, col_e3, col_e4, col_e5 = st.columns([1.5, 2.0, 2.0, 2.5, 2.0])
        with col_e1: st.markdown("")
        with col_e2: st.markdown("ESTA")
        with col_e3: esta_pay = st.selectbox("ESTA 지급처", ["여행사", "출장자"], index=0, key="esta_pay_sel", label_visibility="collapsed")
        with col_e4: esta_krw = st.number_input("ESTA 금액", value=0, step=None, format="%d", key="esta_val", label_visibility="collapsed")
        with col_e5: st.markdown("-")

        col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([1.5, 2.0, 2.0, 2.5, 2.0])
        with col_h1: st.markdown("출장비")
        with col_h2: st.markdown("숙박비")
        with col_h3: hotel_pay = st.selectbox("숙박비 지급처", ["여행사", "출장자"], index=1, key="hotel_pay_sel", label_visibility="collapsed")
        with col_h4: hotel_krw = st.number_input("숙박비 금액", value=int(calc_h_init), step=None, format="%d", key="hotel_val", label_visibility="collapsed")
        with col_h5: st.markdown("자동산정/수정가능")

        col_d1, col_d2, col_d3, col_d4, col_d5 = st.columns([1.5, 2.0, 2.0, 2.5, 2.0])
        with col_d1: st.markdown("")
        with col_d2: st.markdown("일당")
        with col_d3: daily_pay = st.selectbox("일당 지급처", ["여행사", "출장자"], index=1, key="daily_pay_sel", label_visibility="collapsed")
        with col_d4: daily_krw = st.number_input("일당 금액", value=int(calc_d_init), step=None, format="%d", key="daily_val", label_visibility="collapsed")
        with col_d5: st.markdown("자동산정/수정가능")

        total_etc_krw = 0
        etc_payments = []
        for idx, item in enumerate(st.session_state.custom_etc_items):
            gc1, gc2, gc3, gc4, gc5 = st.columns([1.5, 2.0, 2.0, 2.5, 2.0])
            with gc1: 
                if idx == 0: st.markdown("기타")
                else: st.markdown("")
            with gc2: st.markdown(item["name"])
            with gc3: 
                p_val = st.selectbox(f"지급처_{idx}", ["여행사", "출장자"], index=0 if item["pay"]=="여행사" else 1, key=f"etc_pay_{idx}", label_visibility="collapsed")
                item["pay"] = p_val
            with gc4: 
                a_val = st.number_input(f"금액_{idx}", value=int(item["amount"]), step=None, format="%d", key=f"etc_val_{idx}", label_visibility="collapsed")
                item["amount"] = a_val
            with gc5:
                if st.form_submit_button(f"삭제_{idx}", use_container_width=True):
                    st.session_state.custom_etc_items.pop(idx)
                    st.rerun()
            
            total_etc_krw += a_val
            etc_payments.append((a_val, p_val))

        # 총액 합계 행을 하나의 표 내부에 완벽히 통합 렌더링
        total_sum_preview = air_krw + esta_krw + hotel_krw + daily_krw + total_etc_krw
        
        closing_table_html = f"""
            <tr class="total-row">
                <td colspan="3" style="text-align: right; padding-right: 20px;">입력 항목 총액 합계</td>
                <td colspan="2" style="text-align: center;">₩ {int(total_sum_preview):,}</td>
            </tr>
            </tbody>
        </table>
        """
        st.markdown(closing_table_html, unsafe_allow_html=True)

        submitted = st.form_submit_button("➕ 출장자 정보 등록 / 추가", use_container_width=True)

        if submitted:
            final_region = region if region else auto_reg
            if not name.strip():
                st.error("성명을 입력해주세요.")
            elif not final_region:
                st.error("지역구분을 선택해주세요.")
            elif not rank:
                st.error("직급을 선택해주세요.")
            else:
                days = (date_end - date_start).days + 1
                nights = max(0, days - 1)
                if flight_night:
                    nights = max(0, nights - 1)

                if days <= 0:
                    st.error("도착일은 출발일 이후여야 합니다.")
                else:
                    new_data = {
                        "부서": dep,
                        "성명": name.strip(),
                        "출장지": location.strip(),
                        "지역구분": final_region,
                        "직급": position,
                        "직급구분": rank,
                        "출발일": date_start.strftime("%Y-%m-%d"),
                        "도착일": date_end.strftime("%Y-%m-%d"),
                        "숙박일수": nights,
                        "출장일수": days,
                        "적용환율": main_rate,
                        "일당_KRW": daily_krw,
                        "일당_지급처": daily_pay,
                        "숙박비_KRW": hotel_krw,
                        "숙박비_지급처": hotel_pay,
                        "교통비_KRW": air_krw + esta_krw,
                        "교통비_지급처": air_pay,
                        "기타경비_KRW": total_etc_krw,
                        "기타_지급처": etc_payments[0][1] if etc_payments else "여행사"
                    }
                    
                    new_df = pd.DataFrame([new_data])
                    st.session_state.df_input = pd.concat([st.session_state.df_input, new_df], ignore_index=True)
                    st.success(f"성공적으로 등록되었습니다: {name.strip()} ({location})")

    st.markdown("---")
    st.subheader(f"📋 등록된 출장자 목록 (총 {len(st.session_state.df_input)}건)")
    
    if not st.session_state.df_input.empty:
        display_df = calculate_expenses_df(st.session_state.df_input, main_rate)
        st.dataframe(display_df, use_container_width=True)
        
        if st.button("🗑️ 전체 목록 초기화"):
            st.session_state.df_input = pd.DataFrame(columns=st.session_state.df_input.columns)
            st.rerun()
    else:
        st.info("등록된 출장자 정보가 없습니다. 위 폼을 통해 입력해주세요.")

# ----------------------------------------------------
# 탭 2: 자금팀 연결용 이미지 파일 생성
# ----------------------------------------------------
with tab2:
    st.subheader("🖼️ 자금팀 제출용 정산 내역 이미지 생성")
    st.markdown("입력된 모든 출장자의 내역을 종합하여 자금팀 확인용 요약 이미지(PNG)를 생성합니다.")

    if st.session_state.df_input.empty:
        st.warning("먼저 '1. 출장자 정보 직접 입력' 탭에서 출장자 데이터를 입력해주세요.")
    else:
        if st.button("🚀 자금팀 제출용 이미지 생성하기", type="primary"):
            try:
                df_res = calculate_expenses_df(st.session_state.df_input, main_rate)
                
                try:
                    font_title = ImageFont.truetype("malgun.ttf", 26)
                    font_bold = ImageFont.truetype("malgunbd.ttf", 15)
                    font_regular = ImageFont.truetype("malgun.ttf", 14)
                except:
                    try:
                        font_title = ImageFont.truetype("DejaVuSans-Bold.ttf", 26)
                        font_bold = ImageFont.truetype("DejaVuSans-Bold.ttf", 15)
                        font_regular = ImageFont.truetype("DejaVuSans.ttf", 14)
                    except:
                        font_title = font_bold = font_regular = ImageFont.load_default()

                cols = ["순번", "출장지", "부서", "출장자", "출발", "도착", "출장기간", "교통비", "출장비", "기타금액", "집계", "출장자 지급", "여행사", "총계", "지급요청일"]
                col_widths = [50, 110, 110, 110, 110, 110, 80, 110, 110, 100, 120, 120, 120, 120, 130]
                row_height = 36
                header_h1 = 30
                
                img_width = sum(col_widths) + 40
                img_height = 140 + header_h1 + row_height + (len(df_res) * row_height)
                
                img = Image.new("RGB", (img_width, img_height), color=(255, 255, 255))
                draw = ImageDraw.Draw(img)

                current_year = datetime.now().year
                draw.text((20, 20), f"{current_year}년 해외출장비 지급 내역", fill=(0, 0, 0), font=font_title)

                x_start = 20
                y_start = 80

                x_hr_start = x_start + sum(col_widths[:7])
                x_hr_end = x_start + sum(col_widths[:11])
                draw.rectangle([x_hr_start, y_start, x_hr_end, y_start + header_h1], fill=(30, 78, 161), outline=(0, 0, 0))
                draw.text((x_hr_start + (x_hr_end - x_hr_start)/2 - 60, y_start + 5), "인사지원팀 확인", fill=(255, 255, 255), font=font_bold)

                x_fin_start = x_hr_end
                x_fin_end = x_start + sum(col_widths[:15])
                draw.rectangle([x_fin_start, y_start, x_fin_end, y_start + header_h1], fill=(245, 205, 170), outline=(0, 0, 0))
                draw.text((x_fin_start + (x_fin_end - x_fin_start)/2 - 40, y_start + 5), "자금팀 확인", fill=(0, 0, 0), font=font_bold)

                y_head = y_start + header_h1
                x_curr = x_start
                for idx, c_name in enumerate(cols):
                    w = col_widths[idx]
                    box = [x_curr, y_head, x_curr + w, y_head + row_height]
                    
                    if 7 <= idx <= 10:
                        bg_color = (30, 78, 161)
                        txt_color = (255, 255, 255)
                    elif 11 <= idx <= 14:
                        bg_color = (245, 205, 170)
                        txt_color = (0, 0, 0)
                    else:
                        bg_color = (230, 230, 230)
                        txt_color = (0, 0, 0)

                    draw.rectangle(box, fill=bg_color, outline=(0, 0, 0))
                    
                    bbox = draw.textbbox((0, 0), c_name, font=font_bold)
                    tw = bbox[2] - bbox[0]
                    th = bbox[3] - bbox[1]
                    tx = x_curr + (w - tw) / 2
                    ty = y_head + (row_height - th) / 2 - 2
                    draw.text((tx, ty), c_name, fill=txt_color, font=font_bold)
                    x_curr += w

                weekdays = ("월", "화", "수", "목", "금", "토", "일")
                y_data = y_head + row_height

                for row_i, r in df_res.iterrows():
                    try:
                        d_start_dt = datetime.strptime(str(r['출발일']), "%Y-%m-%d")
                        start_str = d_start_dt.strftime("%m월 %d일")
                        end_str = datetime.strptime(str(r['도착일']), "%Y-%m-%d").strftime("%m월 %d일")
                        pay_req_dt = d_start_dt - timedelta(days=1)
                        pay_req_str = f"{pay_req_dt.strftime('%m월 %d일')}({weekdays[pay_req_dt.weekday() ]})"
                    except:
                        start_str = str(r['출발일'])
                        end_str = str(r['도착일'])
                        pay_req_str = str(r['출발일'])

                    days_num = str(r['출장기간']).split('일')[0].split('박')[-1].strip() + "일" if '일' in str(r['출장기간']) else str(r['출장기간'])
                    sub_total = r["교통비_KRW"] + r["숙박비_KRW"] + r["기타경비_KRW"] + r["일당_KRW"]

                    row_vals = [
                        str(row_i + 1),
                        r["출장지"],
                        r["부서"],
                        r["성명"],
                        start_str,
                        end_str,
                        days_num,
                        f"₩{int(r['교통비_KRW']):,}",
                        f"₩{int(r['숙박비_KRW'] + r['일당_KRW']):,}",
                        f"₩{int(r['기타경비_KRW']):,}",
                        f"₩{int(sub_total):,}",
                        f"₩{int(r['출장자지급액_KRW']):,}",
                        f"₩{int(r['여행사지급액_KRW']):,}",
                        f"₩{int(r['총출장비_KRW']):,}",
                        pay_req_str
                    ]

                    x_curr = x_start
                    for idx, val in enumerate(row_vals):
                        w = col_widths[idx]
                        box = [x_curr, y_data, x_curr + w, y_data + row_height]
                        
                        f_to_use = font_bold if idx in [14, 10, 13] else font_regular
                        
                        if idx == 14:
                            draw.rectangle(box, fill=(220, 38, 38), outline=(0, 0, 0))
                            txt_color = (255, 255, 255)
                        else:
                            draw.rectangle(box, fill=(255, 255, 255), outline=(0, 0, 0))
                            txt_color = (0, 0, 0)

                        bbox = draw.textbbox((0, 0), str(val), font=f_to_use)
                        tw = bbox[2] - bbox[0]
                        th = bbox[3] - bbox[1]
                        tx = x_curr + (w - tw) / 2
                        ty = y_data + (row_height - th) / 2 - 2
                        
                        draw.text((tx, ty), str(val), fill=txt_color, font=f_to_use)
                        x_curr += w
                    y_data += row_height

                buf = io.BytesIO()
                img.save(buf, format="PNG")
                byte_im = buf.getvalue()

                st.success("이미지가 성공적으로 생성되었습니다!")
                st.image(byte_im, caption="생성된 자금팀 확인용 내역서", use_container_width=True)

                file_name = f"자금팀연결용_해외출장비지급내역_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                st.download_button(
                    label="💾 이미지 파일 다운로드 (.png)",
                    data=byte_im,
                    file_name=file_name,
                    mime="image/png"
                )

            except Exception as e:
                st.error(f"이미지 생성 중 오류 발생: {str(e)}")

# ----------------------------------------------------
# 탭 3: 출장자 출장비 산정내역서 (PDF)
# ----------------------------------------------------
with tab3:
    st.subheader("📄 출장자별 출장비 산정내역서 PDF 생성")
    st.markdown("등록된 출장자별 공식 산정내역서 PDF를 개별 또는 일괄 생성합니다.")

    if st.session_state.df_input.empty:
        st.warning("먼저 '1. 출장자 정보 직접 입력' 탭에서 출장자 데이터를 입력해주세요.")
    else:
        df_res = calculate_expenses_df(st.session_state.df_input, main_rate)
        
        selected_index = st.selectbox(
            "PDF를 생성할 출장자를 선택하세요:",
            options=df_res.index,
            format_func=lambda i: f"[{df_res.loc[i, '부서']}] {df_res.loc[i, '성명']} (출장지: {df_res.loc[i, '출장지']})"
        )

        if st.button("📥 선택한 출장자 산정내역서 PDF 생성", type="primary"):
            r = df_res.loc[selected_index]
            
            try:
                try:
                    pdfmetrics.registerFont(TTFont("MalgunGothic", "malgun.ttf"))
                    pdfmetrics.registerFont(TTFont("MalgunGothicBold", "malgunbd.ttf"))
                    font_name = "MalgunGothic"
                    font_bold = "MalgunGothicBold"
                except:
                    font_name = "Helvetica"
                    font_bold = "Helvetica-Bold"

                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(
                    pdf_buffer,
                    pagesize=A4,
                    rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
                )
                elements = []

                style_title = ParagraphStyle('TitleStyle', fontName=font_bold, fontSize=18, leading=22, alignment=1)
                style_h2 = ParagraphStyle('H2Style', fontName=font_bold, fontSize=11, leading=14, textColor=colors.HexColor("#000000"))
                style_center = ParagraphStyle('CenterStyle', fontName=font_name, fontSize=9, leading=12, alignment=1)
                style_center_bold = ParagraphStyle('CenterBold', fontName=font_bold, fontSize=9, leading=12, alignment=1)
                style_header_cell = ParagraphStyle('HeaderCell', fontName=font_bold, fontSize=9, leading=11, alignment=1, textColor=colors.black)

                elements.append(Paragraph("해외출장비 산정 내역서", style_title))
                elements.append(Spacer(1, 10))

                rate_val = float(r['적용환율'])
                if str(r['지역구분']) == "특":
                    rate_val_str = f"1 ¥ = {rate_val:,.2f} 원 (적용환율: {rate_val / 100.0:,.4f})"
                else:
                    rate_val_str = f"1 USD = {rate_val:,.2f} 원"

                info_data = [
                    [
                        Paragraph("소속", style_center_bold), Paragraph(str(r["부서"]), style_center),
                        Paragraph("성명", style_center_bold), Paragraph(str(r["성명"]), style_center),
                        Paragraph("직급", style_center_bold), Paragraph(str(r["직급"]), style_center)
                    ],
                    [
                        Paragraph("출장지", style_center_bold), Paragraph(str(r["출장지"]), style_center),
                        Paragraph("지역구분", style_center_bold), Paragraph(str(r["지역구분"]), style_center),
                        Paragraph("직급구분", style_center_bold), Paragraph(str(r["직급구분"]), style_center)
                    ],
                    [
                        Paragraph("출발일", style_center_bold), Paragraph(str(r["출발일"]), style_center),
                        Paragraph("도착일", style_center_bold), Paragraph(str(r["도착일"]), style_center),
                        Paragraph("출장기간", style_center_bold), Paragraph(str(r["출장기간"]), style_center)
                    ],
                    [
                        Paragraph("적용환율", style_center_bold), Paragraph("", style_center), Paragraph("", style_center),
                        Paragraph(rate_val_str, style_center), Paragraph("", style_center), Paragraph("", style_center)
                    ]
                ]

                info_table = Table(info_data, colWidths=[65, 110, 65, 110, 65, 125])
                info_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F1F5F9")),
                    ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F1F5F9")),
                    ('BACKGROUND', (4,0), (4,-1), colors.HexColor("#F1F5F9")),
                    ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor("#FEF08A")),
                    ('SPAN', (0, 3), (2, 3)),
                    ('SPAN', (3, 3), (5, 3)),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94A3B8")),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                elements.append(info_table)
                elements.append(Spacer(1, 12))

                elements.append(Paragraph("■ 적용 출장비 산정 기준", style_h2))
                elements.append(Spacer(1, 4))

                region = str(r["지역구분"])
                rank = str(r["직급구분"])
                rule = RULES_TABLE.get((region, rank), {"daily": 60, "hotel": 90, "curr": "USD"})
                
                daily_std = rule["daily"]
                hotel_std = rule["hotel"]
                curr_type = rule["curr"]

                daily_std_str = f"{daily_std:,} {curr_type} / 일" if curr_type == "JPY" else f"{daily_std} USD / 일"
                calc_rate_display = (rate_val / 100.0) if region == "특" else rate_val
                rate_str_val = f"{calc_rate_display:,.4f}"

                daily_formula = f"{daily_std} {curr_type} * {r['출장일수']}일 * {rate_str_val} (천원단위절사)"
                if hotel_std == -1:
                    hotel_std_str = "실비"
                    hotel_formula = "실비 정산"
                    hotel_amt_str = "실비"
                else:
                    hotel_std_str = f"{hotel_std:,} {curr_type} / 박" if curr_type == "JPY" else f"{hotel_std} USD / 박"
                    hotel_formula = f"{hotel_std} {curr_type} * {r['숙박일수']}박 * {rate_str_val} (천원단위절사)"
                    hotel_amt_str = f"₩ {int(r['숙박비_KRW']):,}"

                total_sum_val = int(r['숙박비_KRW'] + r['일당_KRW']) if hotel_std != -1 else int(r['일당_KRW'])

                calc_data = [
                    [Paragraph("구분", style_header_cell), Paragraph("산정 기준", style_header_cell), Paragraph("기간 적용", style_header_cell), Paragraph("금액", style_header_cell), Paragraph("원화 환산 산식", style_header_cell)],
                    [Paragraph("숙박비", style_center), Paragraph(hotel_std_str, style_center), Paragraph(f"{r['숙박일수']} 박", style_center), Paragraph(hotel_amt_str, style_center), Paragraph(hotel_formula, style_center)],
                    [Paragraph("일당", style_center), Paragraph(daily_std_str, style_center), Paragraph(f"{r['출장일수']} 일", style_center), Paragraph(f"₩ {int(r['일당_KRW']):,}", style_center), Paragraph(daily_formula, style_center)],
                    [Paragraph("지급 총액", style_center_bold), Paragraph("", style_center), Paragraph("", style_center), Paragraph(f"₩ {total_sum_val:,}", style_center_bold), Paragraph("숙박비 + 일당", style_center)]
                ]

                calc_table = Table(calc_data, colWidths=[70, 110, 70, 110, 190])
                calc_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2E8F0")),
                    ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#FEF08A")),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94A3B8")),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))
                elements.append(calc_table)
                elements.append(Spacer(1, 12))

                elements.append(Paragraph("■ 해외출장 지급규정", style_h2))
                elements.append(Spacer(1, 4))

                rule_header = [
                    Paragraph("지역", style_header_cell),
                    Paragraph("직급", style_header_cell),
                    Paragraph("일당", style_header_cell),
                    Paragraph("숙박", style_header_cell),
                    Paragraph("비고", style_header_cell)
                ]
                rule_rows = [rule_header]

                for (reg, rk), info in RULES_TABLE.items():
                    d_val = f"${info['daily']}" if info['curr'] == "USD" else f"¥{info['daily']:,}"
                    h_val = "실비" if info['hotel'] == -1 else (f"${info['hotel']}" if info['curr'] == "USD" else f"¥{info['hotel']:,}")
                    note = "엔화" if info['curr'] == "JPY" else ""
                    
                    rule_rows.append([
                        Paragraph(reg, style_center),
                        Paragraph(rk, style_center),
                        Paragraph(d_val, style_center),
                        Paragraph(h_val, style_center),
                        Paragraph(note, style_center)
                    ])

                rule_table = Table(rule_rows, colWidths=[70, 130, 110, 110, 130])
                rule_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#CBD5E1")),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94A3B8")),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 2.5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
                ]))
                elements.append(rule_table)
                elements.append(Spacer(1, 12))

                today_str = datetime.now().strftime('%Y년 %m월 %d일')
                footer_data = [
                    [Paragraph("위와 같이 해외출장비를 정산 및 지급합니다.", style_center)],
                    [Spacer(1, 4)],
                    [Paragraph(f"신청일 : {today_str}", style_center)]
                ]
                footer_table = Table(footer_data, colWidths=[550])
                footer_table.setStyle(TableStyle([
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ]))
                elements.append(footer_table)

                doc.build(elements)
                pdf_data = pdf_buffer.getvalue()

                st.success(f"{r['성명']} 님의 출장비 산정내역서 PDF가 생성되었습니다.")
                st.download_button(
                    label=f"💾 {r['성명']}_산정내역서.pdf 다운로드",
                    data=pdf_data,
                    file_name=f"해외출장비_산정내역서_{r['성명']}_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf"
                )

            except Exception as e:
                st.error(f"PDF 생성 중 오류 발생: {str(e)}")
