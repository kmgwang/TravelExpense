import io
import pandas as pd
import streamlit as st

# 페이지 설정 (반드시 최상단 위치)
try:
    st.set_page_config(
        page_title="화천기공 해외출장비 정산 자동화 프로그램",
        page_layout="wide",
    )
except Exception:
    pass

st.title("✈️ 화천기공 해외출장비 정산 및 자금팀 연결 자동화 시스템")
st.markdown(
    "엑셀 '1. 계산 시트' 프로세스(기본데이터 수기 입력 및 자동 산정)와 '2. 자금팀연결용 시트' 송금 요청 데이터 생성 프로그램입니다."
)
st.markdown("---")

# 세션 스테이트 초기화
if "travel_list" not in st.session_state:
    st.session_state.travel_list = []


# 1. 직급에 따른 직급 구분 함수
def get_position_group(position):
    executive_high = ["명예회장", "회장", "사장", "부사장"]
    executive = ["전무", "상무", "이사"]
    grade_1 = ["부장", "차장"]
    grade_2 = ["과장", "대리"]

    if position in executive_high:
        return "임원(부사장이상)"
    elif position in executive:
        return "임원"
    elif position in grade_1:
        return "1급"
    elif position in grade_2:
        return "2급"
    else:
        return "3급이하"


# 2. 국가별 지역 구분 매핑 함수
def get_region_group(country):
    country = country.strip()
    if "일본" in country:
        return "특"
    elif "중국" in country:
        return "을"
    elif any(
        x in country
        for x in [
            "베트남",
            "태국",
            "말레이시아",
            "인도네시아",
            "필리핀",
            "싱가포르",
            "인도",
            "파키스탄",
            "방글라데시",
            "카자흐스탄",
            "우즈베키스탄",
        ]
    ):
        if "싱가포르" in country:
            return "갑"
        return "병"
    else:
        return "갑"


# 3. 규정 기준 단가 반환 함수 (C4:G24 시뮬레이션)
def get_standard_rates(region, pos_group):
    rates = {
        "갑": {
            "임원(부사장이상)": (135, "실비"),
            "임원": (90, 130),
            "1급": (70, 100),
            "2급": (65, 95),
            "3급이하": (60, 90),
        },
        "을": {
            "임원(부사장이상)": (130, "실비"),
            "임원": (85, 125),
            "1급": (65, 95),
            "2급": (60, 90),
            "3급이하": (55, 85),
        },
        "병": {
            "임원(부사장이상)": (130, "실비"),
            "임원": (80, 110),
            "1급": (60, 90),
            "2급": (55, 85),
            "3급이하": (55, 80),
        },
        "특": {
            "임원(부사장이상)": (23000, "실비"),
            "임원": (11000, 17000),
            "1급": (8000, 12000),
            "2급": (7000, 11000),
            "3급이하": (7000, 10000),
        },
    }
    return rates.get(region, rates["갑"]).get(pos_group, (60, 90))


def process_travel_data(data_list):
    processed = []
    for item in data_list:
        d = item.copy()
        pos_group = get_position_group(d["직급"])
        region = get_region_group(d["출장지"])
        std_daily, std_hotel = get_standard_rates(region, pos_group)

        d["직급구분"] = pos_group
        d["지역구분"] = region
        d["기준일당_외화"] = std_daily
        d["기준숙박_외화"] = std_hotel

        # 숙박일 및 출장기간 자동 계산 (출발일~도착일 기준)
        start = pd.to_datetime(d["출발일"])
        end = pd.to_datetime(d["도착일"])
        trip_days = (end - start).days + 1
        if trip_days < 1:
            trip_days = 1
        hotel_nights = trip_days - 1 if trip_days > 1 else 0

        d["출장기간_일수"] = trip_days
        d["숙박일_일수"] = hotel_nights

        # 환율 적용 (특 지역은 100 나눔)
        raw_rate = d["환율"]
        applied_rate = raw_rate / 100.0 if region == "특" else raw_rate

        # 일당 및 숙박비 원화 환산 (백원단위 이하 절사: // 100 * 100)
        calc_daily = std_daily * applied_rate * trip_days
        d["산정일당_원화"] = int(calc_daily // 100 * 100)

        if std_hotel == "실비":
            calc_hotel = d.get("실비숙박_입력", 0)
        else:
            calc_hotel = std_hotel * applied_rate * hotel_nights
        d["산정숙박_원화"] = int(calc_hotel // 100 * 100)

        # Q3:V17 및 Y3:AK5 매핑 항목 계산
        d["출장자지급액_소계"] = d["산정일당_원화"] + d["산정숙박_원화"]
        d["여행사지급액_항공료"] = d.get("항공료", 0)
        d["여행사지급액_기타"] = (
            d.get("ESTA기타", 0) + d.get("여행자보험", 0) + d.get("수수료", 0)
        )
        d["여행사지급액_소계"] = (
            d["여행사지급액_항공료"] + d["여행사지급액_기타"]
        )
        d["총합계"] = d["출장자지급액_소계"] + d["여행사지급액_소계"]

        processed.append(d)

    return pd.DataFrame(processed)


# 2가지 핵심 시트(탭) 구성
tab1, tab2 = st.tabs(
    [
        "📊 1. 계산 시트 (기본데이터 입력 및 산정)",
        "💰 2. 자금팀연결용 시트 (송금 요청)",
    ]
)

# ----------------------------------------------------
# [Tab 1] 1. 계산 시트
# ----------------------------------------------------
with tab1:
    st.header("📊 [1. 계산 시트] 해외출장비 기본데이터 입력 및 자동 산정")
    st.markdown(
        "출장자 정보, 환율, 출장지, 기간을 입력하면 규정(C4:G24)을 자동 조회하여 Q3:V17 / Y3:AK5 영역의 산정 결과가 도출됩니다."
    )
    st.markdown(
        "🔗 [서울외국환중개 환율 조회 사이트 바로가기](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
    )

    with st.form("calc_sheet_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("👤 출장자 기본 정보 입력")
            name = st.text_input("출장자")
            exchange_date = st.date_input("환율적용일자")
            position = st.selectbox(
                "직급",
                [
                    "사장",
                    "부사장",
                    "전무",
                    "상무",
                    "이사",
                    "부장",
                    "차장",
                    "과장",
                    "대리",
                    "계장",
                    "사원",
                    "1급기능장",
                    "2급기능장",
                ],
            )
            preview_pos_group = get_position_group(position)
            st.text_input(
                "직급구분 (자동 판정)",
                value=preview_pos_group,
                disabled=True,
            )

        with col2:
            st.subheader("🌍 출장지 및 일정 입력")
            country = st.text_input(
                "출장지 (예: 미국, 일본, 베트남, 독일 등)"
            )
            preview_region = get_region_group(country if country else "미국")
            st.text_input(
                "지역구분 (자동 판정)", value=preview_region, disabled=True
            )

            start_date = st.date_input("출발일")
            end_date = st.date_input("도착일")
            exchange_rate = st.number_input(
                "적용환율 (서울외국환중개 고시 환율 수기 입력)",
                min_value=0.0,
                value=1350.0,
                step=1.0,
            )

        st.markdown("---")
        st.subheader("💳 실비 및 여행사 송금 경비 입력 (원화)")
        col3, col4 = st.columns(2)
        with col3:
            flight = st.number_input(
                "항공료 (여행사 지급)", min_value=0, value=0, step=10000
            )
            esta_ins = st.number_input(
                "ESTA / 여행자보험 / 수수료 등 (여행사 지급)",
                min_value=0,
                value=0,
                step=1000,
            )
        with col4:
            actual_hotel = st.number_input(
                "숙박비 실비 (임원 부사장 이상인 경우 입력)",
                min_value=0,
                value=0,
                step=10000,
            )
            st.info(
                "※ 출발일과 도착일에 따라 숙박일 및 출장기간이 자동 계산되며, C4:G24 규정 단가표에 연동됩니다."
            )

        submitted = st.form_submit_button(
            "➕ 계산 시트에 데이터 반영 및 산정 실행"
        )

        if submitted:
            if not name or not country:
                st.error("⚠️ 출장자와 출장지는 필수 입력 항목입니다.")
            else:
                new_data = {
                    "출장자": name,
                    "환율적용일자": str(exchange_date),
                    "출장지": country,
                    "직급": position,
                    "출발일": str(start_date),
                    "도착일": str(end_date),
                    "환율": exchange_rate,
                    "항공료": flight,
                    "ESTA기타": esta_ins,
                    "여행자보험": 0,
                    "수수료": 0,
                    "실비숙박_입력": actual_hotel,
                }
                st.session_state.travel_list.append(new_data)
                st.success(
                    f"✅ [{name}] 님의 해외출장비 계산이 완료되었습니다!"
                )

    st.markdown("---")
    st.subheader("📋 [Q3:V17 영역] 해외출장비 자동 산정 결과 미리보기")
    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)
        st.dataframe(processed_df, use_container_width=True)

        if st.button("🗑️ 계산 시트 데이터 초기화"):
            st.session_state.travel_list = []
            st.rerun()
    else:
        st.info("입력된 출장 기본데이터가 없습니다.")

# ----------------------------------------------------
# [Tab 2] 2. 자금팀연결용 시트
# ----------------------------------------------------
with tab2:
    st.header("💰 [2. 자금팀연결용 시트] 송금 요청 데이터 생성")
    st.markdown(
        "1. 계산 시트의 산정 결과(Y3:AK5 영역 값)를 연동하여, 교통비 및 기타 항목은 여행사로, 숙박비 및 일당은 출장자에게 송금하도록 정리한 자금팀 제출용 시트입니다."
    )

    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)

        # 자금팀 연결용 전용 뷰 생성 (Y3:AK5 핵심 매핑 구조 반영)
        fund_sheet_df = processed_df[
            [
                "출장자",
                "출장지",
                "직급구분",
                "출장기간_일수",
                "숙박일_일수",
                "환율",
                "산정일당_원화",
                "산정숙박_원화",
                "출장자지급액_소계",
                "여행사지급액_항공료",
                "여행사지급액_기타",
                "여행사지급액_소계",
                "총합계",
            ]
        ].copy()

        fund_sheet_df.columns = [
            "출장자",
            "출장지",
            "직급",
            "출장기간",
            "숙박일",
            "적용환율",
            "일당(출장자)",
            "숙박비(출장자)",
            "출장자 지급 소계",
            "항공료(여행사)",
            "기타비용(여행사)",
            "여행사 지급 소계",
            "총합계",
        ]

        st.subheader("📑 자금팀 송금 집계 표 (Y3:AK5 연동 데이터)")
        st.dataframe(fund_sheet_df, use_container_width=True)

        # 엑셀 다운로드 (계산 시트 및 자금팀연결용 시트 분리 포함)
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            processed_df.to_excel(
                writer, index=False, sheet_name="1. 계산 시트"
            )
            fund_sheet_df.to_excel(
                writer, index=False, sheet_name="2. 자금팀연결용 시트"
            )
        output_excel.seek(0)

        st.download_button(
            label="📥 엑셀 파일 다운로드 (계산 시트 + 자금팀연결용 시트)",
            data=output_excel,
            file_name="화천기공_해외출장비_정산_및_자금팀연결용.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning(
            "⚠️ 입력된 출장 정보가 없습니다. [1. 계산 시트] 탭에서 데이터를 먼저 입력해 주세요."
        )

1. 보완점(Risk): 출발일과 도착일 입력 시 날짜 계산 로직이 정확히 작동하도록 날짜 형식을 엄격히 준수해야 합니다.
2. 수정사항(Revision): 사용자가 설명한 '1. 계산 시트'의 수기 입력 프로세스와 '2. 자금팀연결용 시트'의 Y3:AK5 데이터 복사·붙여넣기 업무 흐름에 맞춰 웹 화면을 2개 탭으로 구조화했습니다.
3. 진행여부(Next Step): 위 프로세스 반영 코드로 테스트해 보신 뒤, 엑셀 서식이나 추가 연동 항목에 대해 보완할 점이 있으신가요?
