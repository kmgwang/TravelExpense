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

st.title("✈️ 화천기공 해외출장비 정산 및 내역서 자동 생성 시스템")
st.markdown(
    "인사지원팀 해외출장 경비 산정, 자금팀 제출용 정산표 분리, 출장자용 산정 내역서 자동 생성 프로그램입니다."
)
st.markdown("---")

# 세션 스테이트 초기화 (기본 데이터 구성)
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


# 3. 규정 기준 단가 반환 함수
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

        raw_rate = d["환율"]
        if region == "특":
            applied_rate = raw_rate / 100.0
        else:
            applied_rate = raw_rate

        calc_daily = std_daily * applied_rate
        d["산정일당_원화"] = int(calc_daily // 100 * 100)

        if std_hotel == "실비":
            calc_hotel = d.get("실비숙박_입력", 0)
        else:
            calc_hotel = std_hotel * applied_rate
        d["산정숙박_원화"] = int(calc_hotel // 100 * 100)

        d["직원_계좌입금액"] = d["산정일당_원화"] + d["산정숙박_원화"]
        d["여행사_지급액"] = (
            d.get("항공료", 0)
            + d.get("ESTA기타", 0)
            + d.get("여행자보험", 0)
            + d.get("수수료", 0)
        )
        d["총출장비"] = d["직원_계좌입금액"] + d["여행사_지급액"]
        processed.append(d)

    return pd.DataFrame(processed)


# 3가지 탭 구성
tab1, tab2, tab3 = st.tabs(
    [
        "📋 1. 출장 정보 입력",
        "💰 2. 자금팀 연결 자료 생성",
        "📄 3. 출장비 산정 내역서 생성",
    ]
)

# ----------------------------------------------------
# [Tab 1] 출장 정보 입력
# ----------------------------------------------------
with tab1:
    st.header("📋 해외출장 신청 및 경비 정보 입력")
    st.markdown(
        "출장자 정보와 출장지, 직급을 입력하면 규정에 따른 일당과 숙박비가 자동 산정됩니다."
    )
    st.markdown(
        "🔗 [서울외국환중개 환율 조회 사이트 바로가기](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
    )

    with st.form("travel_input_form"):
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("👤 출장자 정보")
            name = st.text_input("출장자 성명")
            department = st.text_input("부서", value="인사지원팀")
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
            # 직급 구분 자동 표시 기능 (참고용)
            preview_pos_group = get_position_group(position)
            st.text_input(
                "직급 구분",
                value=preview_pos_group,
                disabled=True,
                help="선택한 직급에 따라 자동 판정됩니다.",
            )

        with col_b:
            st.subheader("🌍 출장지 및 일정")
            country = st.text_input("출장지")
            start_date = st.date_input("출장 시작일")
            end_date = st.date_input("출장 종료일")
            exchange_rate = st.number_input(
                "적용 환율 (서울외국환중개 고시 환율)",
                min_value=0.0,
                value=1350.0,
                step=1.0,
            )

        st.markdown("---")
        st.subheader("💵 실비 및 여행사 대행 경비 입력 (원화)")
        col_c, col_d = st.columns(2)
        with col_c:
            st.markdown("##### [ 여행사 송금 항목 ]")
            flight = st.number_input(
                "항공료", min_value=0, value=0, step=10000
            )
            esta = st.number_input(
                "ESTA / 비자 비용", min_value=0, value=0, step=1000
            )
            insurance = st.number_input(
                "여행자 보험", min_value=0, value=0, step=1000
            )
            fee = st.number_input(
                "변경/취소 수수료", min_value=0, value=0, step=1000
            )
        with col_d:
            st.markdown("##### [ 숙박비 실비 적용 대상자용 ]")
            actual_hotel = st.number_input(
                "숙박비 실비 (임원 부사장 이상인 경우 입력)",
                min_value=0,
                value=0,
                step=10000,
            )
            st.info(
                "※ 임원(부사장 이상)의 숙박비는 '실비'로 적용되며, 그 외 직급은 규정 정액이 자동 적용됩니다."
            )

        submitted = st.form_submit_button(
            "➕ 입력한 출장 내역 규정 적용 및 추가"
        )

        if submitted:
            if not name or not country:
                st.error("⚠️ 출장자 성명과 출장지는 필수 입력 항목입니다.")
            else:
                new_data = {
                    "출장자성명": name,
                    "부서": department,
                    "직급": position,
                    "출장지": country,
                    "출장시작일": str(start_date),
                    "출장종료일": str(end_date),
                    "환율": exchange_rate,
                    "항공료": flight,
                    "ESTA기타": esta,
                    "여행자보험": insurance,
                    "수수료": fee,
                    "실비숙박_입력": actual_hotel,
                }
                st.session_state.travel_list.append(new_data)
                st.success(f"✅ {name} 님의 출장 경비가 규정에 맞춰 산정되었습니다!")

    st.markdown("---")
    st.subheader("📊 현재 등록된 전체 출장 내역 목록")
    if len(st.session_state.travel_list) > 0:
        current_df = process_travel_data(st.session_state.travel_list)
        st.dataframe(current_df, use_container_width=True)
        if st.button("🗑️ 전체 데이터 초기화"):
            st.session_state.travel_list = []
            st.rerun()
    else:
        st.info("등록된 출장 내역이 없습니다.")

# ----------------------------------------------------
# [Tab 2] 자금팀 연결 자료 생성
# ----------------------------------------------------
with tab2:
    st.header("💰 자금팀 제출용 정산 집계표 생성")
    st.markdown(
        "출장자에게 송금할 금액(일비+숙박비)과 여행사에 송금할 금액(항공권+ESTA+보험 등)이 완벽히 분리된 자금팀 제출용 표입니다."
    )

    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric(
                label="총 출장 건수", value=f"{len(processed_df)} 건"
            )
        with m2:
            total_agency = processed_df["여행사_지급액"].sum()
            st.metric(
                label="총 여행사 송금 총액",
                value=f"{total_agency:,.0f} 원",
            )
        with m3:
            total_employee = processed_df["직원_계좌입금액"].sum()
            st.metric(
                label="총 직원 계좌 입금 총액",
                value=f"{total_employee:,.0f} 원",
            )

        st.markdown("---")
        st.subheader("📑 자금팀 송금 요청 분리 집계표")

        fund_view_df = processed_df[
            [
                "출장자성명",
                "부서",
                "직급",
                "직급구분",
                "출장지",
                "직원_계좌입금액",
                "여행사_지급액",
                "총출장비",
            ]
        ].copy()
        fund_view_df.columns = [
            "성명",
            "부서",
            "직급",
            "직급구분",
            "출장지",
            "직원지급액(일비/숙박)",
            "여행사지급액(항공/보험등)",
            "총합계",
        ]
        st.dataframe(fund_view_df, use_container_width=True)

        output_agency = io.BytesIO()
        with pd.ExcelWriter(output_agency, engine="openpyxl") as writer:
            processed_df.to_excel(
                writer, index=False, sheet_name="자금팀_정산집계표"
            )
        output_agency.seek(0)

        st.download_button(
            label="📥 자금팀 제출용 정산 집계표 엑셀 다운로드",
            data=output_agency,
            file_name="화천기공_자금팀_출장정산집계표.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning(
            "⚠️ 입력된 출장 정보가 없습니다. [1. 출장 정보 입력] 탭에서 데이터를 먼저 입력해 주세요."
        )

# ----------------------------------------------------
# [Tab 3] 출장비 산정 내역서 생성
# ----------------------------------------------------
with tab3:
    st.header("📄 출장자용 해외출장비 산정 내역서 생성")
    st.markdown(
        "규정 기준, 환율 적용 방식, 백원 단위 절사 내역이 상세히 포함된 개인별 산정 내역서입니다."
    )

    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)

        selected_person = st.selectbox(
            "내역서를 생성할 출장자를 선택하세요",
            processed_df["출장자성명"].unique(),
        )

        person_data = processed_df[
            processed_df["출장자성명"] == selected_person
        ].iloc[0]

        st.markdown("---")
        st.markdown(f"### 👤 [ {selected_person} ] 님 해외출장비 산정 내역서")

        det_c1, det_c2 = st.columns(2)
        with det_c1:
            st.info(
                f"""
            - **소속 부서**: {person_data.get('부서', '-')}
            - **직급 / 직급 구분**: {person_data.get('직급', '-')} ({person_data.get('직급구분', '-')})
            - **출장지**: {person_data.get('출장지', '-')} (지역: {person_data.get('지역구분', '-')}급)
            """
            )
        with det_c2:
            st.success(
                f"""
            - **출장 기간**: {person_data.get('출장시작일', '')} ~ {person_data.get('출장종료일', '')}
            - **적용 환율**: {person_data.get('환율', 0):,.2f} 원 {'(엔화 100 환산 적용)' if person_data.get('지역구분')=='특' else ''}
            - **직원 계좌 입금 총액**: {person_data.get('직원_계좌입금액', 0):,.0f} 원
            """
            )

        hotel_display = (
            f"{person_data.get('산정숙박_원화', 0):,.0f} 원"
            if person_data.get("기준숙박_외화") != "실비"
            else f"실비 정산 ({person_data.get('실비숙박_입력', 0):,.0f} 원)"
        )
        currency_unit = "엔(¥)" if person_data.get("지역구분") == "특" else "달러($)"

        detail_table = pd.DataFrame(
            {
                "경비 항목": ["일비", "숙박비", "항공료 및 대행비"],
                "규정 기준단가": [
                    f"{person_data.get('기준일당_외화', 0):,d} {currency_unit}",
                    (
                        f"{person_data.get('기준숙박_외화', 0):,d} {currency_unit}"
                        if person_data.get("기준숙박_외화") != "실비"
                        else "실비"
                    ),
                    "실비 청구",
                ],
                "산정 금액 (원화)": [
                    f"{person_data.get('산정일당_원화', 0):,.0f} 원",
                    hotel_display,
                    f"{person_data.get('여행사_지급액', 0):,.0f} 원",
                ],
                "지급처 및 비고": [
                    "출장자 개인 계좌 입금 (백원 이하 절사)",
                    "출장자 개인 계좌 입금 (백원 이하 절사)",
                    "여행사 송금 지급",
                ],
            }
        )
        st.table(detail_table)

        output_person = io.BytesIO()
        with pd.ExcelWriter(output_person, engine="openpyxl") as writer:
            pd.DataFrame([person_data]).to_excel(
                writer, index=False, sheet_name="산정내역서"
            )
        output_person.seek(0)

        st.download_button(
            label=f"📥 [{selected_person}] 출장자용 산정 내역서 엑셀 다운로드",
            data=output_person,
            file_name=f"화천기공_해외출장산정내역서_{selected_person}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning(
            "⚠️ 입력된 출장 정보가 없습니다. [1. 출장 정보 입력] 탭에서 데이터를 먼저 입력해 주세요."
        )

