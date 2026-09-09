import datetime
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

# 경비 입력 폼의 동적 행 관리를 위한 세션 상태 초기화 (교통비, 출장비, 기타별 분리)
if "transport_rows" not in st.session_state:
    st.session_state.transport_rows = [
        {"item": "항공권", "amount": 0, "payer": "여행사"},
        {"item": "ESTA", "amount": 0, "payer": "여행사"}
    ]

if "travel_exp_rows" not in st.session_state:
    st.session_state.travel_exp_rows = [
        {"item": "숙박비", "amount": 0, "payer": "출장자"},
        {"item": "일당", "amount": 0, "payer": "출장자"}
    ]

if "other_rows" not in st.session_state:
    st.session_state.other_rows = [
        {"item": "", "amount": 0, "payer": "여행사"}
    ]


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
        pos_group = d["직급구분"]
        region = d["지역구분"]
        std_daily, std_hotel = get_standard_rates(region, pos_group)

        d["기준일당_외화"] = std_daily
        d["기준숙박_외화"] = std_hotel

        raw_rate = d["환율"]
        if region == "특":
            applied_rate = raw_rate / 100.0
        else:
            applied_rate = raw_rate

        user_daily = d.get("일당_입력값", 0)
        if user_daily > 0:
            calc_daily = user_daily
        else:
            calc_daily = std_daily * applied_rate * d.get("출장일수", 1)
            calc_daily = int(calc_daily // 100 * 100)
        d["산정일당_원화"] = calc_daily

        user_hotel = d.get("실비숙박_입력", 0)
        if user_hotel > 0:
            calc_hotel = user_hotel
        elif std_hotel == "실비":
            calc_hotel = 0
        else:
            calc_hotel = std_hotel * applied_rate * d.get("출장박수", 0)
            calc_hotel = int(calc_hotel // 100 * 100)
        d["산정숙박_원화"] = calc_hotel

        agency_total = 0
        employee_total = 0

        for t in d.get("교통비항목리스트", []):
            amt = t.get("amount", 0)
            payer = t.get("payer", "여행사")
            if payer == "여행사":
                agency_total += amt
            else:
                employee_total += amt

        for te in d.get("출장비항목리스트", []):
            amt = te.get("amount", 0)
            payer = te.get("payer", "출장자")
            if payer == "여행사":
                agency_total += amt
            else:
                employee_total += amt

        for other in d.get("기타항목리스트", []):
            o_amt = other.get("amount", 0)
            o_payer = other.get("payer", "여행사")
            if o_payer == "여행사":
                agency_total += o_amt
            else:
                employee_total += o_amt

        d["직원_계좌입금액"] = employee_total
        d["여행사_지급액"] = agency_total
        d["총출장비"] = employee_total + agency_total
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

    col_rate_info, col_rate_input = st.columns([2, 1])
    with col_rate_info:
        st.markdown(
            "🔗 [서울외국환중개 환율 조회 사이트 바로가기](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
        )
    with col_rate_input:
        exchange_rate = st.number_input(
            "적용 환율 입력", min_value=0.0, value=1350.0, step=1.0, format="%.2f"
        )

    department_list = [
        "임원",
        "경영지원본부",
        "경영지원실",
        "인사지원팀",
        "관리팀",
        "재무전략실",
        "노동조합",
        "재무팀",
        "자금팀",
        "정보실",
        "정보팀",
        "IBU",
        "성장전략실",
        "프로젝트팀",
        "구매전략본부",
        "HTB 대만지사",
        "구매팀",
        "VI팀",
        "품질혁신본부",
        "QM팀",
        "보전팀",
        "생산본부",
        "생산관리팀",
        "생산기술팀",
        "가공팀",
        "F/S가공",
        "정밀가공",
        "가공지원",
        "UNIT팀",
        "UNIT준비",
        "UNIT조립",
        "UNIT서비스",
        "생산1팀",
        "생산2팀",
        "서비스센터",
        "서비스1팀",
        "서비스2팀",
        "서비스3팀",
        "서비스4팀",
        "기술개발연구소",
        "MC개발팀",
        "TC개발팀",
        "5축개발팀",
        "UNIT개발팀",
        "제어개발팀",
        "제어SW개발팀",
        "가공기술1팀",
        "가공기술2팀",
        "소재사업부문",
        "기타",
    ]

    st.subheader("👤 출장자 정보 및 🌍 출장지 일정 설정")
    col_a, col_b = st.columns(2)

    with col_a:
        name = st.text_input("출장자 성명")
        department = st.selectbox(
            "부서",
            department_list,
            index=(
                department_list.index("인사지원팀")
                if "인사지원팀" in department_list
                else 0
            ),
        )
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

        auto_pos_group = get_position_group(position)
        pos_group_options = [
            "임원(부사장이상)",
            "임원",
            "1급",
            "2급",
            "3급이하",
        ]
        default_pos_idx = (
            pos_group_options.index(auto_pos_group)
            if auto_pos_group in pos_group_options
            else 4
        )
        position_group = st.selectbox(
            "직급 구분", pos_group_options, index=default_pos_idx
        )

    with col_b:
        default_start = datetime.date.today() + datetime.timedelta(days=1)
        default_end = default_start + datetime.timedelta(days=7)

        start_date = st.date_input("출장 시작일", value=default_start)
        end_date = st.date_input("출장 종료일", value=default_end)
        country = st.text_input("출장지")

        auto_region = get_region_group(country) if country else "갑"
        region_options = ["갑", "을", "병", "특"]
        default_reg_idx = (
            region_options.index(auto_region)
            if auto_region in region_options
            else 0
        )
        region_group = st.selectbox(
            "지역 구분", region_options, index=default_reg_idx
        )

    raw_days = (end_date - start_date).days + 1
    if raw_days < 1:
        raw_days = 1

    st.markdown("---")
    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        is_flight_minus = st.checkbox(
            "기내 박 적용(숙박 1박 차감)", value=False
        )
    with col_opt2:
        calculated_days = raw_days
        calculated_nights = (
            calculated_days - 1 if calculated_days > 1 else 0
        )
        if is_flight_minus:
            calculated_nights = max(0, calculated_nights - 1)

        st.info(
            f"📅 최종 산정된 출장 기간: **{calculated_nights}박 {calculated_days}일**"
        )

    st.markdown("---")
    with st.form("travel_input_sub_form"):
        st.subheader("💵 실비 및 여행사 대행 경비 입력 (원화)")
        st.markdown(
            "요청하신 표 형식의 입력 구조에 맞춰 아래 항목별 금액을 직접 입력해주세요."
        )

        col_ratios = [1.2, 1.5, 2, 1.5]

        # 헤더
        th1, th2, th3, th4 = st.columns(col_ratios)
        with th1:
            st.markdown("<div style='text-align: center;'><b>구분</b></div>", unsafe_allow_html=True)
        with th2:
            st.markdown("<div style='text-align: center;'><b>항목</b></div>", unsafe_allow_html=True)
        with th3:
            st.markdown("<div style='text-align: center;'><b>금액</b></div>", unsafe_allow_html=True)
        with th4:
            st.markdown("<div style='text-align: center;'><b>지급처</b></div>", unsafe_allow_html=True)

        st.markdown("---")
        payer_options = ["여행사", "출장자", "직접입력"]

        # --- 1. 교통비 섹션 ---
        updated_transport_rows = []
        transport_sum = 0
        t_len = len(st.session_state.transport_rows)
        for idx, row_data in enumerate(st.session_state.transport_rows):
            tc1, tc2, tc3, tc4 = st.columns(col_ratios)
            with tc1:
                if idx == 0:
                    st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>교통비</b></div>", unsafe_allow_html=True)
                else:
                    st.markdown("")
            with tc2:
                t_name = st.text_input(
                    f"교통비 항목 {idx}", value=row_data["item"], key=f"t_item_{idx}", label_visibility="collapsed"
                )
            with tc3:
                t_amt_str = st.text_input(
                    f"교통비 금액 {idx}", value=str(row_data["amount"]), key=f"t_amt_{idx}", label_visibility="collapsed"
                )
            with tc4:
                p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 0
                t_payer = st.selectbox(
                    f"교통비 지급처 {idx}", payer_options, index=p_idx, key=f"t_payer_{idx}", label_visibility="collapsed"
                )

            try:
                parsed_amt = int(t_amt_str.replace(",", ""))
            except ValueError:
                parsed_amt = 0

            transport_sum += parsed_amt
            updated_transport_rows.append({"item": t_name, "amount": parsed_amt, "payer": t_payer})

        st.session_state.transport_rows = updated_transport_rows

        # 교통비 총계 행
        sc1, sc2, sc3, sc4 = st.columns(col_ratios)
        with sc1:
            st.markdown("")
        with sc2:
            st.markdown("<div style='text-align: right;'><b>교통비 총계</b></div>", unsafe_allow_html=True)
        with sc3:
            st.markdown(f"<div style='text-align: right;'><b>{transport_sum:,.0f} 원</b></div>", unsafe_allow_html=True)
        with sc4:
            st.markdown("")

        st.markdown("---")

        # --- 2. 출장비 섹션 ---
        updated_travel_exp_rows = []
        travel_exp_sum = 0
        te_len = len(st.session_state.travel_exp_rows)
        for idx, row_data in enumerate(st.session_state.travel_exp_rows):
            tec1, tec2, tec3, tec4 = st.columns(col_ratios)
            with tec1:
                if idx == 0:
                    st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>출장비</b></div>", unsafe_allow_html=True)
                else:
                    st.markdown("")
            with tec2:
                te_name = st.text_input(
                    f"출장비 항목 {idx}", value=row_data["item"], key=f"te_item_{idx}", label_visibility="collapsed"
                )
            with tec3:
                te_amt_str = st.text_input(
                    f"출장비 금액 {idx}", value=str(row_data["amount"]), key=f"te_amt_{idx}", label_visibility="collapsed"
                )
            with tec4:
                p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 1
                te_payer = st.selectbox(
                    f"출장비 지급처 {idx}", payer_options, index=p_idx, key=f"te_payer_{idx}", label_visibility="collapsed"
                )

            try:
                parsed_amt = int(te_amt_str.replace(",", ""))
            except ValueError:
                parsed_amt = 0

            travel_exp_sum += parsed_amt
            updated_travel_exp_rows.append({"item": te_name, "amount": parsed_amt, "payer": te_payer})

        st.session_state.travel_exp_rows = updated_travel_exp_rows

        # 출장비 총계 행
        tc1, tc2, tc3, tc4 = st.columns(col_ratios)
        with tc1:
            st.markdown("")
        with tc2:
            st.markdown("<div style='text-align: right;'><b>출장비 총계</b></div>", unsafe_allow_html=True)
        with tc3:
            st.markdown(f"<div style='text-align: right;'><b>{travel_exp_sum:,.0f} 원</b></div>", unsafe_allow_html=True)
        with tc4:
            st.markdown("")

        st.markdown("---")

        # --- 3. 기타 섹션 (동적 행 관리) ---
        updated_other_rows = []
        other_sum = 0
        for idx, row_data in enumerate(st.session_state.other_rows):
            oc1, oc2, oc3, oc4 = st.columns(col_ratios)
            with oc1:
                if idx == 0:
                    st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>기타</b></div>", unsafe_allow_html=True)
                else:
                    st.markdown("")
            with oc2:
                it_name = st.text_input(
                    f"기타 항목명 {idx}",
                    value=row_data["item"],
                    placeholder="항목 입력",
                    key=f"other_item_{idx}",
                    label_visibility="collapsed",
                )
            with oc3:
                it_amt_str = st.text_input(
                    f"기타 금액 {idx}",
                    value=str(row_data["amount"]),
                    key=f"other_amt_{idx}",
                    label_visibility="collapsed",
                )
            with oc4:
                p_idx = (
                    payer_options.index(row_data["payer"])
                    if row_data["payer"] in payer_options
                    else 0
                )
                it_payer = st.selectbox(
                    f"기타 지급처 {idx}",
                    payer_options,
                    index=p_idx,
                    key=f"other_payer_{idx}",
                    label_visibility="collapsed",
                )

            try:
                parsed_amt = int(it_amt_str.replace(",", ""))
            except ValueError:
                parsed_amt = 0

            other_sum += parsed_amt
            updated_other_rows.append(
                {"item": it_name, "amount": parsed_amt, "payer": it_payer}
            )

        st.session_state.other_rows = updated_other_rows

        # 기타 마지막 행 아래에 행 추가/삭제 버튼 배치
        b_c1, b_c2, b_c3 = st.columns([1.2, 2.5, 2.5])
        with b_c1:
            st.markdown("")
        with b_c2:
            if st.form_submit_button("➕ 기타 행 추가", use_container_width=True):
                st.session_state.other_rows.append({"item": "", "amount": 0, "payer": "여행사"})
                st.rerun()
        with b_c3:
            if len(st.session_state.other_rows) > 1:
                if st.form_submit_button("➖ 기타 마지막 행 삭제", use_container_width=True):
                    st.session_state.other_rows.pop()
                    st.rerun()
            else:
                st.markdown("")

        # 기타 총계 행
        oc1, oc2, oc3, oc4 = st.columns(col_ratios)
        with oc1:
            st.markdown("")
        with oc2:
            st.markdown("<div style='text-align: right;'><b>기타 총계</b></div>", unsafe_allow_html=True)
        with oc3:
            st.markdown(f"<div style='text-align: right;'><b>{other_sum:,.0f} 원</b></div>", unsafe_allow_html=True)
        with oc4:
            st.markdown("")

        st.markdown("---")

        # --- 4. 총액 행 (전체 금액 합계) ---
        grand_total = transport_sum + travel_exp_sum + other_sum
        tot_c1, tot_c2, tot_c3, tot_c4 = st.columns(col_ratios)
        with tot_c1:
            st.markdown("")
        with tot_c2:
            st.markdown("<div style='text-align: right;'><span style='font-size: 1.1em;'><b>총액</b></span></div>", unsafe_allow_html=True)
        with tot_c3:
            st.markdown(f"<div style='text-align: right;'><span style='font-size: 1.1em; color: #ff4b4b;'><b>{grand_total:,.0f} 원</b></span></div>", unsafe_allow_html=True)
        with tot_c4:
            st.markdown("")

        st.markdown("---")
        submitted = st.form_submit_button(
            "➕ 입력한 출장 내역 규정 적용 및 추가", use_container_width=True
        )

        if submitted:
            if not name or not country:
                st.error("⚠️ 출장자 성명과 출장지는 필수 입력 항목입니다.")
            else:
                new_data = {
                    "출장자성명": name,
                    "부서": department,
                    "직급": position,
                    "직급구분": position_group,
                    "출장지": country,
                    "지역구분": region_group,
                    "출장시작일": str(start_date),
                    "출장종료일": str(end_date),
                    "출장일수": calculated_days,
                    "출장박수": calculated_nights,
                    "환율": exchange_rate,
                    "교통비항목리스트": st.session_state.transport_rows.copy(),
                    "출장비항목리스트": st.session_state.travel_exp_rows.copy(),
                    "기타항목리스트": st.session_state.other_rows.copy(),
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
        "출장자에게 송금할 금액과 여행사에 송금할 금액 등이 완벽히 분리된 자금팀 제출용 표입니다."
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
                "지역구분",
                "출장일수",
                "출장박수",
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
            "지역구분",
            "출장일수",
            "출장박수",
            "직원지급액",
            "여행사지급액",
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
            - **출장 기간**: {person_data.get('출장시작일', '')} ~ {person_data.get('출장종료일', '')} ({person_data.get('출장박수', 0)}박 {person_data.get('출장일수', 0)}일)
            - **적용 환율**: {person_data.get('환율', 0):,.2f} 원 {'(엔화 100 환산 적용)' if person_data.get('지역구분')=='특' else ''}
            - **직원 계좌 입금 총액**: {person_data.get('직원_계좌입금액', 0):,.0f} 원
            """
            )

        hotel_display = f"{person_data.get('산정숙박_원화', 0):,.0f} 원"
        currency_unit = "엔(¥)" if person_data.get("지역구분") == "특" else "달러($)"

        detail_table = pd.DataFrame(
            {
                "경비 항목": [
                    f"일비 ({person_data.get('출장일수', 0)}일)",
                    f"숙박비 ({person_data.get('출장박수', 0)}박)",
                    "대행 및 기타 경비",
                ],
                "규정 기준단가": [
                    f"{person_data.get('기준일당_외화', 0):,d} {currency_unit}/일",
                    f"{person_data.get('기준숙박_외화', 0):,d} {currency_unit}/박" if person_data.get("기준숙박_외화") != "실비" else "실비",
                    "실비 청구",
                ],
                "산정 금액 (원화)": [
                    f"{person_data.get('산정일당_원화', 0):,.0f} 원",
                    hotel_display,
                    f"{person_data.get('여행사_지급액', 0):,.0f} 원",
                ],
                "지급처 및 비고": [
                    "지급처 설정 반영",
                    "지급처 설정 반영",
                    "지급처 설정 REF",
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
