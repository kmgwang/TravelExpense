import datetime
import io
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import streamlit as st
import platform
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

if platform.system() == "Windows":
    matplotlib.rc("font", family="Malgun Gothic")
elif platform.system() == "Darwin":
    matplotlib.rc("font", family="AppleGothic")
else:
    matplotlib.rcParams["font.family"] = "NanumGothic"
matplotlib.rcParams["axes.unicode_minus"] = False

try:
    st.set_page_config(
        page_title="화천기공 해외출장비 정산 자동화 프로그램",
        page_layout="wide",
    )
except Exception:
    pass

st.markdown(
    """
    <style>
    .block-container {
        max-width: 95% !important;
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    input[aria-label*="항목"] {
        text-align: center !important;
    }
    input[aria-label*="금액"] {
        text-align: right !important;
    }
    .row-highlight-yellow {
        background-color: #fff9c4;
        padding: 8px 12px;
        border-radius: 4px;
        width: 100%;
    }
    .row-highlight-blue {
        background-color: #e1f5fe;
        padding: 10px 12px;
        border-radius: 4px;
        width: 100%;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("✈️ 화천기공 해외출장비 프로그램")
st.markdown(
    "인사지원팀 해외출장 경비 산정, 자금팀 제출용 정산표 분리, 출장자용 산정 내역서 자동 생성 프로그램입니다."
)
st.markdown("---")

if "travel_list" not in st.session_state:
    st.session_state.travel_list = []

if "edit_target_index" not in st.session_state:
    st.session_state.edit_target_index = None

if "transport_rows" not in st.session_state:
    st.session_state.transport_rows = [
        {"item": "항공권", "amount": 0, "payer": "여행사"},
        {"item": "ESTA", "amount": 0, "payer": "여행사"},
    ]

if "travel_exp_rows" not in st.session_state:
    st.session_state.travel_exp_rows = [
        {"item": "숙박비", "amount": 0, "payer": "출장자"},
        {"item": "일당", "amount": 0, "payer": "출장자"},
    ]

if "other_rows" not in st.session_state:
    st.session_state.other_rows = [{"item": "", "amount": 0, "payer": "여행사"}]


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

        raw_rate = d["환율"]
        if region == "특":
            applied_rate = raw_rate / 100.0
        else:
            applied_rate = raw_rate

        calc_daily = std_daily * applied_rate * d.get("출장일수", 1)
        calc_daily = int(calc_daily // 1000 * 1000)

        if std_hotel == "실비":
            calc_hotel = 0
        else:
            calc_hotel = std_hotel * applied_rate * d.get("출장박수", 0)
            calc_hotel = int(calc_hotel // 1000 * 1000)

        agency_total = 0
        employee_total = 0

        for te in d.get("출장비항목리스트", []):
            item_name = te.get("item", "")
            payer = te.get("payer", "출장자")
            if "일당" in item_name:
                amt = calc_daily
            elif "숙박" in item_name:
                amt = calc_hotel
            else:
                amt = te.get("amount", 0)

            if payer == "여행사":
                agency_total += amt
            else:
                employee_total += amt

        for t in d.get("교통비항목리스트", []):
            amt = t.get("amount", 0)
            payer = t.get("payer", "여행사")
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

        ordered_d = {
            "출장자성명": d.get("출장자성명"),
            "부서": d.get("부서"),
            "직급": d.get("직급"),
            "직급구분": d.get("직급구분"),
            "출장지": d.get("출장지"),
            "지역구분": d.get("지역구분"),
            "출장시작일": d.get("출장시작일"),
            "출장종료일": d.get("출장종료일"),
            "출장박수": d.get("출장박수"),
            "출장일수": d.get("출장일수"),
            "환율": d.get("환율"),
            "직원_계좌입금액": d.get("직원_계좌입금액"),
            "여행사_지급액": d.get("여행사_지급액"),
            "총출장비": d.get("총출장비"),
            "교통비항목리스트": d.get("교통비항목리스트"),
            "출장비항목리스트": d.get("출장비항목리스트"),
            "기타항목리스트": d.get("기타항목리스트"),
        }

        processed.append(ordered_d)

    df = pd.DataFrame(processed)
    return df


tab1, tab2, tab3 = st.tabs(
    [
        "📋 1. 출장 정보 입력",
        "💰 2. 자금팀 연결 자료 생성",
        "📄 3. 출장비 산정 내역서 생성",
    ]
)

with tab1:
    st.header("📋 해외출장비 산정")
    if st.session_state.edit_target_index is not None:
        st.info(
            f"✏️ 현재 **[인덱스 {st.session_state.edit_target_index}]** 번 출장 내역 수정 중입니다. 수정 후 아래 버튼을 누르면 내용이 갱신됩니다."
        )

    col_rate_info, col_rate_input = st.columns([2, 1])
    with col_rate_info:
        st.markdown(
            "🔗 [서울외국환중개 환율 조회 사이트 바로가기](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
        )

    target_edit_data = None
    if st.session_state.edit_target_index is not None and len(
        st.session_state.travel_list
    ) > st.session_state.edit_target_index:
        target_edit_data = st.session_state.travel_list[
            st.session_state.edit_target_index
        ]

    default_exchange_rate = (
        target_edit_data["환율"] if target_edit_data else 1350.0
    )
    with col_rate_input:
        exchange_rate = st.number_input(
            "적용 환율 입력",
            min_value=0.0,
            value=float(default_exchange_rate),
            step=1.0,
            format="%.2f",
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

    st.subheader("🌍 출장 정보 등록")
    col_a, col_b = st.columns(2)

    with col_a:
        default_name = (
            target_edit_data["출장자성명"] if target_edit_data else ""
        )
        name = st.text_input("출장자 성명", value=default_name)

        default_dept = target_edit_data["부서"] if target_edit_data else "인사지원팀"
        dept_idx = (
            department_list.index(default_dept)
            if default_dept in department_list
            else 0
        )
        department = st.selectbox("부서", department_list, index=dept_idx)

        position_list = [
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
        ]
        default_pos = target_edit_data["직급"] if target_edit_data else "사원"
        pos_idx = (
            position_list.index(default_pos)
            if default_pos in position_list
            else 10
        )
        position = st.selectbox("직급", position_list, index=pos_idx)

        auto_pos_group = get_position_group(position)
        pos_group_options = [
            "임원(부사장이상)",
            "임원",
            "1급",
            "2급",
            "3급이하",
        ]

        if target_edit_data and "loaded_edit_idx" not in st.session_state:
            default_pos_group = target_edit_data.get("직급구분", auto_pos_group)
        else:
            default_pos_group = auto_pos_group

        default_pos_idx = (
            pos_group_options.index(default_pos_group)
            if default_pos_group in pos_group_options
            else 4
        )
        position_group = st.selectbox(
            "직급 구분", pos_group_options, index=default_pos_idx
        )

    with col_b:
        default_start = (
            datetime.date.fromisoformat(target_edit_data["출장시작일"])
            if target_edit_data
            else datetime.date.today() + datetime.timedelta(days=1)
        )
        default_end = (
            datetime.date.fromisoformat(target_edit_data["출장종료일"])
            if target_edit_data
            else default_start + datetime.timedelta(days=7)
        )

        start_date = st.date_input("출장 시작일", value=default_start)
        end_date = st.date_input("출장 종료일", value=default_end)

        default_country = target_edit_data["출장지"] if target_edit_data else ""
        country = st.text_input("출장지", value=default_country)

        auto_region = get_region_group(country) if country else "갑"
        region_options = ["갑", "을", "병", "특"]

        if target_edit_data and "loaded_edit_idx" not in st.session_state:
            default_region = target_edit_data.get("지역구분", auto_region)
        else:
            default_region = auto_region

        default_reg_idx = (
            region_options.index(default_region)
            if default_region in region_options
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
        calculated_nights = calculated_days - 1 if calculated_days > 1 else 0
        if is_flight_minus:
            calculated_nights = max(0, calculated_nights - 1)

        st.info(
            f"📅 최종 산정된 출장 기간: **{calculated_nights}박 {calculated_days}일**"
        )

    if target_edit_data and "loaded_edit_idx" not in st.session_state:
        st.session_state.transport_rows = target_edit_data.get(
            "교통비항목리스트", st.session_state.transport_rows
        )
        st.session_state.travel_exp_rows = target_edit_data.get(
            "출장비항목리스트", st.session_state.travel_exp_rows
        )
        st.session_state.other_rows = target_edit_data.get(
            "기타항목리스트", st.session_state.other_rows
        )
        st.session_state.loaded_edit_idx = st.session_state.edit_target_index

    std_daily, std_hotel = get_standard_rates(region_group, position_group)
    applied_rate = (
        exchange_rate / 100.0 if region_group == "특" else exchange_rate
    )

    auto_calc_daily = int(
        (std_daily * applied_rate * calculated_days) // 1000 * 1000
    )
    if std_hotel == "실비":
        auto_calc_hotel = 0
    else:
        auto_calc_hotel = int(
            (std_hotel * applied_rate * calculated_nights) // 1000 * 1000
        )

    if target_edit_data is None:
        if (
            "prev_auto_calc_hotel" not in st.session_state
            or st.session_state.prev_auto_calc_hotel != auto_calc_hotel
        ):
            st.session_state.prev_auto_calc_hotel = auto_calc_hotel
            st.session_state["te_amt_str_0"] = f"{auto_calc_hotel:,}"

        if (
            "prev_auto_calc_daily" not in st.session_state
            or st.session_state.prev_auto_calc_daily != auto_calc_daily
        ):
            st.session_state.prev_auto_calc_daily = auto_calc_daily
            st.session_state["te_amt_str_1"] = f"{auto_calc_daily:,}"

    st.markdown("---")
    st.subheader("💵 금액 입력")
    st.markdown("<br>", unsafe_allow_html=True)

    col_ratios = [1.2, 1.5, 2, 1.5]

    th1, th2, th3, th4 = st.columns(col_ratios)
    with th1:
        st.markdown(
            "<div style='text-align: center;'><b>구분</b></div>",
            unsafe_allow_html=True,
        )
    with th2:
        st.markdown(
            "<div style='text-align: center;'><b>항목</b></div>",
            unsafe_allow_html=True,
        )
    with th3:
        st.markdown(
            "<div style='text-align: center;'><b>금액</b></div>",
            unsafe_allow_html=True,
        )
    with th4:
        st.markdown(
            "<div style='text-align: center;'><b>지급처</b></div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    payer_options = ["여행사", "출장자", "직접입력"]

    updated_transport_rows = []
    transport_sum = 0
    for idx, row_data in enumerate(st.session_state.transport_rows):
        tc1, tc2, tc3, tc4 = st.columns(col_ratios)
        with tc1:
            st.markdown(
                "<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>교통비</b></div>",
                unsafe_allow_html=True,
            )
        with tc2:
            t_name = st.text_input(
                f"교통비 항목 {idx}",
                value=row_data["item"],
                key=f"t_item_{idx}",
                label_visibility="collapsed",
            )
        with tc3:
            key_prefix = f"t_amt_str_{idx}"
            if key_prefix not in st.session_state:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"


            def make_on_change(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = (
                        f"{int(digits):,}" if digits else "0"
                    )

                return callback


            amt_str = st.text_input(
                f"교통비 금액 {idx}",
                key=key_prefix,
                on_change=make_on_change(key_prefix),
                label_visibility="collapsed",
            )
            digits = "".join(filter(str.isdigit, amt_str))
            t_amt = int((int(digits or 0) // 1000) * 1000)
        with tc4:
            p_idx = (
                payer_options.index(row_data["payer"])
                if row_data["payer"] in payer_options
                else 0
            )
            t_payer = st.selectbox(
                f"교통비 지급처 {idx}",
                payer_options,
                index=p_idx,
                key=f"t_payer_{idx}",
                label_visibility="collapsed",
            )

        transport_sum += t_amt
        updated_transport_rows.append(
            {"item": t_name, "amount": t_amt, "payer": t_payer}
        )

    st.session_state.transport_rows = updated_transport_rows

    sc1, sc2, sc3, sc4 = st.columns(col_ratios)
    with sc1:
        st.markdown("")
    with sc2:
        st.markdown(
            "<div class='row-highlight-yellow' style='text-align: center;'><b>교통비 총계</b></div>",
            unsafe_allow_html=True,
        )
    with sc3:
        st.markdown(
            f"<div class='row-highlight-yellow' style='text-align: center;'><b>{transport_sum:,.0f} 원</b></div>",
            unsafe_allow_html=True,
        )
    with sc4:
        st.markdown("")

    st.markdown("---")

    updated_travel_exp_rows = []
    travel_exp_sum = 0
    for idx, row_data in enumerate(st.session_state.travel_exp_rows):
        tec1, tec2, tec3, tec4 = st.columns(col_ratios)
        with tec1:
            st.markdown(
                "<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>출장비</b></div>",
                unsafe_allow_html=True,
            )
        with tec2:
            te_name = st.text_input(
                f"출장비 항목 {idx}",
                value=row_data["item"],
                key=f"te_item_{idx}",
                label_visibility="collapsed",
            )
        with tec3:
            key_prefix = f"te_amt_str_{idx}"
            if key_prefix not in st.session_state and target_edit_data:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"


            def make_on_change_te(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = (
                        f"{int(digits):,}" if digits else "0"
                    )

                return callback


            amt_str = st.text_input(
                f"출장비 금액 {idx}",
                key=key_prefix,
                on_change=make_on_change_te(key_prefix),
                label_visibility="collapsed",
            )
            digits = "".join(filter(str.isdigit, amt_str))
            te_amt = int((int(digits or 0) // 1000) * 1000)
        with tec4:
            p_idx = (
                payer_options.index(row_data["payer"])
                if row_data["payer"] in payer_options
                else 1
            )
            te_payer = st.selectbox(
                f"출장비 지급처 {idx}",
                payer_options,
                index=p_idx,
                key=f"te_payer_{idx}",
                label_visibility="collapsed",
            )

        travel_exp_sum += te_amt
        updated_travel_exp_rows.append(
            {"item": te_name, "amount": te_amt, "payer": te_payer}
        )

    st.session_state.travel_exp_rows = updated_travel_exp_rows

    tc1, tc2, tc3, tc4 = st.columns(col_ratios)
    with tc1:
        st.markdown("")
    with tc2:
        st.markdown(
            "<div class='row-highlight-yellow' style='text-align: center;'><b>출장비 총계</b></div>",
            unsafe_allow_html=True,
        )
    with tc3:
        st.markdown(
            f"<div class='row-highlight-yellow' style='text-align: center;'><b>{travel_exp_sum:,.0f} 원</b></div>",
            unsafe_allow_html=True,
        )
    with tc4:
        st.markdown("")

    st.markdown("---")

    updated_other_rows = []
    other_sum = 0
    for idx, row_data in enumerate(st.session_state.other_rows):
        oc1, oc2, oc3, oc4 = st.columns(col_ratios)
        with oc1:
            st.markdown(
                "<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>기타</b></div>",
                unsafe_allow_html=True,
            )
        with oc2:
            it_name = st.text_input(
                f"기타 항목명 {idx}",
                value=row_data["item"],
                placeholder="항목 입력",
                key=f"other_item_{idx}",
                label_visibility="collapsed",
            )
        with oc3:
            key_prefix = f"other_amt_str_{idx}"
            if key_prefix not in st.session_state:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"


            def make_on_change_other(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = (
                        f"{int(digits):,}" if digits else "0"
                    )

                return callback


            amt_str = st.text_input(
                f"기타 금액 {idx}",
                key=key_prefix,
                on_change=make_on_change_other(key_prefix),
                label_visibility="collapsed",
            )
            digits = "".join(filter(str.isdigit, amt_str))
            it_amt = int((int(digits or 0) // 1000) * 1000)
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

        other_sum += it_amt
        updated_other_rows.append(
            {"item": it_name, "amount": it_amt, "payer": it_payer}
        )

    st.session_state.other_rows = updated_other_rows

    b_c1, b_c2, b_c3 = st.columns([1.2, 2.5, 2.5])
    with b_c1:
        st.markdown("")
    with b_c2:
        if st.button("➕ 기타 행 추가", use_container_width=True):
            st.session_state.other_rows.append(
                {"item": "", "amount": 0, "payer": "여행사"}
            )
            st.rerun()
    with b_c3:
        if len(st.session_state.other_rows) > 1:
            if st.button("➖ 기타 마지막 행 삭제", use_container_width=True):
                st.session_state.other_rows.pop()
                st.rerun()
        else:
            st.markdown("")

    oc1, oc2, oc3, oc4 = st.columns(col_ratios)
    with oc1:
        st.markdown("")
    with oc2:
        st.markdown(
            "<div class='row-highlight-yellow' style='text-align: center;'><b>기타 총계</b></div>",
            unsafe_allow_html=True,
        )
    with oc3:
        st.markdown(
            f"<div class='row-highlight-yellow' style='text-align: center;'><b>{other_sum:,.0f} 원</b></div>",
            unsafe_allow_html=True,
        )
    with oc4:
        st.markdown("")

    st.markdown("---")

    real_employee_total = 0
    real_agency_total = 0

    for te_item in st.session_state.travel_exp_rows:
        amt = te_item.get("amount", 0)
        if te_item.get("payer") == "여행사":
            real_agency_total += amt
        else:
            real_employee_total += amt

    for t_item in st.session_state.transport_rows:
        amt = t_item.get("amount", 0)
        if t_item.get("payer") == "여행사":
            real_agency_total += amt
        else:
            real_employee_total += amt

    for o_item in st.session_state.other_rows:
        amt = o_item.get("amount", 0)
        if o_item.get("payer") == "여행사":
            real_agency_total += amt
        else:
            real_employee_total += amt

    grand_total = real_employee_total + real_agency_total

    tot_c1, tot_c2, tot_c3, tot_c4 = st.columns(col_ratios)
    with tot_c1:
        st.markdown("")
    with tot_c2:
        st.markdown(
            "<div class='row-highlight-blue' style='text-align: center;'><span style='font-size: 1.1em;'><b>총액</b></span></div>",
            unsafe_allow_html=True,
        )
    with tot_c3:
        st.markdown(
            f"<div class='row-highlight-blue' style='text-align: center;'><span style='font-size: 1.1em; color: #000000;'><b>{grand_total:,.0f} 원</b></span></div>",
            unsafe_allow_html=True,
        )
    with tot_c4:
        st.markdown("")

    st.markdown("---")
    btn_label = (
        "🔄 수정 사항 반영하기"
        if st.session_state.edit_target_index is not None
        else "➕ 입력한 출장 내역 규정 적용 및 추가"
    )
    submitted = st.button(btn_label, use_container_width=True)

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
                "직원_계좌입금액": real_employee_total,
                "여행사_지급액": real_agency_total,
                "총출장비": grand_total,
                "교통비항목리스트": st.session_state.transport_rows.copy(),
                "출장비항목리스트": st.session_state.travel_exp_rows.copy(),
                "기타항목리스트": st.session_state.other_rows.copy(),
            }

            if st.session_state.edit_target_index is not None:
                st.session_state.travel_list[
                    st.session_state.edit_target_index
                ] = new_data
                st.success(
                    f"✅ [{name}] 님의 출장 내역이 성공적으로 수정(갱신)되었습니다!"
                )
                st.session_state.edit_target_index = None
                if "loaded_edit_idx" in st.session_state:
                    del st.session_state["loaded_edit_idx"]
            else:
                st.session_state.travel_list.append(new_data)
                st.success(
                    f"✅ {name} 님의 출장 경비가 규정에 맞춰 산정되었습니다!"
                )
            st.rerun()

    st.markdown("---")
    st.subheader("📊 현재 등록된 전체 출장 내역 목록")
    st.caption("💡 팁: 아래 표의 행을 **더블클릭**하면 해당 내역을 다시 수정할 수 있습니다.")

    if len(st.session_state.travel_list) > 0:
        raw_df = process_travel_data(st.session_state.travel_list)

        display_df = raw_df.drop(
            columns=[
                "교통비항목리스트",
                "출장비항목리스트",
                "기타항목리스트",
                "지역구분",
                "직급구분",
                "환율",
            ]
        ).reset_index(drop=True).copy()

        display_df.insert(0, "순번", range(1, len(display_df) + 1))

        if "출장지" in display_df.columns and "순번" in display_df.columns:
            cols = list(display_df.columns)
            cols.remove("출장지")
            seq_idx = cols.index("순번")
            cols.insert(seq_idx + 1, "출장지")
            display_df = display_df[cols]

        if "출장박수" in display_df.columns:
            display_df["출장박수"] = display_df["출장박수"].apply(
                lambda x: f"{int(x):,}"
            )
        if "출장일수" in display_df.columns:
            display_df["출장일수"] = display_df["출장일수"].apply(
                lambda x: f"{int(x):,}"
            )
        if "직원_계좌입금액" in display_df.columns:
            display_df["직원_계좌입금액"] = display_df[
                "직원_계좌입금액"
            ].apply(lambda x: f"{int(x):,}")
        if "여행사_지급액" in display_df.columns:
            display_df["여행사_지급액"] = display_df["여행사_지급액"].apply(
                lambda x: f"{int(x):,}"
            )
        if "총출장비" in display_df.columns:
            display_df["총출장비"] = display_df["총출장비"].apply(
                lambda x: f"{int(x):,}"
            )

        gb = GridOptionsBuilder.from_dataframe(display_df)
        gb.configure_selection(
            selection_mode="single",
            use_checkbox=False,
            rowMultiSelectWithClick=False,
        )
        gb.configure_default_column(
            filterable=False, sortable=True, resizable=True
        )
        gb.configure_grid_options(
            rowSelection="single", suppressRowClickSelection=False
        )
        grid_options = gb.build()

        grid_response = AgGrid(
            display_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            fit_columns_on_grid_load=True,
            height=250,
            theme="balham",
        )

        selected_rows = grid_response.get("selected_rows", None)

        if selected_rows is not None:
            if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
                selected_idx = int(selected_rows.index[0])
            elif isinstance(selected_rows, list) and len(selected_rows) > 0:
                sel_row_dict = selected_rows[0]
                matched = display_df[
                    (display_df["출장자성명"] == sel_row_dict.get("출장자성명"))
                    & (display_df["출장지"] == sel_row_dict.get("출장지"))
                    & (display_df["출장시작일"] == sel_row_dict.get("출장시작일"))
                ]
                if not matched.empty:
                    selected_idx = int(matched.index[0])
                else:
                    selected_idx = None
            else:
                selected_idx = None

            if selected_idx is not None and st.session_state.edit_target_index is None:
                st.session_state.edit_target_index = selected_idx
                if "loaded_edit_idx" in st.session_state:
                    del st.session_state["loaded_edit_idx"]
                st.success(
                    f"📌 [{display_df.iloc[selected_idx]['출장자성명']}] 님의 내역이 선택되었습니다. 위쪽 입력 폼에서 내용을 수정한 뒤 '수정 사항 반영하기' 버튼을 눌러주세요."
                )
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        col_del1, col_del2 = st.columns([1, 1])
        with col_del1:
            if st.session_state.edit_target_index is not None:
                if st.button("❌ 수정 모드 취소"):
                    st.session_state.edit_target_index = None
                    if "loaded_edit_idx" in st.session_state:
                        del st.session_state["loaded_edit_idx"]
                    st.rerun()
        with col_del2:
            if st.button("🗑️ 전체 데이터 초기화"):
                st.session_state.travel_list = []
                st.session_state.edit_target_index = None
                if "loaded_edit_idx" in st.session_state:
                    del st.session_state["loaded_edit_idx"]
                st.rerun()
    else:
        st.info("등록된 출장 내역이 없습니다.")

with tab2:
    st.header("💰 자금팀 제출용 정산 집계표 생성")
    st.markdown(
        "출장자에게 송금할 금액과 여행사에 송금할 금액 등이 완벽히 분리된 자금팀 제출용 표입니다."
    )

    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric(label="총 출장 건수", value=f"{len(processed_df)} 건")
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
                "출장지",
                "직원_계좌입금액",
                "여행사_지급액",
                "총출장비",
            ]
        ].copy()

        fund_view_df.insert(0, "순번", range(1, len(fund_view_df) + 1))

        if "출장지" in fund_view_df.columns and "순번" in fund_view_df.columns:
            cols = list(fund_view_df.columns)
            cols.remove("출장지")
            seq_idx = cols.index("순번")
            cols.insert(seq_idx + 1, "출장지")
            fund_view_df = fund_view_df[cols]

        fund_view_df.columns = [
            "순번",
            "출장지",
            "성명",
            "부서",
            "직급",
            "직원지급액",
            "여행사지급액",
            "총합계",
        ]
        st.dataframe(fund_view_df, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        output_agency = io.BytesIO()
        excel_save_df = processed_df.drop(
            columns=[
                "직급구분",
                "지역구분",
                "환율",
                "교통비항목리스트",
                "출장비항목리스트",
                "기타항목리스트",
                "출장시작일",
                "출장종료일",
                "출장박수",
                "출장일수",
            ]
        )
        excel_save_df.insert(0, "순번", range(1, len(excel_save_df) + 1))

        if "출장지" in excel_save_df.columns and "순번" in excel_save_df.columns:
            cols = list(excel_save_df.columns)
            cols.remove("출장지")
            seq_idx = cols.index("순번")
            cols.insert(seq_idx + 1, "출장지")
            excel_save_df = excel_save_df[cols]

        with pd.ExcelWriter(output_agency, engine="openpyxl") as writer:
            excel_save_df.to_excel(
                writer, index=False, sheet_name="자금팀_정산집계표"
            )

            worksheet = writer.sheets["자금팀_정산집계표"]

            fill_sky_blue = PatternFill(
                start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
            )
            font_mild = Font(
                name="맑은 고딕", size=10, bold=False, color="333333"
            )

            fill_fg_strong = PatternFill(
                start_color="1F4E78", end_color="1F4E78", fill_type="solid"
            )
            font_strong = Font(
                name="맑은 고딕", size=10, bold=True, color="FFFFFF"
            )

            thick_side = Side(style="medium", color="000000")
            thin_side = Side(style="thin", color="000000")
            dashed_side = Side(style="dashed", color="000000")

            max_row = len(excel_save_df) + 1
            max_col = len(excel_save_df.columns)

            for row_idx in range(1, max_row + 1):
                worksheet.row_dimensions[row_idx].height = 35.0

                for col_idx in range(1, max_col + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)

                    if col_idx in [6, 7]:
                        cell.fill = fill_sky_blue

                    if row_idx == 1:
                        top_b = thick_side
                        bottom_b = thick_side
                    else:
                        top_b = dashed_side
                        bottom_b = (
                            thick_side if row_idx == max_row else dashed_side
                        )

                    left_b = thick_side if col_idx == 1 else thin_side
                    right_b = thick_side if col_idx == max_col else thin_side

                    cell.border = Border(
                        top=top_b, bottom=bottom_b, left=left_b, right=right_b
                    )

                    if col_idx <= 5:
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
                    else:
                        if row_idx == 1:
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )
                            if col_idx == 8:
                                cell.fill = fill_fg_strong
                                cell.font = Font(
                                    name="맑은 고딕",
                                    size=11,
                                    bold=True,
                                    color="FFFFFF",
                                )
                        else:
                            cell.alignment = Alignment(
                                horizontal="right", vertical="center"
                            )

                            if col_idx in [6, 7]:
                                cell.font = font_mild
                            elif col_idx == 8:
                                cell.fill = fill_fg_strong
                                cell.font = font_strong

                    col_name = excel_save_df.columns[col_idx - 1]
                    if any(k in col_name for k in ["금액", "지급액", "총출장비"]):
                        cell.number_format = "#,##0"

            for col in worksheet.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                worksheet.column_dimensions[col_letter].width = max(
                    max_length + 5, 14
                )

        output_agency.seek(0)

        st.download_button(
            label="📥 자금팀 정산 집계표 엑셀 다운로드",
            data=output_agency,
            file_name="화천기공_자금팀_출장정산집계표.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning(
            "⚠️ 입력된 출장 정보가 없습니다. [1. 출장 정보 입력] 탭에서 데이터를 먼저 입력해 주세요."
        )

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
            - **출장지**: {person_data.get('출장지', '-')} (지역: {person_data.get('지역구분', '-')})
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

        def generate_exact_statement_excel(p_data):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "산정내역서"

            # 1. 인쇄 영역 설정 (A1:F36)
            ws.page_setup.printArea = "A1:F36"
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize = ws.PAPERSIZE_A4

            ws.views.sheetView[0].showGridLines = True

            font_title = Font(name="맑은 고딕", size=16, bold=True)
            font_bold = Font(name="맑은 고딕", size=10, bold=True)
            font_normal = Font(name="맑은 고딕", size=10, bold=False)

            fill_gray_header = PatternFill(
                start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
            )
            thin_border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )

            ws.merge_cells("A1:E1")
            cell_t = ws["A1"]
            cell_t.value = "해외출장비 산정 내역서"
            cell_t.font = font_title
            cell_t.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 40

            # 2. 지역구분 값 ('급' 제거, 갑/을/병/특 만 표시)
            region_val = p_data.get("지역구분", "")

            info_rows = [
                [
                    "소속",
                    p_data.get("부서", ""),
                    "성명",
                    p_data.get("출장자성명", ""),
                    "직급",
                    p_data.get("직급", ""),
                ],
                [
                    "출장지",
                    p_data.get("출장지", ""),
                    "지역구분",
                    region_val,
                    "직급구분",
                    p_data.get("직급구분", ""),
                ],
                [
                    "출발일",
                    p_data.get("출장시작일", ""),
                    "도착일",
                    p_data.get("출장종료일", ""),
                    "출장기간",
                    f"{p_data.get('출장박수', 0)}박 {p_data.get('출장일수', 0)}일",
                ],
            ]

            for r_idx, r_data in enumerate(info_rows, start=3):
                ws.row_dimensions[r_idx].height = 22
                ws.cell(row=r_idx, column=1, value=r_data[0])
                ws.cell(row=r_idx, column=2, value=r_data[1])
                ws.cell(row=r_idx, column=3, value=r_data[2])
                ws.cell(row=r_idx, column=4, value=r_data[3])
                ws.cell(row=r_idx, column=5, value=r_data[4])
                ws.cell(row=r_idx, column=6, value=r_data[5])

                for c_idx in range(1, 7):
                    c = ws.cell(row=r_idx, column=c_idx)
                    c.border = thin_border
                    c.font = font_normal
                    if c_idx in [1, 3, 5]:
                        c.fill = fill_gray_header
                        c.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
                        c.font = font_bold
                    else:
                        c.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )

            ws["A6"] = "■ 적용 출장비 산정 기준"
            ws["A6"].font = font_bold
            ws.row_dimensions[6].height = 25

            headers_1 = [
                "구분",
                "산정 기준",
                "기간 적용",
                "금액",
                "원화 환산 산식",
            ]
            for c_idx, h_text in enumerate(headers_1, start=1):
                cell = ws.cell(row=7, column=c_idx, value=h_text)
                cell.font = font_bold
                cell.fill = fill_gray_header
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                cell.border = thin_border
            ws.row_dimensions[7].height = 22

            pos_group = p_data.get("직급구분", "3급이하")
            std_daily, std_hotel = get_standard_rates(region_val, pos_group)
            curr_symbol = "¥" if region_val == "특" else "USD"

            hotel_str = (
                "실비"
                if std_hotel == "실비"
                else f"{std_hotel}{curr_symbol} / 박"
            )
            daily_str = f"{std_daily}{curr_symbol} / 일"

            n_nights = p_data.get("출장박수", 0)
            n_days = p_data.get("출장일수", 0)
            raw_rate = p_data.get("환율", 0)
            applied_rate = raw_rate / 100.0 if region_val == "특" else raw_rate

            # 3. 숙박비 값인 D8셀에 자동 산정된 숙박비 값 입력
            if std_hotel == "실비":
                calc_hotel = 0
            else:
                calc_hotel = int((std_hotel * applied_rate * n_nights) // 1000 * 1000)

            # 5. 원화 환산 산식 (E8, E9에 산정 기준 * 기간 적용 * 환율 반영)
            hotel_formula_text = (
                "실비"
                if std_hotel == "실비"
                else f"산정기준({std_hotel}) * 박수({n_nights}) * 환율({applied_rate:,.2f})"
            )

            row_hotel = [
                "숙박비",
                hotel_str,
                f"{n_nights}박",
                calc_hotel,
                hotel_formula_text,
            ]
            for c_idx, val in enumerate(row_hotel, start=1):
                cell = ws.cell(row=8, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_normal
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                if c_idx == 4 and isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
            ws.row_dimensions[8].height = 22

            # 4. 일당 값인 D9셀에 자동 산정된 일당 값 입력
            calc_daily = int((std_daily * applied_rate * n_days) // 1000 * 1000)
            daily_formula_text = f"산정기준({std_daily}) * 일수({n_days}) * 환율({applied_rate:,.2f})"

            row_daily = [
                "일당",
                daily_str,
                f"{n_days}일",
                calc_daily,
                daily_formula_text,
            ]
            for c_idx, val in enumerate(row_daily, start=1):
                cell = ws.cell(row=9, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = font_normal
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                if c_idx == 4 and isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
            ws.row_dimensions[9].height = 22

            ws.cell(row=10, column=1, value="지급 총액").font = font_bold
            ws.cell(row=10, column=1).fill = fill_gray_header
            ws.cell(row=10, column=1).alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.cell(row=10, column=1).border = thin_border

            ws.merge_cells("B10:C10")
            ws.cell(row=10, column=2, value="숙박비 + 일당").font = font_bold
            ws.cell(row=10, column=2).alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.cell(row=10, column=2).border = thin_border
            ws.cell(row=10, column=3).border = thin_border

            ws.merge_cells("D10:E10")
            total_calc_amt = p_data.get("직원_계좌입금액", 0)
            ws.cell(
                row=10, column=4, value=total_calc_amt
            ).font = font_bold
            ws.cell(row=10, column=4).alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.cell(row=10, column=4).border = thin_border
            ws.cell(row=10, column=4).number_format = "#,##0"
            ws.cell(row=10, column=5).border = thin_border

            ws.cell(row=10, column=6, value="").border = thin_border
            ws.row_dimensions[10].height = 22

            ws["A12"] = "■ 해외출장 지급규정"
            ws["A12"].font = font_bold
            ws.row_dimensions[12].height = 25

            headers_2 = ["지역", "직급", "일당", "숙박", "비고"]
            ws.merge_cells("D13:E13")
            ws.cell(row=13, column=1, value=headers_2[0])
            ws.cell(row=13, column=2, value=headers_2[1])
            ws.cell(row=13, column=3, value=headers_2[2])
            ws.cell(row=13, column=4, value=headers_2[3])
            ws.cell(row=13, column=6, value=headers_2[4])

            for c_idx in [1, 2, 3, 4, 5, 6]:
                cell = ws.cell(row=13, column=c_idx)
                cell.font = font_bold
                cell.fill = fill_gray_header
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                cell.border = thin_border
            ws.row_dimensions[13].height = 22

            rules_data = [
                (
                    "갑",
                    [
                        ("임원(부사장이상)", "$135", "실비", ""),
                        ("임원", "$90", "$130", ""),
                        ("1급", "$70", "$100", ""),
                        ("2급", "$65", "$95", ""),
                        ("3급이하", "$60", "$90", ""),
                    ],
                ),
                (
                    "을",
                    [
                        ("임원(부사장이상)", "$130", "실비", ""),
                        ("임원", "$85", "$125", ""),
                        ("1급", "$65", "$95", ""),
                        ("2급", "$60", "$90", ""),
                        ("3급이하", "$55", "$85", ""),
                    ],
                ),
                (
                    "병",
                    [
                        ("임원(부사장이상)", "$130", "실비", ""),
                        ("임원", "$80", "$110", ""),
                        ("1급", "$60", "$90", ""),
                        ("2급", "$55", "$85", ""),
                        ("3급이하", "$55", "$80", ""),
                    ],
                ),
                (
                    "특",
                    [
                        ("임원(부사장이상)", "¥23,000", "실비", ""),
                        ("임원", "¥11,000", "¥17,000", "엔화"),
                        ("1급", "¥8,000", "¥12,000", "엔화"),
                        ("2급", "¥7,000", "¥11,000", "엔화"),
                        ("3급이하", "¥7,000", "¥10,000", "엔화"),
                    ],
                ),
            ]

            curr_row = 14
            for reg_name, pos_list in rules_data:
                start_r = curr_row
                for pos_name, d_val, h_val, note_val in pos_list:
                    ws.cell(row=curr_row, column=2, value=pos_name)
                    ws.cell(row=curr_row, column=3, value=d_val)
                    ws.merge_cells(
                        start_row=curr_row,
                        start_column=4,
                        end_row=curr_row,
                        end_column=5,
                    )
                    ws.cell(row=curr_row, column=4, value=h_val)
                    ws.cell(row=curr_row, column=6, value=note_val)

                    for c_idx in range(1, 7):
                        cell = ws.cell(row=curr_row, column=c_idx)
                        cell.border = thin_border
                        cell.font = font_normal
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )
                    ws.row_dimensions[curr_row].height = 20
                    curr_row += 1

                end_r = curr_row - 1
                if start_r != end_r:
                    ws.merge_cells(
                        start_row=start_r,
                        start_column=1,
                        end_row=end_r,
                        end_column=1,
                    )
                ws.cell(row=start_r, column=1, value=reg_name)
                ws.cell(row=start_r, column=1).alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                ws.cell(row=start_r, column=1).font = font_bold

            footer_row_1 = curr_row + 1
            ws.merge_cells(
                start_row=footer_row_1,
                start_column=1,
                end_row=footer_row_1,
                end_column=6,
            )
            f_cell1 = ws.cell(
                row=footer_row_1,
                column=1,
                value="위와 같이 해외출장비를 정산 및 지급합니다.",
            )
            f_cell1.font = font_bold
            f_cell1.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.row_dimensions[footer_row_1].height = 25

            footer_row_2 = footer_row_1 + 1
            today_str = datetime.date.today().strftime("%Y년 %m월 %d일")
            ws.merge_cells(
                start_row=footer_row_2,
                start_column=1,
                end_row=footer_row_2,
                end_column=6,
            )
            f_cell2 = ws.cell(
                row=footer_row_2,
                column=1,
                value=f"신청일 : {today_str}",
            )
            f_cell2.font = font_normal
            f_cell2.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws.row_dimensions[footer_row_2].height = 22

            col_widths = {
                "A": 16,
                "B": 18,
                "C": 14,
                "D": 16,
                "E": 28,
                "F": 16,
            }
            for col_letter, width in col_widths.items():
                ws.column_dimensions[col_letter].width = width

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output

        output_person = generate_exact_statement_excel(person_data)

        st.download_button(
            label=f"📥 [{selected_person}] 출장자용 산정 내역서 엑셀 다운로드 (양식 적용)",
            data=output_person,
            file_name=f"화천기공_해외출장산정내역서_{selected_person}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning(
            "⚠️ 입력된 출장 정보가 없습니다. [1. 출장 정보 입력] 탭에서 데이터를 먼저 입력해 주세요."
        )
