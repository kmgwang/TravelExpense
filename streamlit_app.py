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

# 사이드바 설정: 웹에서 직접 출장 정보 입력
st.sidebar.header("📝 출장 정보 직접 입력")

with st.sidebar.form("travel_form"):
    st.subheader("출장자 및 기본 정보")
    name = st.text_input("출장자 성명", value="홍길동")
    department = st.text_input("부서", value="인사지원팀")
    position = st.text_input("직급", value="대리")
    account = st.text_input("계좌번호", value="우리은행 123-***")

    st.subheader("출장지 및 일정")
    country = st.text_input("출장 국가 / 도시", value="미국 (LA)")
    start_date = st.date_input("출장 시작일")
    end_date = st.date_input("출장 종료일")

    st.subheader("경비 항목 입력 (원화 기준)")
    flight = st.number_input(
        "항공료 (여행사 지급)", min_value=0, value=2500000, step=10000
    )
    hotel = st.number_input(
        "숙박비 (여행사 지급)", min_value=0, value=1400000, step=10000
    )
    daily_allowance = st.number_input(
        "일비 (직원 지급)", min_value=0, value=350000, step=5000
    )
    meal = st.number_input(
        "식비 (직원 지급)", min_value=0, value=490000, step=5000
    )
    exchange_rate = st.number_input(
        "적용 환율", min_value=0.0, value=1350.0, step=10.0
    )

    submitted = st.form_submit_button("➕ 출장 내역 추가 / 계산")

# 세션 스테이트를 이용해 입력된 출장 데이터 누적 관리
if "travel_list" not in st.session_state:
    st.session_state.travel_list = [
        {
            "출장자성명": "홍길동",
            "부서": "인사지원팀",
            "직급": "대리",
            "계좌번호": "우리은행 123-***",
            "출장국가": "미국 (LA)",
            "출장시작일": "2026-09-10",
            "출장종료일": "2026-09-17",
            "항공료_여행사지급": 2500000,
            "숙박비_여행사지급": 1400000,
            "일비_직인지급": 350000,
            "식비_직인지급": 490000,
            "환율": 1350.0,
        },
        {
            "출장자성명": "김철수",
            "부서": "생산관리팀",
            "직급": "과장",
            "계좌번호": "국민은행 456-***",
            "출장국가": "베트남 (하노이)",
            "출장시작일": "2026-09-15",
            "출장종료일": "2026-09-20",
            "항공료_여행사지급": 1200000,
            "숙박비_여행사지급": 800000,
            "일비_직인지급": 250000,
            "식비_직인지급": 350000,
            "환율": 1350.0,
        },
    ]

if submitted:
    new_data = {
        "출장자성명": name,
        "부서": department,
        "직급": position,
        "계좌번호": account,
        "출장국가": country,
        "출장시작일": str(start_date),
        "출장종료일": str(end_date),
        "항공료_여행사지급": flight,
        "숙박비_여행사지급": hotel,
        "일비_직인지급": daily_allowance,
        "식비_직인지급": meal,
        "환율": exchange_rate,
    }
    st.session_state.travel_list.append(new_data)
    st.sidebar.success(f"{name} 님의 출장 내역이 추가되었습니다!")


def process_travel_data(data_list):
    df = pd.DataFrame(data_list)
    if not df.empty:
        df["총출장비"] = (
            df["항공료_여행사지급"]
            + df["숙박비_여행사지급"]
            + df["일비_직인지급"]
            + df["식비_직인지급"]
        )
        df["여행사_지급액"] = (
            df["항공료_여행사지급"] + df["숙박비_여행사지급"]
        )
        df["직원_계좌입금액"] = df["일비_직인지급"] + df["식비_직인지급"]
    return df


processed_df = process_travel_data(st.session_state.travel_list)

# 메인 화면 로직
if not processed_df.empty:
    st.subheader("📊 입력된 출장비 산정 및 지급처별 집계 결과")

    # 데이터 수정/삭제 기능 안내
    st.caption(
        "💡 사이드바의 입력 폼을 통해 출장 건을 계속 추가할 수 있으며, 아래 표에서 내용을 확인할 수 있습니다."
    )
    st.dataframe(processed_df, use_container_width=True)

    # 초기화 버튼
    if st.button("🗑️ 전체 데이터 초기화"):
        st.session_state.travel_list = []
        st.rerun()

    st.markdown("---")

    # 요약 지표
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="총 출장 건수", value=f"{len(processed_df)} 건")
    with col2:
        total_agency = processed_df["여행사_지급액"].sum()
        st.metric(
            label="총 여행사 지급 총액", value=f"{total_agency:,.0f} 원"
        )
    with col3:
        total_employee = processed_df["직원_계좌입금액"].sum()
        st.metric(
            label="총 직원 계좌 입금 총액", value=f"{total_employee:,.0f} 원"
        )

    st.markdown("---")

    # 1. 자금팀 제출용 정산 집계표 다운로드
    st.subheader("📑 자금팀 제출용 정산 집계표 다운로드")
    st.markdown(
        "자금팀 송금 요청을 위한 여행사/직원 지급액 분리 내역서입니다."
    )

    output_agency = io.BytesIO()
    with pd.ExcelWriter(output_agency, engine="openpyxl") as writer:
        processed_df.to_excel(
            writer, index=False, sheet_name="자금팀_정산집계표"
        )
    output_agency.seek(0)

    st.download_button(
        label="📥 자금팀 제출용 엑셀 다운로드",
        data=output_agency,
        file_name="화천기공_자금팀_출장정산집계표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

    # 2. 출장자용 산정 내역서 개별 생성
    st.subheader("👤 출장자 개인별 산정 내역서 조회 및 다운로드")
    selected_person = st.selectbox(
        "출장자를 선택하세요", processed_df["출장자성명"].unique()
    )

    person_data = processed_df[
        processed_df["출장자성명"] == selected_person
    ].iloc[0]

    st.markdown(f"**[ {selected_person} ] 님의 해외출장비 산정 내역**")

    detail_col1, detail_col2 = st.columns(2)
    with detail_col1:
        st.markdown(f"- **소속 부서**: {person_data.get('부서', '-')}")
        st.markdown(f"- **직급**: {person_data.get('직급', '-')}")
        st.markdown(f"- **출장 국가**: {person_data.get('출장국가', '-')}")
        st.markdown(
            f"- **출장 기간**: {person_data.get('출장시작일', '')} ~ {person_data.get('출장종료일', '')}"
        )
    with detail_col2:
        st.markdown(
            f"- **여행사 지급 (항공/숙박)**: {person_data.get('여행사_지급액', 0):,.0f} 원"
        )
        st.markdown(
            f"- **직원 계좌 입금 (일비/식비)**: {person_data.get('직원_계좌입금액', 0):,.0f} 원"
        )
        st.markdown(
            f"- **총 경비 합계**: {person_data.get('총출장비', 0):,.0f} 원"
        )

    output_person = io.BytesIO()
    with pd.ExcelWriter(output_person, engine="openpyxl") as writer:
        pd.DataFrame([person_data]).to_excel(
            writer, index=False, sheet_name="산정내역서"
        )
    output_person.seek(0)

    st.download_button(
        label=f"📥 {selected_person} 산정 내역서 엑셀 다운로드",
        data=output_person,
        file_name=f"화천기공_해외출장산정내역서_{selected_person}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("👋 사이드바 입력 폼을 통해 출장 정보를 추가해 주세요.")
