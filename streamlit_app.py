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

# 세션 스테이트 초기화 (출장 데이터 누적 관리)
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
        "출장자 정보를 입력하고 [출장 내역 추가] 버튼을 누르면 데이터가 누적됩니다."
    )

    with st.form("travel_input_form"):
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("👤 출장자 정보")
            name = st.text_input("출장자 성명")
            department = st.text_input("부서", value="인사지원팀")
            position = st.text_input("직급", value="사원")
            account = st.text_input("계좌번호 (은행명 계좌번호)")

        with col_b:
            st.subheader("🌍 출장지 및 일정")
            country = st.text_input("출장 국가 / 도시")
            start_date = st.date_input("출장 시작일")
            end_date = st.date_input("출장 종료일")

        st.markdown("---")
        st.subheader("💵 경비 항목 입력 (원화 기준)")
        col_c, col_d = st.columns(2)
        with col_c:
            st.markdown("##### [ 여행사 대행 지급 ]")
            flight = st.number_input(
                "항공료", min_value=0, value=0, step=10000
            )
            hotel = st.number_input(
                "숙박비", min_value=0, value=0, step=10000
            )
        with col_d:
            st.markdown("##### [ 직원 계좌 직접 지급 ]")
            daily_allowance = st.number_input(
                "일비", min_value=0, value=0, step=5000
            )
            meal = st.number_input(
                "식비", min_value=0, value=0, step=5000
            )

        exchange_rate = st.number_input(
            "적용 환율 (기준일 기준)", min_value=0.0, value=1350.0, step=10.0
        )

        submitted = st.form_submit_button(
            "➕ 입력한 출장 내역 저장 및 추가"
        )

        if submitted:
            if not name or not country:
                st.error(
                    "⚠️ 출장자 성명과 출장 국가는 필수 입력 항목입니다."
                )
            else:
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
                st.success(f"✅ {name} 님의 출장 내역이 추가되었습니다!")

    st.markdown("---")
    st.subheader("📊 현재 등록된 전체 출장 내역 목록")
    current_df = process_travel_data(st.session_state.travel_list)
    if not current_df.empty:
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
        "전체 출장 건에 대해 **[여행사 지급액]**과 **[직원 계좌 입금액]**이 자동으로 분리 계산된 정산표입니다."
    )

    processed_df = process_travel_data(st.session_state.travel_list)

    if not processed_df.empty:
        # 요약 지표 카드
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric(
                label="총 출장 건수", value=f"{len(processed_df)} 건"
            )
        with m2:
            total_agency = processed_df["여행사_지급액"].sum()
            st.metric(
                label="총 여행사 지급 총액",
                value=f"{total_agency:,.0f} 원",
            )
        with m3:
            total_employee = processed_df["직원_계좌입금액"].sum()
            st.metric(
                label="총 직원 계좌 입금 총액",
                value=f"{total_employee:,.0f} 원",
            )

        st.markdown("---")
        st.subheader("📑 자금팀 송금 요청 집계 미리보기")
        st.dataframe(processed_df, use_container_width=True)

        # 엑셀 다운로드 처리
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
        "출장자 개인별로 상세 산정 내역을 확인하고 안내용 엑셀 파일을 생성하여 전달할 수 있습니다."
    )

    processed_df = process_travel_data(st.session_state.travel_list)

    if not processed_df.empty:
        selected_person = st.selectbox(
            "내역서를 생성할 출장자를 선택하세요",
            processed_df["출장자성명"].unique(),
        )

        person_data = processed_df[
            processed_df["출장자성명"] == selected_person
        ].iloc[0]

        st.markdown("---")
        st.markdown(
            f"### 👤 [ {selected_person} ] 님 해외출장비 산정 내역서"[cite: 1]
        )

        det_c1, det_c2 = st.columns(2)
        with det_c1:
            st.info(
                f"""
            - **소속 부서**: {person_data.get('부서', '-')}
            - **직급**: {person_data.get('직급', '-')}
            - **출장 국가**: {person_data.get('출장국가', '-')}
            - **출장 계좌**: {person_data.get('계좌번호', '-')}
            """
            )
        with det_c2:
            st.success(
                f"""
            - **출장 기간**: {person_data.get('출장시작일', '')} ~ {person_data.get('출장종료일', '')}
            - **적용 환율**: {person_data.get('환율', 0):,.2f} 원
            - **총 출장 경비**: {person_data.get('총출장비', 0):,.0f} 원
            """
            )

        # 상세 항목 표
        detail_table = pd.DataFrame(
            {
                "구분": [
                    "항공료 (여행사 지급)",
                    "숙박비 (여행사 지급)",
                    "일비 (직원 지급)",
                    "식비 (직원 지급)",
                ],
                "금액 (원)": [
                    f"{person_data.get('항공료_여행사지급', 0):,.0f}",
                    f"{person_data.get('숙박비_여행사지급', 0):,.0f}",
                    f"{person_data.get('일비_직인지급', 0):,.0f}",
                    f"{person_data.get('식비_직인지급', 0):,.0f}",
                ],
                "지급처": [
                    "여행사",
                    "여행사",
                    "직원 계좌",
                    "직원 계좌",
                ],
            }
        )
        st.table(detail_table)

        # 개인별 내역서 엑셀 다운로드
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
