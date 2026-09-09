import io
import pandas as pd
import streamlit as st

# 페이지 설정
st.set_page_config(
    page_title="화천기공 해외출장비 정산 자동화 프로그램",
    page_layout="wide",
)

st.title("✈️ 화천기공 해외출장비 정산 및 내역서 자동 생성 시스템")
st.markdown(
    "인사지원팀 해외출장 경비 산정, 자금팀 제출용 정산표 분리, 출장자용 산정 내역서 자동 생성 프로그램입니다."
)
st.markdown("---")

# 사이드바 설정
st.sidebar.header("📁 데이터 업로드 및 설정")
uploaded_file = st.sidebar.file_uploader(
    "출장 내역 엑셀 파일을 업로드하세요", type=["xlsx", "xls"]
)

# 기본 기준 정보 설정 (엑셀이 없을 경우 테스트용 또는 기본값 적용)
st.sidebar.subheader("📌 출장비 지급 기준 참고")
st.sidebar.info(
    "- **여행사 지급**: 항공료, 숙박비(대행 시)\n- **직원 계좌 지급**: 일비, 식비, 현지 교통비 등"
)

# 엑셀 파일 템플릿 다운로드 제공 기능 (예시)
with st.sidebar.expander("📥 표준 입력 엑셀 양식 구조 안내"):
    st.markdown(
        """
        엑셀 파일에는 아래 컬럼이 포함되어야 합니다.
        - `출장자성명`, `부서`, `직급`, `계좌번호`
        - `출장국가`, `출장시작일`, `출장종료일`
        - `항공료_여행사지급`, `숙박비_여행사지급`
        - `일비_직인지급`, `식비_직인지급`, `환율`
    """
    )


def process_travel_data(df):
    """출장비 데이터를 기반으로 여행사 지급액과 직원 지급액을 자동 분리 계산"""
    # 계산 로직 구현
    df["총출장비"] = (
        df.get("항공료_여행사지급", 0)
        + df.get("숙박비_여행사지급", 0)
        + df.get("일비_직인지급", 0)
        + df.get("식비_직인지급", 0)
    )

    df["여행사_지급액"] = df.get("항공료_여행사지급", 0) + df.get(
        "숙박비_여행사지급", 0
    )
    df["직원_계좌입금액"] = df.get("일비_직인지급", 0) + df.get(
        "식비_직인지급", 0
    )

    return df


# 메인 화면 로직
if uploaded_file is not None:
    try:
        # 데이터 읽기
        raw_df = pd.read_excel(uploaded_file)
        processed_df = process_travel_data(raw_df)

        st.success("데이터가 성공적으로 로드되고 계산되었습니다.")

        # 1. 전체 집계 결과 표시
        st.subheader("📊 출장비 산정 및 지급처별 집계 결과")
        st.dataframe(processed_df, use_container_width=True)

        # 요약 지표
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(
                label="총 출장 건수", value=f"{len(processed_df)} 건"
            )
        with col2:
            total_agency = processed_df["여행사_지급액"].sum()
            st.metric(
                label="총 여행사 지급 총액",
                value=f"{total_agency:,.0f} 원",
            )
        with col3:
            total_employee = processed_df["직원_계좌입금액"].sum()
            st.metric(
                label="총 직원 계좌 입금 총액",
                value=f"{total_employee:,.0f} 원",
            )

        st.markdown("---")

        # 2. 자금팀 제출용 정산 집계표 다운로드
        st.subheader("📑 자금팀 제출용 정산 집계표")
        st.markdown(
            "자금팀 송금 요청을 위한 여행사/직원 지급액 분리 내역서입니다."
        )

        # 엑셀 다운로드 버퍼 생성
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

        # 3. 출장자용 산정 내역서 개별 생성
        st.subheader("👤 출장자 개인별 산정 내역서 조회 및 다운로드")
        selected_person = st.selectbox(
            "출장자를 선택하세요", processed_df["출장자성명"].unique()
        )

        person_data = processed_df[
            processed_df["출장자성명"] == selected_person
        ].iloc[0]

        st.markdown(
            f"**[ {selected_person} ] 님의 해외출장비 산정 내역**"
        )

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

        # 개인별 내역서 엑셀 다운로드
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

    except Exception as e:
        st.error(
            f"엑셀 파일을 읽고 처리하는 중 오류가 발생했습니다. 파일 양식을 확인해주세요. [오류내용: {e}]"
        )

else:
    st.info(
        "👋 사이드바에서 출장 내역 엑셀 파일을 업로드하면 정산 프로그램이 구동됩니다."
    )

    # 테스트용 가상 데이터 안내
    with st.expander("💡 표준 엑셀 입력 양식 예시 보기"):
        sample_data = pd.DataFrame(
            {
                "출장자성명": ["홍길동", "김철수"],
                "부서": ["인사지원팀", "생산관리팀"],
                "직급": ["대리", "과장"],
                "계좌번호": ["우리은행 123-***", "국민은행 456-***"],
                "출장국가": ["미국", "베트남"],
                "출장시작일": ["2026-09-10", "2026-09-15"],
                "출장종료일": ["2026-09-17", "2026-09-20"],
                "항공료_여행사지급": [2500000, 1200000],
                "숙박비_여행사지급": [1400000, 800000],
                "일비_직인지급": [350000, 250000],
                "식비_직인지급": [490000, 350000],
                "환율": [1350.0, 1350.0],
            }
        )
        st.dataframe(sample_data, use_container_width=True)
