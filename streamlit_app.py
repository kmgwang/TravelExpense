import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="해외 출장비 정산 시스템", page_icon="✈️", layout="wide"
)

st.title("✈️ 해외 출장비 정산 및 환율 자동 계산 프로그램")
st.markdown("---")

# 1. 기본 정보 입력 섹션
st.subheader("1. 출장 기본 정보 및 환율 설정")
col1, col2, col3 = st.columns(3)

with col1:
    traveler_name = st.text_input("출장자 성명", "홍길동")
with col2:
    destination = st.selectbox(
        "출장 국가", ["일본 (JPY)", "미국 (USD)", "유럽 (EUR)", "기타"]
    )
with col3:
    if "일본" in destination:
        currency = "JPY"
    elif "미국" in destination:
        currency = "USD"
    elif "유럽" in destination:
        currency = "EUR"
    else:
        currency = "USD"

col_r1, col_r2 = st.columns(2)
with col_r1:
    # 환율 입력 (일본의 경우 고시 환율이 100엔 기준인 점을 고려)
    if currency == "JPY":
        raw_exchange_rate = st.number_input(
            "적용 환율 (엔화 100엔 기준 원화 환산액)",
            min_value=0.0,
            value=950.0,
            step=1.0,
        )
        # 핵심 요구사항: 일본 환율인 경우 환율 / 100 적용
        effective_exchange_rate = raw_exchange_rate / 100.0
        st.info(
            f"ℹ️ 일본(JPY) 적용: 입력하신 환율({raw_exchange_rate:,.2f})에 따라 1엔당 실적용 환율은 **{effective_exchange_rate:,.4f}원**으로 계산됩니다."
        )
    else:
        effective_exchange_rate = st.number_input(
            f"적용 환율 (1 {currency}당 원화)",
            min_value=0.0,
            value=1350.0,
            step=1.0,
        )

with col_r2:
    days = st.number_input("출장 일수 (일)", min_value=1, value=5, step=1)

st.markdown("---")

# 2. 지출 비용 입력 섹션 (현지 통화 기준)
st.subheader("2. 지출 내역 입력 (현지 통화)")

col_e1, col_e2, col_e3 = st.columns(3)
with col_e1:
    daily_allowance_local = st.number_input(
        f"일일 체제비 (1일당 {currency})", min_value=0.0, value=10000.0, step=100.0
    )
with col_e2:
    lodging_local = st.number_input(
        f"숙박비 총액 ({currency})", min_value=0.0, value=80000.0, step=1000.0
    )
with col_e3:
    transport_local = st.number_input(
        f"기타 교통비/부대비 ({currency})",
        min_value=0.0,
        value=15000.0,
        step=500.0,
    )

# 3. 환율 및 계산식 적용 (환율이 들어가는 모든 연산에 effective_exchange_rate 일괄 반영)
total_daily_allowance_local = daily_allowance_local * days
total_expense_local = (
    total_daily_allowance_local + lodging_local + transport_local
)

# 원화 환산액 산출
total_daily_allowance_krw = total_daily_allowance_local * effective_exchange_rate
lodging_krw = lodging_local * effective_exchange_rate
transport_krw = transport_local * effective_exchange_rate
total_expense_krw = total_expense_local * effective_exchange_rate

st.markdown("---")

# 4. 정산 결과 출력 섹션
st.subheader("3. 최종 정산 내역서")

summary_data = {
    "항목": ["일일 체제비", "숙박비", "기타 교통비/부대비", "총 합계"],
    f"현지 통화 ({currency})": [
        total_daily_allowance_local,
        lodging_local,
        transport_local,
        total_expense_local,
    ],
    "적용 환율": [
        effective_exchange_rate,
        effective_exchange_rate,
        effective_exchange_rate,
        effective_exchange_rate,
    ],
    "원화 환산액 (KRW)": [
        total_daily_allowance_krw,
        lodging_krw,
        transport_krw,
        total_expense_krw,
    ],
}

df_summary = pd.DataFrame(summary_data)

st.dataframe(
    df_summary.style.format(
        {
            f"현지 통화 ({currency})": "{:,.2f}",
            "적용 환율": "{:,.4f}",
            "원화 환산액 (KRW)": "{:,.0f}",
        }
    ),
    use_container_width=True,
)

st.success(
    f"💡 **총 청구 원화 금액**: **₩{total_expense_krw:,.0f}** (적용 통화: {currency})"
)
