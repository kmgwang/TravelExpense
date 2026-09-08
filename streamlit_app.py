import streamlit as st
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

st.set_page_config(
    page_title="화천기공 업무 지원 시스템",
    page_icon="⚙️",
    layout="wide"
)

# 폰트 등록 (malgun.ttf가 같은 경로에 있다고 가정)
font_path = "malgun.ttf"
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('Malgun', font_path))
    DEFAULT_FONT = 'Malgun'
else:
    DEFAULT_FONT = 'Helvetica'

st.title("⚙️ 화천기공 인사·총무 업무 지원 시스템")
st.markdown("---")

# 세션 상태 초기화
if 'data_log' not in st.session_state:
    st.session_state.data_log = []

tab1, tab2, tab3 = st.tabs(["📋 데이터 관리 및 입력", "📊 대시보드 및 리포트", "🖨️ PDF/출력물 생성"])

with tab1:
    st.subheader("신규 데이터 등록")
    with st.form("input_form"):
        col1, col2 = st.columns(2)
        with col1:
            emp_name = st.text_input("성명")
            dept = st.selectbox("부서", ["인사총무팀", "생산관리팀", "기술지원팀", "경영기획팀"])
        with col2:
            task_type = st.selectbox("업무 구분", ["사원증 관리", "급여 및 제수당", "출장비 정산", "기타 행정"])
            amount = st.number_input("관련 금액 / 수량", min_value=0, value=0, step=1000)
        
        submitted = st.form_submit_button("데이터 추가")
        if submitted:
            if emp_name:
                new_row = {
                    "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "성명": emp_name,
                    "부서": dept,
                    "업무구분": task_type,
                    "금액": amount
                }
                st.session_state.data_log.append(new_row)
                st.success(f"'{emp_name}'님의 데이터가 성공적으로 등록되었습니다.")
            else:
                st.warning("성명을 입력해주세요.")

    st.markdown("### 등록된 데이터 목록")
    if st.session_state.data_log:
        df = pd.DataFrame(st.session_state.data_log)
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    else:
        info_df = pd.DataFrame(columns=["등록일시", "성명", "부서", "업무구분", "금액"])
        st.data_editor(info_df, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("업무 현황 대시보드")
    if st.session_state.data_log:
        df_dash = pd.DataFrame(st.session_state.data_log)
        col1, col2, col3 = st.columns(3)
        col1.metric("총 등록 건수", f"{len(df_dash)} 건")
        col2.metric("총 금액 합계", f"{df_dash['금액'].sum():,0f} 원")
        col3.metric("참여 부서 수", f"{df_dash['부서'].nunique()} 개")
        
        st.markdown("#### 부서별 처리 현황")
        dept_summary = df_dash.groupby("부서")["금액"].sum().reset_index()
        st.bar_chart(dept_summary.set_index("부서"))
    else:
        st.info("시각화할 데이터가 없습니다. '데이터 관리 및 입력' 탭에서 데이터를 추가해주세요.")

with tab3:
    st.subheader("공식 문서 및 PDF 출력")
    st.markdown("입력된 데이터를 바탕으로 보고서용 PDF 파일을 생성합니다.")
    
    doc_title = st.text_input("보고서 제목", "화천기공 업무 수행 보고서")
    
    if st.button("PDF 보고서 생성"):
        pdf_filename = "hwacheon_report.pdf"
        c = canvas.Canvas(pdf_filename, pagesize=A4)
        width, height = A4
        
        # 상단 헤더
        c.setFont(DEFAULT_FONT, 18)
        c.drawString(50, height - 50, doc_title)
        
        c.setFont(DEFAULT_FONT, 10)
        c.drawString(50, height - 70, f"출력 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        c.drawString(50, height - 85, "발신: 화천기공 주식회사 인사총무팀")
        c.line(50, height - 95, width - 50, height - 95)
        
        # 내용 작성
        y_pos = height - 130
        c.setFont(DEFAULT_FONT, 12)
        c.drawString(50, y_pos, "[ 세부 내역 ]")
        y_pos -= 25
        
        if st.session_state.data_log:
            for item in st.session_state.data_log:
                text = f"- [{item['부서']}] {item['성명']} ({item['업무구분']}): {item['금액']:,}원 (등록: {item['등록일시']})"
                c.setFont(DEFAULT_FONT, 10)
                c.drawString(55, y_pos, text)
                y_pos -= 20
                if y_pos < 50:
                    c.showPage()
                    y_pos = height - 50
        else:
            c.setFont(DEFAULT_FONT, 10)
            c.drawString(55, y_pos, "등록된 내역이 없습니다.")
            
        c.save()
        
        with open(pdf_filename, "rb") as f:
            pdf_bytes = f.read()
            
        st.download_button(
            label="📄 PDF 보고서 다운로드",
            data=pdf_bytes,
            file_name="화천기공_업무보고서.pdf",
            mime="application/pdf"
        )
        st.success("PDF 파일 생성이 완료되었습니다!")
