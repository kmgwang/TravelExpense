import io
import os
import tempfile
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Windows COM 객체 제어를 통한 엑셀 영역 이미지 변환 모듈 (pywin32)
try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

st.set_page_config(page_title="자금팀 정산 집계표 생성기", layout="wide")

st.title("💼 자금팀 정산 집계표 자동화 프로그램")

# 예시 데이터프레임 로드 (실무 데이터 연동부)
@st.cache_data
def load_sample_data():
    data = {
        "출장자": ["김철수", "박영희", "이민수", "정지현"],
        "부서": ["인사지원팀", "재경팀", "생산관리팀", "품질보증팀"],
        "출장목적": ["법인카탈로그 검수", "세무조사 대응", "설비 점검", "품질 협의"],
        "출장지": ["서울", "부산", "울산", "대구"],
        "출장기간": ["2026-03-01 ~ 2026-03-03", "2026-03-05 ~ 2026-03-06", "2026-03-10 ~ 2026-03-12", "2026-03-15 ~ 2026-03-15"],
        "교통비": [150000, 250000, 180000, 120000],
        "숙박비": [200000, 300000, 240000, 0],
        "일당": [120000, 80000, 120000, 40000]
    }
    return pd.DataFrame(data)

df = load_sample_data()

def create_excel_file(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='정산집계표')
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    # 스타일 정의 (1행 하단 굵은 실선 적용)
    header_border = Border(
        bottom=Side(style='medium', color='000000'),
        top=Side(style='thin', color='000000'),
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3')
    )

    # 1행 헤더 서식 지정
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(name="맑은 고딕", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        cell.border = header_border

    # 데이터 셀 테두리 및 정렬 지정
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = Font(name="맑은 고딕", size=10)
            cell.border = Border(
                bottom=Side(style='thin', color='E0E0E0'),
                top=Side(style='thin', color='E0E0E0'),
                left=Side(style='thin', color='E0E0E0'),
                right=Side(style='thin', color='E0E0E0')
            )

    # 열 너비 자동 조정
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

def convert_excel_range_to_image(excel_bytes):
    if not WIN32_AVAILABLE:
        return None

    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, "temp_excel.xlsx")
        with open(excel_path, "wb") as f:
            f.write(excel_bytes.getbuffer())

        image_path = os.path.join(tmpdir, "settlement_image.png")

        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            ws = wb.Sheets(1)

            # H열 기준 데이터가 입력된 마지막 행 탐색 (-4162 = xlUp)
            last_row = ws.Cells(ws.Rows.Count, "H").End(-4162).Row
            if last_row < 1:
                last_row = ws.UsedRange.Rows.Count

            # A1부터 H열 마지막 행까지 범위 지정 및 그림 복사
            range_address = f"A1:H{last_row}"
            rng = ws.Range(range_address)
            rng.CopyPicture(Appearance=1, Format=2) # xlScreen, xlBitmap

            # 차트 객체를 활용하여 이미지 내보내기
            chart_page = ws.ChartObjects().Add(Left=rng.Left, Top=rng.Top, Width=rng.Width, Height=rng.Height)
            chart_page.Chart.Paste()
            chart_page.Chart.Export(os.path.abspath(image_path), "PNG")
            chart_page.Delete()

            wb.Close(SaveChanges=False)
            excel.Quit()

            with open(image_path, "rb") as img_f:
                return img_f.read()

        except Exception as e:
            excel.Quit()
            st.error(f"이미지 변환 중 오류 발생: {e}")
            return None

# 파일 생성 실행
excel_data = create_excel_file(df)

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.download_button(
        label="📥 자금팀 정산 집계표 엑셀 다운로드",
        data=excel_data,
        file_name="자금팀_정산집계표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col2:
    if st.button("🖼️ 자금팀 정산 집계표 이미지 다운로드"):
        if not WIN32_AVAILABLE:
            st.error("Windows 환경 및 pywin32 라이브러리가 필요합니다. (터미널에 `pip install pywin32` 실행 필요)")
        else:
            with st.spinner("엑셀 영역을 이미지로 변환 중입니다..."):
                excel_data.seek(0)
                img_bytes = convert_excel_range_to_image(excel_data)
                if img_bytes:
                    st.image(img_bytes, caption="자금팀 정산 집계표 (A1:H열 마지막행)", use_container_width=True)
                    st.download_button(
                        label="💾 이미지 파일(PNG) 저장",
                        data=img_bytes,
                        file_name="자금팀_정산집계표.png",
                        mime="image/png"
                    )
