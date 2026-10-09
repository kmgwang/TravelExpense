import base64
import datetime
import io
import os
import platform
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# ----------------------------------------------------
# 1. 폰트 및 페이지 설정
# ----------------------------------------------------
if platform.system() == "Windows":
    matplotlib.rc("font", family="Malgun Gothic")
elif platform.system() == "Darwin":
    matplotlib.rc("font", family="AppleGothic")
else:
    matplotlib.rcParams["font.family"] = "NanumGothic"
matplotlib.rcParams["axes.unicode_minus"] = False

try:
    st.set_page_config(
        page_title="HWACHEON - 해외출장비 정산 및 기록 관리 시스템",
        page_icon="✈️",
        layout="wide",
        initial_sidebar_state="expanded",
    )
except Exception:
    pass

# ----------------------------------------------------
# 2. 화천(HWACHEON) 홈페이지 스타일 커스텀 CSS
#    브랜드 컬러: PANTONE 2728C / C96 M69 Y0 K0 / R0 G92 B171 (#005CAB)
# ----------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

    :root {
        --hw-blue: #005CAB;
        --hw-blue-dark: #004A8C;
        --hw-blue-soft: #EAF2FA;
        --hw-gray-bg: #F5F6F8;
        --hw-line: #E5E7EB;
        --hw-text: #1F2328;
        --hw-text-sub: #6B7280;
    }

    html, body, .stApp, .stApp p, .stApp label, .stApp input, .stApp textarea,
    .stApp button, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5,
    .stApp td, .stApp th, .stApp li, .stApp a, .stApp div[data-baseweb] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto,
                     'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
    }

    .stApp { background-color: #ffffff; color: var(--hw-text); }
    header[data-testid="stHeader"] { background: transparent; }
    .block-container {
        max-width: 1200px !important;
        padding-top: 2.5rem;
        padding-bottom: 4rem;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
    }

    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid var(--hw-line);
        min-width: 290px;
        max-width: 290px;
        transform: none !important;
        margin-left: 0 !important;
        visibility: visible !important;
        width: 290px !important;
    }
    section[data-testid="stSidebar"] > div:first-child { padding-top: 0.5rem; }

    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarResizeHandle"],
    section[data-testid="stSidebar"] button[kind="header"],
    section[data-testid="stSidebar"] button[kind="headerNoPadding"] {
        display: none !important;
    }

    .hw-side-brand {
        padding: 18px 8px 22px 8px;
        border-bottom: 1px solid var(--hw-line);
        margin-bottom: 22px;
    }
    .hw-side-logo {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hw-side-logo svg { height: 30px; width: auto; flex-shrink: 0; }
    .hw-side-logo-text {
        font-size: 22px;
        font-weight: 800;
        letter-spacing: 1.5px;
        color: var(--hw-blue);
        line-height: 1;
    }
    .hw-side-desc {
        font-size: 12.5px;
        color: var(--hw-text-sub);
        line-height: 1.6;
        margin-top: 14px;
    }
    .hw-side-caption {
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 1px;
        color: var(--hw-text-sub);
        padding: 0 8px 10px 8px;
    }

    section[data-testid="stSidebar"] div[role="radiogroup"] {
        gap: 6px;
        width: 100%;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label {
        width: 100%;
        padding: 14px 18px;
        border-radius: 12px;
        background-color: transparent;
        cursor: pointer;
        transition: background-color 0.15s ease;
        margin: 0;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
        display: none;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label p {
        font-size: 15px;
        font-weight: 600;
        color: var(--hw-text);
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
        background-color: var(--hw-gray-bg);
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
        background-color: var(--hw-blue);
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p {
        color: #ffffff;
    }

    .hw-page-head {
        padding-bottom: 18px;
        border-bottom: 1px solid var(--hw-line);
        margin-bottom: 28px;
    }
    .hw-page-eyebrow {
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 1.2px;
        color: var(--hw-blue);
        margin-bottom: 6px;
    }
    .hw-page-title {
        font-size: 30px;
        font-weight: 800;
        color: var(--hw-text);
        letter-spacing: -0.5px;
    }
    .hw-page-sub {
        color: var(--hw-text-sub);
        font-size: 14px;
        margin-top: 6px;
    }

    .stApp h3 {
        font-weight: 800;
        letter-spacing: -0.3px;
        color: var(--hw-text);
    }

    .hw-card {
        background: var(--hw-gray-bg);
        padding: 24px;
        border-radius: 16px;
        margin-bottom: 20px;
    }

    .row-highlight-blue {
        background-color: var(--hw-blue-soft);
        border: none;
        color: var(--hw-blue);
        padding: 12px 16px;
        border-radius: 10px;
        font-size: 16px;
    }

    div[data-baseweb="input"],
    div[data-baseweb="base-input"],
    div[data-baseweb="select"] > div,
    div[data-testid="stDateInput"] div[data-baseweb="input"] {
        background-color: var(--hw-gray-bg) !important;
        border: 1px solid transparent !important;
        border-radius: 10px !important;
    }
    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="select"] > div:focus-within {
        border: 1px solid var(--hw-blue) !important;
        background-color: #ffffff !important;
    }
    div[data-baseweb="input"] input { background-color: transparent !important; }

    .stButton > button {
        background-color: var(--hw-blue);
        color: #ffffff;
        border-radius: 10px;
        border: none;
        font-weight: 600;
        padding: 10px 22px;
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        background-color: var(--hw-blue-dark);
        color: #ffffff;
    }

    div[data-testid="stMetric"] {
        background-color: var(--hw-gray-bg);
        border-radius: 16px;
        padding: 18px 22px;
    }
    div[data-testid="stMetricLabel"] p { color: var(--hw-text-sub); font-weight: 600; }
    div[data-testid="stMetricValue"] { color: var(--hw-blue); font-weight: 800; }

    .stApp hr { border-color: var(--hw-line); }
    </style>
""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 3. 사이드바 (로고 + 메뉴 내비게이션)
# ----------------------------------------------------
MENU_1 = "1. 출장 정보 입력"
MENU_2 = "2. 자금팀 연결 자료 생성"
MENU_3 = "3. 출장비 산정 내역서 생성"
MENU_4 = "4. 해외출장 기록 관리"

LOGO_SVG = (
    '<svg viewBox="0 53 906 264" xmlns="http://www.w3.org/2000/svg" aria-label="HWACHEON">'
    '<rect x="0" y="53" width="88" height="264" fill="#005CAB"/>'
    '<rect x="193" y="53" width="519" height="88" fill="#005CAB"/>'
    '<rect x="193" y="229" width="519" height="88" fill="#005CAB"/>'
    '<rect x="818" y="53" width="88" height="264" fill="#005CAB"/>'
    "</svg>"
)

SIDEBAR_BRAND_HTML = (
    '<div class="hw-side-brand">'
    '<div class="hw-side-logo">'
    + LOGO_SVG
    + '<div class="hw-side-logo-text">HWACHEON</div>'
    "</div>"
    '<div class="hw-side-desc">화천기공 해외출장비 산정 및 기록 관리 시스템</div>'
    "</div>"
    '<div class="hw-side-caption">MENU</div>'
)

with st.sidebar:
    st.markdown(SIDEBAR_BRAND_HTML, unsafe_allow_html=True)
    menu = st.radio(
        "메뉴",
        [MENU_1, MENU_2, MENU_3, MENU_4],
        key="nav_menu",
        label_visibility="collapsed",
    )

PAGE_HEAD = {
    MENU_1: ("STEP 01", "출장 정보 입력", "출장자 정보와 출장 목적 및 경비 항목을 입력하고 등록합니다."),
    MENU_2: ("STEP 02", "자금팀 연결 자료 생성", "등록된 출장 내역을 자금팀 제출용으로 집계합니다."),
    MENU_3: ("STEP 03", "출장비 산정 내역서 생성", "출장자별 해외출장비 산정 내역서를 확인합니다."),
    MENU_4: ("STEP 04", "해외출장 기록 관리", "완료된 출장 기록을 통합 보관하고 조건별로 다차원 검색합니다."),
}
_eyebrow, _title, _sub = PAGE_HEAD[menu]
st.markdown(
    f"""
    <div class="hw-page-head">
        <div class="hw-page-eyebrow">{_eyebrow}</div>
        <div class="hw-page-title">{_title}</div>
        <div class="hw-page-sub">{_sub}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 세션 상태 초기화
if "travel_list" not in st.session_state:
    st.session_state.travel_list = []

if "completed_travel_list" not in st.session_state:
    # 초기 예시 데이터 생성 (검색 및 관리가 즉각 확인 가능하도록 설정)
    st.session_state.completed_travel_list = [
        {
            "출장자성명": "김화천",
            "부서": "생산기술팀",
            "직급": "차장",
            "직급구분": "1급",
            "출장지": "일본 (오사카)",
            "지역구분": "특",
            "출장목적": "JIMTOF 2026 공작기계 전시회 참관 및 기술 협의",
            "출장시작일": "2026-03-10",
            "출장종료일": "2026-03-15",
            "출장일수": 6,
            "출장박수": 5,
            "환율": 915.20,
            "직원_지급액": 1250000,
            "여행사_지급액": 1800000,
            "총출장비": 3050000,
            "교통비항목리스트": [{"item": "항공권", "amount": 1800000, "payer": "여행사"}],
            "출장비항목리스트": [{"item": "숙박비", "amount": 800000, "payer": "출장자"}, {"item": "일당", "amount": 450000, "payer": "출장자"}],
            "기타항목리스트": [],
            "완료일자": "2026-03-16"
        },
        {
            "출장자성명": "이경영",
            "부서": "인사지원팀",
            "직급": "과장",
            "직급구분": "2급",
            "출장지": "미국 (카슨)",
            "지역구분": "갑",
            "출장목적": "미국 법인(HWACHEON Machinery America) 인사제도 점검",
            "출장시작일": "2026-04-01",
            "출장종료일": "2026-04-08",
            "출장일수": 8,
            "출장박수": 7,
            "환율": 1345.00,
            "직원_지급액": 2100000,
            "여행사_지급액": 3200000,
            "총출장비": 5300000,
            "교통비항목리스트": [{"item": "항공권", "amount": 3200000, "payer": "여행사"}],
            "출장비항목리스트": [{"item": "숙박비", "amount": 1400000, "payer": "출장자"}, {"item": "일당", "amount": 700000, "payer": "출장자"}],
            "기타항목리스트": [],
            "완료일자": "2026-04-09"
        }
    ]

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
            "베트남", "태국", "말레이시아", "인도네시아", "필리핀",
            "싱가포르", "인도", "파키스탄", "방글라데시", "카자흐스탄", "우즈베키스탄",
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
        applied_rate = raw_rate / 100.0 if region == "특" else raw_rate

        calc_daily = int((std_daily * applied_rate * d.get("출장일수", 1)) // 1000 * 1000)

        if std_hotel == "실비":
            calc_hotel = 0
        else:
            calc_hotel = int((std_hotel * applied_rate * d.get("출장박수", 0)) // 1000 * 1000)

        agency_total = 0
        employee_total = 0

        for te in d.get("출장비항목리스트", []):
            amt = te.get("amount", 0)
            if te.get("payer", "출장자") == "여행사":
                agency_total += amt
            else:
                employee_total += amt

        for t in d.get("교통비항목리스트", []):
            amt = t.get("amount", 0)
            if t.get("payer", "여행사") == "여행사":
                agency_total += amt
            else:
                employee_total += amt

        for other in d.get("기타항목리스트", []):
            o_amt = other.get("amount", 0)
            if other.get("payer", "여행사") == "여행사":
                agency_total += o_amt
            else:
                employee_total += o_amt

        d["직원_지급액"] = employee_total
        d["여행사_지급액"] = agency_total
        d["총출장비"] = employee_total + agency_total

        ordered_d = {
            "출장자성명": d.get("출장자성명"),
            "부서": d.get("부서"),
            "직급": d.get("직급"),
            "직급구분": d.get("직급구분"),
            "출장지": d.get("출장지"),
            "지역구분": d.get("지역구분"),
            "출장목적": d.get("출장목적", "-"),
            "출장시작일": d.get("출장시작일"),
            "출장종료일": d.get("출장종료일"),
            "출장박수": d.get("출장박수"),
            "출장일수": d.get("출장일수"),
            "환율": d.get("환율"),
            "직원_지급액": d.get("직원_지급액"),
            "여행사_지급액": d.get("여행사_지급액"),
            "총출장비": d.get("총출장비"),
            "교통비항목리스트": d.get("교통비항목리스트"),
            "출장비항목리스트": d.get("출장비항목리스트"),
            "기타항목리스트": d.get("기타항목리스트"),
        }
        if "완료일자" in d:
            ordered_d["완료일자"] = d["완료일자"]

        processed.append(ordered_d)

    return pd.DataFrame(processed)

# ----------------------------------------------------
# 엑셀 생성 함수
# ----------------------------------------------------
_THIN = Side(style="thin", color="BFBFBF")
XL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
XL_HEADER_FILL = PatternFill("solid", fgColor="005CAB")
XL_SUB_FILL = PatternFill("solid", fgColor="EAF2FA")
XL_HEADER_FONT = Font(bold=True, color="FFFFFF")
XL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
XL_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
XL_RIGHT = Alignment(horizontal="right", vertical="center")
XL_NUM = "#,##0"
XL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def payer_display(payer, rec):
    if payer == "출장자":
        return (rec.get("출장자성명") or "").strip() or "출장자"
    return payer

def calc_item_rows(record):
    region = record["지역구분"]
    std_daily, std_hotel = get_standard_rates(region, record["직급구분"])
    cur = "JPY" if region == "특" else "USD"

    rows = []
    for t in record.get("교통비항목리스트", []):
        rows.append(("교통비", t.get("item", ""), int(t.get("amount", 0)), payer_display(t.get("payer", "여행사"), record)))
    for te in record.get("출장비항목리스트", []):
        rows.append(("출장비", te.get("item", ""), int(te.get("amount", 0)), payer_display(te.get("payer", "출장자"), record)))
    for o in record.get("기타항목리스트", []):
        if o.get("item") or o.get("amount"):
            rows.append(("기타", o.get("item", ""), int(o.get("amount", 0)), payer_display(o.get("payer", "여행사"), record)))

    return rows, std_daily, std_hotel, cur

def _xl_header(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.fill = XL_HEADER_FILL
        c.font = XL_HEADER_FONT
        c.alignment = XL_CENTER
        c.border = XL_BORDER

def _xl_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

def _xl_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def build_fund_excel(records):
    df = process_travel_data(records)
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "자금팀 정산 집계표"
    ws["A1"] = "해외출장비 자금팀 정산 집계표"
    ws["A1"].font = Font(bold=True, size=16, color="005CAB")
    ws.merge_cells("A1:I1")
    ws["A2"] = f"작성일: {datetime.date.today():%Y-%m-%d}"
    ws["A2"].alignment = XL_RIGHT
    ws.merge_cells("A2:I2")

    headers = ["순번", "출장자", "부서", "직급", "출장지", "출장 기간", "직원 지급액(원)", "여행사 지급액(원)", "총 출장비(원)"]
    hr = 4
    _xl_header(ws, hr, headers)
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        row = hr + i
        vals = [
            i, r["출장자성명"], r["부서"], r["직급"], r["출장지"],
            f'{r["출장시작일"]} ~ {r["출장종료일"]}',
            int(r["직원_지급액"]), int(r["여행사_지급액"]), int(r["총출장비"]),
        ]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=j, value=v)
            c.border = XL_BORDER
            if j >= 7:
                c.number_format = XL_NUM
                c.alignment = XL_RIGHT
            else:
                c.alignment = XL_CENTER

    first, last = hr + 1, hr + len(df)
    tot = last + 1
    ws.cell(row=tot, column=1, value="합 계")
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=6)
    for j in range(1, 10):
        c = ws.cell(row=tot, column=j)
        c.fill = XL_SUB_FILL
        c.font = Font(bold=True)
        c.border = XL_BORDER
        c.alignment = XL_CENTER if j < 7 else XL_RIGHT
    for j in (7, 8, 9):
        col = openpyxl.utils.get_column_letter(j)
        c = ws.cell(row=tot, column=j, value=f"=SUM({col}{first}:{col}{last})")
        c.number_format = XL_NUM
    _xl_widths(ws, [7, 12, 16, 10, 14, 24, 22, 20, 18])
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)

    ws2 = wb.create_sheet("항목별 상세")
    _xl_header(ws2, 1, ["순번", "출장자", "출장지", "구분", "항목", "금액(원)", "지급처"])
    r2 = 2
    for i, rec in enumerate(records, start=1):
        items, _, _, _ = calc_item_rows(rec)
        for cat, name, amt, payer in items:
            vals = [i, rec["출장자성명"], rec["출장지"], cat, name, amt, payer]
            for j, v in enumerate(vals, start=1):
                c = ws2.cell(row=r2, column=j, value=v)
                c.border = XL_BORDER
                c.alignment = XL_RIGHT if j == 6 else XL_CENTER
                if j == 6:
                    c.number_format = XL_NUM
            r2 += 1
    _xl_widths(ws2, [7, 12, 14, 10, 24, 16, 12])
    ws2.freeze_panes = "A2"

    return _xl_bytes(wb)

TEMPLATE_FILENAME = "출장비산정내역_양식.xlsx"
TEMPLATE_B64 = (
    "UEsDBBQABgAIAAAAIQBBN4LPbgEAAAQFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsVMluwjAQvVfqP0S+Vomhh6qqCBy6HFsk6AeYeJJY"
    "JLblGSj8fSdmUVWxCMElUWzPWybzPBit2iZZQkDjbC76WU8kYAunja1y8T39SJ9FgqSsVo2zkIs1oBgN7+8G07UHTLjaYi5qIv8i"
    "JRY1tAoz58HyTulCq4g/QyW9KuaqAvnY6z3JwlkCSyl1GGI4eINSLRpK3le8vFEyM1Ykr5tzHVUulPeNKRSxULm0+h9J6srSFKBd"
    "sWgZOkMfQGmsAahtMh8MM4YJELExFPIgZ4AGLyPdusq4MgrD2nh8YOtHGLqd4662dV/8O4LRkIxVoE/Vsne5auSPC/OZc/PsNMil"
    "rYktylpl7E73Cf54GGV89W8spPMXgc/oIJ4xkPF5vYQIc4YQad0A3rrtEfQcc60C6Anx9FY3F/AX+5QOjtQ4OI+c2gCXd2EXka46"
    "9QwEgQzsQ3Jo2PaMHPmr2w7dnaJBH+CW8Q4b/gIAAP//AwBQSwMEFAAGAAgAAAAhALVVMCP0AAAATAIAAAsACAJfcmVscy8ucmVs"
    "cyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1P"
    "wzAMhu9I/IfI99XdkBBCS3dBSLshVH6ASdwPtY2jJBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZy"
    "lkZxrOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYWPZmBWsZNWd5i+K4B1UJT7a2GsLc3oOqTz5t/15am"
    "6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80SbvxP9fC1OnMhSIjQS+DLPR8cloPV/WrQ08cudecQ3CcOr"
    "yPDJgosfqN4BAAD//wMAUEsDBBQABgAIAAAAIQCBPpSX8wAAALoCAAAaAAgBeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHMgogQB"
    "KKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsUk1LxDAQvQv+hzB3m3YVEdl0LyLsVesPCMm0KdsmITN+9N8bKrpdWNZLLwNvhnnvzcd2"
    "9zUO4gMT9cErqIoSBHoTbO87BW/N880DCGLtrR6CRwUTEuzq66vtCw6acxO5PpLILJ4UOOb4KCUZh6OmIkT0udKGNGrOMHUyanPQ"
    "HcpNWd7LtOSA+oRT7K2CtLe3IJopZuX/uUPb9gafgnkf0fMZCUk8DXkA0ejUISv4wUX2CPK8/GZNec5rwaP6DOUcq0seqjU9fIZ0"
    "IIfIRx9/KZJz5aKZu1Xv4XRC+8opv9vyLMv072bkycfV3wAAAP//AwBQSwMEFAAGAAgAAAAhANCQ5fOnAgAA+AUAAA8AAAB4bC93"
    "b3JrYm9vay54bWykVMtq20AU3Rf6D8PsFWn8kB0ROSRWQg1tMW2TbAxhLI2twZJGnRnFDiGb0F8o3ZTSTZeFLkqh3xR/RO9IlhPH"
    "mzQR9p3HFWfOufdo9vYXaYIumFRcZD4mOw5GLAtFxLOpj08+HFtdjJSmWUQTkTEfXzKF93svX+zNhZyNhZghAMiUj2Otc8+2VRiz"
    "lKodkbMMMhMhU6phKae2yiWjkYoZ02liNxzHtVPKM1whePIxGGIy4SELRFikLNMViGQJ1UBfxTxXNVoaPgYuAnx9FY3F/AX+5QOj"
    "tQ4OI+c2gCXd2EXka469QwEgQzsQ3Jo2PaMHPmr2w7dnaJBH+CW8Q4b/gIAAP//AwBQSwMEFAAGAAgAAAAhALVVMCP0AAAATAIA"
    "AAsACAJfcmVscy8ucmVs"
)

def load_statement_template():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATE_FILENAME)
    if os.path.exists(path):
        return openpyxl.load_workbook(path)
    return openpyxl.load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)))

def fill_statement_sheet(ws, rec):
    items, std_daily, std_hotel, cur = calc_item_rows(rec)
    jpy = rec["지역구분"] == "특"
    days = rec["출장일수"]
    nights = rec["출장박수"]
    hotel_amt = sum(a for cat, n, a, p in items if cat == "출장비" and "숙박" in n)
    daily_amt = sum(a for cat, n, a, p in items if cat == "출장비" and "일당" in n)
    unit = "(100엔 기준)" if jpy else ""

    ws["B2"] = rec["부서"]
    ws["D2"] = rec["출장자성명"]
    ws["F2"] = rec["직급"]
    ws["B3"] = rec["출장지"]
    ws["D3"] = rec["지역구분"]
    ws["F3"] = rec["직급구분"]
    ws["B4"] = datetime.date.fromisoformat(rec["출장시작일"])
    ws["D4"] = datetime.date.fromisoformat(rec["출장종료일"])
    ws["B4"].number_format = "yyyy-mm-dd"
    ws["D4"].number_format = "yyyy-mm-dd"
    ws["F4"] = f"{nights}박 {days}일"

    if std_hotel == "실비":
        ws["B8"] = "실비"
        ws["E8"] = "실비 정산"
    else:
        ws["B8"] = f"{std_hotel:,}{cur} / 박"
        ws["E8"] = f"{std_hotel:,} {cur} * {nights}박 x 환율{unit}"
    ws["C8"] = f"{nights}박"
    ws["D8"] = hotel_amt

    ws["B9"] = f"{std_daily:,}{cur} / 일"
    ws["C9"] = f"{days}일"
    ws["D9"] = daily_amt
    ws["E9"] = f"{std_daily:,} {cur} * {days}일 x 환율{unit}"

    ws["D10"] = "=D8+D9"
    for cell in ("D8", "D9", "D10"):
        ws[cell].number_format = "#,##0"

    ws["A37"] = f"신청일 : {datetime.date.today():%Y년 %m월 %d일}"

def build_statement_excel(records, indices):
    wb = load_statement_template()
    tpl = wb.worksheets[0]
    used = set()
    for i in indices:
        rec = records[i]
        ws = wb.copy_worksheet(tpl)
        base = "".join(ch for ch in f'{i + 1}_{rec["출장자성명"]}' if ch not in '[]:*?/\\')[:31]
        title, n = base, 1
        while title in used:
            n += 1
            title = f"{base[:28]}_{n}"
        used.add(title)
        ws.title = title
        fill_statement_sheet(ws, rec)
        ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.remove(tpl)
    wb.calculation.fullCalcOnLoad = True
    return _xl_bytes(wb)

# ----------------------------------------------------
# 4. 페이지 구성
# ----------------------------------------------------
department_list = [
    "임원", "경영지원본부", "경영지원실", "인사지원팀", "관리팀", "재무전략실",
    "노동조합", "재무팀", "자금팀", "정보실", "정보팀", "IBU", "성장전략실",
    "프로젝트팀", "구매전략본부", "HTB 대만지사", "구매팀", "VI팀", "품질혁신본부",
    "QM팀", "보전팀", "생산본부", "생산관리팀", "생산기술팀", "가공팀", "F/S가공",
    "정밀가공", "가공지원", "UNIT팀", "UNIT준비", "UNIT조립", "UNIT서비스",
    "생산1팀", "생산2팀", "서비스센터", "서비스1팀", "서비스2팀", "서비스3팀",
    "서비스4팀", "기술개발연구소", "MC개발팀", "TC개발팀", "5축개발팀", "UNIT개발팀",
    "제어개발팀", "제어SW개발팀", "가공기술1팀", "가공기술2팀", "소재사업부문", "기타",
]

if menu == MENU_1:
    st.markdown("### 📋 출장 기본 정보 입력")
    if st.session_state.edit_target_index is not None:
        st.info(f"✏️ 현재 **[인덱스 {st.session_state.edit_target_index}]** 번 출장 내역 수정 중입니다.")

    target_edit_data = None
    if st.session_state.edit_target_index is not None and len(st.session_state.travel_list) > st.session_state.edit_target_index:
        target_edit_data = st.session_state.travel_list[st.session_state.edit_target_index]

    col_a, col_b = st.columns(2)

    with col_a:
        default_name = target_edit_data["출장자성명"] if target_edit_data else ""
        name = st.text_input("출장자 성명", value=default_name)

        default_dept = target_edit_data["부서"] if target_edit_data else "인사지원팀"
        dept_idx = department_list.index(default_dept) if default_dept in department_list else 0
        department = st.selectbox("부서", department_list, index=dept_idx)

        position_list = [
            "사장", "부사장", "전무", "상무", "이사", "부장", "차장",
            "과장", "대리", "계장", "사원", "1급기능장", "2급기능장",
        ]
        default_pos = target_edit_data["직급"] if target_edit_data else "사원"
        pos_idx = position_list.index(default_pos) if default_pos in position_list else 10
        position = st.selectbox("직급", position_list, index=pos_idx)

        auto_pos_group = get_position_group(position)
        pos_group_options = ["임원(부사장이상)", "임원", "1급", "2급", "3급이하"]

        default_pos_group = target_edit_data.get("직급구분", auto_pos_group) if target_edit_data else auto_pos_group
        default_pos_idx = pos_group_options.index(default_pos_group) if default_pos_group in pos_group_options else 4
        position_group = st.selectbox("직급 구분", pos_group_options, index=default_pos_idx)

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
        country = st.text_input("출장지 (국가/도시)", value=default_country)

        auto_region = get_region_group(country) if country else "갑"
        region_options = ["갑", "을", "병", "특"]

        default_region = target_edit_data.get("지역구분", auto_region) if target_edit_data else auto_region
        default_reg_idx = region_options.index(default_region) if default_region in region_options else 0
        region_group = st.selectbox("지역 구분", region_options, index=default_reg_idx)

    default_purpose = target_edit_data.get("출장목적", "") if target_edit_data else ""
    purpose = st.text_input("출장 목적 (예: JIMTOF 2026 전시회 참관 및 해외 고객사 미팅)", value=default_purpose)

    is_jpy = region_group == "특"
    rate_label = "적용 환율 입력 (일본 엔화 JPY · 100엔 기준)" if is_jpy else "적용 환율 입력 (미국 달러 USD · 1달러 기준)"
    rate_key = "exchange_rate_jpy" if is_jpy else "exchange_rate_usd"
    rate_fallback = 900.00 if is_jpy else 1345.30

    default_exchange_rate = rate_fallback
    if target_edit_data and (target_edit_data.get("지역구분") == "특") == is_jpy:
        default_exchange_rate = target_edit_data["환율"]

    col_rate_info, col_rate_input = st.columns(2)
    with col_rate_info:
        st.markdown("🔗 [서울외국환중개 환율 조회 링크](http://www.smbs.biz/ExRate/TodayExRate.jsp)")
        st.caption("🌐 조회한 환율을 우측 칸에 입력해주세요.")
    with col_rate_input:
        exchange_rate = st.number_input(
            rate_label, min_value=0.0, value=float(default_exchange_rate), step=1.0, format="%.2f", key=rate_key
        )

    raw_days = max(1, (end_date - start_date).days + 1)
    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        is_flight_minus = st.checkbox("기내 박 적용(숙박 1박 차감)", value=False)
    with col_opt2:
        calculated_days = raw_days
        calculated_nights = calculated_days - 1 if calculated_days > 1 else 0
        if is_flight_minus:
            calculated_nights = max(0, calculated_nights - 1)
        st.info(f"📅 최종 산정된 출장 기간: **{calculated_nights}박 {calculated_days}일**")

    std_daily, std_hotel = get_standard_rates(region_group, position_group)
    applied_rate = exchange_rate / 100.0 if region_group == "특" else exchange_rate
    auto_calc_daily = int((std_daily * applied_rate * calculated_days) // 1000 * 1000)
    auto_calc_hotel = 0 if std_hotel == "실비" else int((std_hotel * applied_rate * calculated_nights) // 1000 * 1000)

    if target_edit_data is None:
        if "prev_auto_calc_hotel" not in st.session_state or st.session_state.prev_auto_calc_hotel != auto_calc_hotel:
            st.session_state.prev_auto_calc_hotel = auto_calc_hotel
            st.session_state["te_amt_str_0"] = f"{auto_calc_hotel:,}"
        if "prev_auto_calc_daily" not in st.session_state or st.session_state.prev_auto_calc_daily != auto_calc_daily:
            st.session_state.prev_auto_calc_daily = auto_calc_daily
            st.session_state["te_amt_str_1"] = f"{auto_calc_daily:,}"

    st.markdown("---")
    st.markdown("### 💵 금액 상세 입력")

    col_ratios = [1.2, 1.5, 2, 1.5]
    th1, th2, th3, th4 = st.columns(col_ratios)
    with th1: st.markdown("<div style='text-align: center;'><b>구분</b></div>", unsafe_allow_html=True)
    with th2: st.markdown("<div style='text-align: center;'><b>항목</b></div>", unsafe_allow_html=True)
    with th3: st.markdown("<div style='text-align: center;'><b>금액</b></div>", unsafe_allow_html=True)
    with th4: st.markdown("<div style='text-align: center;'><b>지급처</b></div>", unsafe_allow_html=True)

    st.markdown("---")
    payer_options = ["여행사", "출장자", "직접입력"]
    def payer_fmt(v): return (name.strip() or "출장자") if v == "출장자" else v

    updated_transport_rows = []
    transport_sum = 0
    for idx, row_data in enumerate(st.session_state.transport_rows):
        tc1, tc2, tc3, tc4 = st.columns(col_ratios)
        with tc1: st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>교통비</b></div>", unsafe_allow_html=True)
        with tc2: t_name = st.text_input(f"교통비 항목 {idx}", value=row_data["item"], key=f"t_item_{idx}", label_visibility="collapsed")
        with tc3:
            key_prefix = f"t_amt_str_{idx}"
            if key_prefix not in st.session_state: st.session_state[key_prefix] = f"{int(row_data['amount']):,}"
            def make_on_change(k):
                def callback():
                    digits = "".join(filter(str.isdigit, st.session_state[k]))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback
            amt_str = st.text_input(f"교통비 금액 {idx}", key=key_prefix, on_change=make_on_change(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            t_amt = int((int(digits or 0) // 1000) * 1000)
        with tc4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 0
            t_payer = st.selectbox(f"교통비 지급처 {idx}", payer_options, index=p_idx, key=f"t_payer_{idx}", format_func=payer_fmt, label_visibility="collapsed")
        transport_sum += t_amt
        updated_transport_rows.append({"item": t_name, "amount": t_amt, "payer": t_payer})
    st.session_state.transport_rows = updated_transport_rows

    updated_travel_exp_rows = []
    travel_exp_sum = 0
    for idx, row_data in enumerate(st.session_state.travel_exp_rows):
        tec1, tec2, tec3, tec4 = st.columns(col_ratios)
        with tec1: st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>출장비</b></div>", unsafe_allow_html=True)
        with tec2: te_name = st.text_input(f"출장비 항목 {idx}", value=row_data["item"], key=f"te_item_{idx}", label_visibility="collapsed")
        with tec3:
            key_prefix = f"te_amt_str_{idx}"
            if key_prefix not in st.session_state: st.session_state[key_prefix] = f"{int(row_data['amount']):,}"
            def make_on_change_te(k):
                def callback():
                    digits = "".join(filter(str.isdigit, st.session_state[k]))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback
            amt_str = st.text_input(f"출장비 금액 {idx}", key=key_prefix, on_change=make_on_change_te(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            te_amt = int((int(digits or 0) // 1000) * 1000)
        with tc4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 1
            te_payer = st.selectbox(f"출장비 지급처 {idx}", payer_options, index=p_idx, key=f"te_payer_{idx}", format_func=payer_fmt, label_visibility="collapsed")
        travel_exp_sum += te_amt
        updated_travel_exp_rows.append({"item": te_name, "amount": te_amt, "payer": te_payer})
    st.session_state.travel_exp_rows = updated_travel_exp_rows

    updated_other_rows = []
    other_sum = 0
    for idx, row_data in enumerate(st.session_state.other_rows):
        oc1, oc2, oc3, oc4 = st.columns(col_ratios)
        with oc1: st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>기타</b></div>", unsafe_allow_html=True)
        with oc2: it_name = st.text_input(f"기타 항목명 {idx}", value=row_data["item"], placeholder="항목 입력", key=f"other_item_{idx}", label_visibility="collapsed")
        with oc3:
            key_prefix = f"other_amt_str_{idx}"
            if key_prefix not in st.session_state: st.session_state[key_prefix] = f"{int(row_data['amount']):,}"
            def make_on_change_other(k):
                def callback():
                    digits = "".join(filter(str.isdigit, st.session_state[k]))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback
            amt_str = st.text_input(f"기타 금액 {idx}", key=key_prefix, on_change=make_on_change_other(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            it_amt = int((int(digits or 0) // 1000) * 1000)
        with oc4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 0
            it_payer = st.selectbox(f"기타 지급처 {idx}", payer_options, index=p_idx, key=f"other_payer_{idx}", format_func=payer_fmt, label_visibility="collapsed")
        other_sum += it_amt
        updated_other_rows.append({"item": it_name, "amount": it_amt, "payer": it_payer})
    st.session_state.other_rows = updated_other_rows

    real_employee_total = sum(x["amount"] for x in st.session_state.travel_exp_rows + st.session_state.transport_rows + st.session_state.other_rows if x.get("payer") != "여행사")
    real_agency_total = sum(x["amount"] for x in st.session_state.travel_exp_rows + st.session_state.transport_rows + st.session_state.other_rows if x.get("payer") == "여행사")
    grand_total = real_employee_total + real_agency_total

    st.markdown("<br>", unsafe_allow_html=True)
    tot_c1, tot_c2, tot_c3, tot_c4 = st.columns(col_ratios)
    with tot_c2: st.markdown("<div class='row-highlight-blue' style='text-align: center;'><b>총 출장 경비 합계</b></div>", unsafe_allow_html=True)
    with tot_c3: st.markdown(f"<div class='row-highlight-blue' style='text-align: right;'><b>{grand_total:,.0f} 원</b></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    btn_label = "🔄 수정 사항 반영하기" if st.session_state.edit_target_index is not None else "➕ 입력 내역 저장 및 등록"
    if st.button(btn_label, use_container_width=True):
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
                "출장목적": purpose if purpose else "해외 출장 업무 수행",
                "출장시작일": str(start_date),
                "출장종료일": str(end_date),
                "출장일수": calculated_days,
                "출장박수": calculated_nights,
                "환율": exchange_rate,
                "직원_지급액": real_employee_total,
                "여행사_지급액": real_agency_total,
                "총출장비": grand_total,
                "교통비항목리스트": st.session_state.transport_rows.copy(),
                "출장비항목리스트": st.session_state.travel_exp_rows.copy(),
                "기타항목리스트": st.session_state.other_rows.copy(),
            }

            if st.session_state.edit_target_index is not None:
                st.session_state.travel_list[st.session_state.edit_target_index] = new_data
                st.success(f"✅ [{name}] 님의 내역이 수정되었습니다.")
                st.session_state.edit_target_index = None
            else:
                st.session_state.travel_list.append(new_data)
                st.success(f"✅ [{name}] 님의 내역이 등록되었습니다.")
            st.rerun()

    st.markdown("---")
    st.markdown("### 📊 현재 정산 대기 중인 출장 내역")

    if len(st.session_state.travel_list) > 0:
        raw_df = process_travel_data(st.session_state.travel_list)
        display_df = raw_df.drop(columns=["교통비항목리스트", "출장비항목리스트", "기타항목리스트", "지역구분", "직급구분", "환율"]).reset_index(drop=True)
        display_df.insert(0, "순번", range(1, len(display_df) + 1))

        grid_response = AgGrid(
            display_df,
            gridOptions=GridOptionsBuilder.from_dataframe(display_df).build(),
            update_mode=GridUpdateMode.MODEL_CHANGED,
            fit_columns_on_grid_load=True,
            height=200,
            theme="alpine",
        )

        col_act1, col_act2 = st.columns(2)
        with col_act1:
            if st.button("✈️ 결재/출장 완료 처리 (완료 기록으로 이관)", use_container_width=True):
                today_str = str(datetime.date.today())
                for item in st.session_state.travel_list:
                    item["완료일자"] = today_str
                    st.session_state.completed_travel_list.append(item)
                st.session_state.travel_list = []
                st.success("✅ 모든 정산건이 내부 결재 및 출장 완료 처리되어 '4. 해외출장 기록 관리' 페이지로 이관되었습니다.")
                st.rerun()
        with col_act2:
            if st.button("🗑️ 대기목록 초기화", use_container_width=True):
                st.session_state.travel_list = []
                st.session_state.edit_target_index = None
                st.rerun()
    else:
        st.info("현재 정산 대기 중인 출장 내역이 없습니다.")

elif menu == MENU_2:
    st.markdown("### 💰 자금팀 제출용 정산 집계표")
    target_data = st.session_state.travel_list if len(st.session_state.travel_list) > 0 else st.session_state.completed_travel_list
    
    if len(target_data) > 0:
        processed_df = process_travel_data(target_data)
        m1, m2, m3 = st.columns(3)
        with m1: st.metric(label="총 출장 건수", value=f"{len(processed_df)} 건")
        with m2: st.metric(label="총 여행사 송금액", value=f"{processed_df['여행사_지급액'].sum():,.0f} 원")
        with m3: st.metric(label="총 직원 지급액", value=f"{processed_df['직원_지급액'].sum():,.0f} 원")

        fund_view = processed_df[["출장자성명", "부서", "직급", "출장지", "출장목적", "직원_지급액", "여행사_지급액", "총출장비"]].copy()
        fund_view.index = range(1, len(fund_view) + 1)
        fund_view.index.name = "순번"
        st.dataframe(
            fund_view.style.format({"직원_지급액": "{:,.0f}", "여행사_지급액": "{:,.0f}", "총출장비": "{:,.0f}"}),
            use_container_width=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            "📥 자금팀 연결 자료 엑셀 다운로드",
            data=build_fund_excel(target_data),
            file_name=f"자금팀_연결자료_{datetime.date.today():%Y%m%d}.xlsx",
            mime=XL_MIME,
            use_container_width=True,
        )
    else:
        st.info("등록된 출장 내역이 없습니다.")

elif menu == MENU_3:
    st.markdown("### 📄 해외출장비 산정 내역서")
    target_data = st.session_state.travel_list if len(st.session_state.travel_list) > 0 else st.session_state.completed_travel_list

    if len(target_data) > 0:
        records = target_data
        sel_idx = st.selectbox(
            "출장자 선택",
            list(range(len(records))),
            format_func=lambda i: f"{i + 1}. {records[i]['출장자성명']} ({records[i]['출장지']}, {records[i]['출장시작일']})",
        )
        selected_person = records[sel_idx]["출장자성명"]
        st.success(f"[{selected_person}] 님의 해외출장 산정 내역서가 준비되었습니다.")

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                f"📥 [{selected_person}] 산정 내역서 엑셀 다운로드",
                data=build_statement_excel(records, [sel_idx]),
                file_name=f"출장비_산정내역서_{selected_person}_{datetime.date.today():%Y%m%d}.xlsx",
                mime=XL_MIME,
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "📥 전체 출장자 일괄 다운로드 (출장자별 시트)",
                data=build_statement_excel(records, list(range(len(records)))),
                file_name=f"출장비_산정내역서_전체_{datetime.date.today():%Y%m%d}.xlsx",
                mime=XL_MIME,
                use_container_width=True,
            )
    else:
        st.info("등록된 출장 내역이 없습니다.")

# ----------------------------------------------------
# 5. 메뉴 4: 해외출장 기록 관리 (완료 건 검색 및 리포팅)
# ----------------------------------------------------
elif menu == MENU_4:
    st.markdown("### 🗂️ 완료된 해외출장 기록 조회 및 다차원 검색")

    if len(st.session_state.completed_travel_list) == 0:
        st.warning("현재 보관된 해외출장 완료 기록이 없습니다. '1. 출장 정보 입력' 메뉴에서 출장 처리 후 승인/완료를 진행해주세요.")
    else:
        c_df = process_travel_data(st.session_state.completed_travel_list)

        with st.expander("🔍 다차원 조건 검색 필터 (출장목적 / 부서 / 출장자 / 기간 / 국가)", expanded=True):
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                search_purpose = st.text_input("🎯 출장 목적 키워드 검색", value="")
                search_name = st.text_input("👤 출장자 성명 검색", value="")
            with f_col2:
                all_depts = ["전체"] + sorted(list(c_df["부서"].unique()))
                search_dept = st.selectbox("🏢 부서 선택", all_depts)
                search_country = st.text_input("🌍 출장 국가/지역 검색", value="")
            with f_col3:
                min_date = datetime.date.fromisoformat(c_df["출장시작일"].min())
                max_date = datetime.date.fromisoformat(c_df["출장종료일"].max())
                search_period = st.date_input("📅 출장 기간 검색 범위", value=(min_date, max_date))

        # 조건 필터링 엔진
        filtered_df = c_df.copy()

        if search_purpose.strip():
            filtered_df = filtered_df[filtered_df["출장목적"].str.contains(search_purpose.strip(), case=False, na=False)]
        if search_name.strip():
            filtered_df = filtered_df[filtered_df["출장자성명"].str.contains(search_name.strip(), case=False, na=False)]
        if search_dept != "전체":
            filtered_df = filtered_df[filtered_df["부서"] == search_dept]
        if search_country.strip():
            filtered_df = filtered_df[filtered_df["출장지"].str.contains(search_country.strip(), case=False, na=False)]
        if isinstance(search_period, tuple) and len(search_period) == 2:
            s_start, s_end = search_period
            filtered_df = filtered_df[
                (pd.to_datetime(filtered_df["출장시작일"]) >= pd.to_datetime(s_start)) &
                (pd.to_datetime(filtered_df["출장종료일"]) <= pd.to_datetime(s_end))
            ]

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📈 검색 결과 통계 요약")
        s1, s2, s3, s4 = st.columns(4)
        with s1: st.metric("조회된 출장 건수", f"{len(filtered_df)} 건")
        with s2: st.metric("총 집행 경비", f"{filtered_df['총출장비'].sum():,.0f} 원")
        with s3: st.metric("총 직원 지급액", f"{filtered_df['직원_지급액'].sum():,.0f} 원")
        with s4: st.metric("총 여행사 정산액", f"{filtered_df['여행사_지급액'].sum():,.0f} 원")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 상세 출장 기록 명세")

        disp_cols = ["출장자성명", "부서", "직급", "출장지", "출장목적", "출장시작일", "출장종료일", "출장일수", "총출장비", "완료일자"]
        existing_cols = [c for c in disp_cols if c in filtered_df.columns]
        view_result = filtered_df[existing_cols].copy().reset_index(drop=True)
        view_result.index = range(1, len(view_result) + 1)
        view_result.index.name = "순번"

        st.dataframe(
            view_result.style.format({"총출장비": "{:,.0f}"}),
            use_container_width=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="해외출장_기록_검색결과")

        st.download_button(
            "📥 검색 결과 엑셀 파일 다운로드",
            data=out_buf.getvalue(),
            file_name=f"해외출장기록_조회결과_{datetime.date.today():%Y%m%d}.xlsx",
            mime=XL_MIME,
            use_container_width=True,
        )
