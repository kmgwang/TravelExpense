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
        page_title="HWACHEON - 해외출장비 정산 자동화 시스템",
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

    /* 폰트 (아이콘 폰트가 깨지지 않도록 span 은 제외) */
    html, body, .stApp, .stApp p, .stApp label, .stApp input, .stApp textarea,
    .stApp button, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5,
    .stApp td, .stApp th, .stApp li, .stApp a, .stApp div[data-baseweb] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto,
                     'Helvetica Neue', 'Segoe UI', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
    }

    /* 메인 배경 및 레이아웃 */
    .stApp { background-color: #ffffff; color: var(--hw-text); }
    header[data-testid="stHeader"] { background: transparent; }
    .block-container {
        max-width: 1200px !important;
        padding-top: 2.5rem;
        padding-bottom: 4rem;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
    }

    /* ---------- 사이드바 ---------- */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid var(--hw-line);
        min-width: 290px;
        max-width: 290px;
    }
    section[data-testid="stSidebar"] > div:first-child { padding-top: 0.5rem; }

    /* 사이드바 항상 열림 고정 (접기/펼치기 버튼 제거) */
    section[data-testid="stSidebar"] {
        transform: none !important;
        margin-left: 0 !important;
        visibility: visible !important;
        min-width: 290px !important;
        max-width: 290px !important;
        width: 290px !important;
    }
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarResizeHandle"],
    section[data-testid="stSidebar"] button[kind="header"],
    section[data-testid="stSidebar"] button[kind="headerNoPadding"] {
        display: none !important;
    }

    /* 사이드바 상단 브랜드 영역 (로고 + 시스템 설명 문구) */
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

    /* 사이드바 메뉴 (라디오 → 메뉴 리스트) */
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
        display: none;               /* 라디오 동그라미 숨김 */
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label p {
        font-size: 16px;
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

    /* ---------- 페이지 헤드 ---------- */
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

    /* 제목 */
    .stApp h3 {
        font-weight: 800;
        letter-spacing: -0.3px;
        color: var(--hw-text);
    }

    /* 카드 스타일 섹션 (홈페이지 게시물 카드 느낌) */
    .hw-card {
        background: var(--hw-gray-bg);
        padding: 28px;
        border-radius: 20px;
        margin-bottom: 20px;
    }

    /* 하이라이트 박스 */
    .row-highlight-yellow {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 10px 14px;
        border-radius: 10px;
        font-weight: 600;
    }
    .row-highlight-blue {
        background-color: var(--hw-blue-soft);
        border: none;
        color: var(--hw-blue);
        padding: 12px 16px;
        border-radius: 10px;
        font-size: 16px;
    }

    /* 입력 필드 (홈페이지 검색창 스타일) */
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
    div[data-testid="stNumberInput"] button { background-color: transparent; }

    /* Input 텍스트 비주얼 개선 */
    input[aria-label*="항목"] { text-align: center !important; }
    input[aria-label*="금액"] { text-align: right !important; }

    /* 버튼 */
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
        border: none;
    }
    .stButton > button:focus:not(:active) { color: #ffffff; border: none; }

    /* 체크박스 포인트 컬러 */
    label[data-baseweb="checkbox"] > span:first-child[data-checked="true"],
    div[data-baseweb="checkbox"] > div[aria-checked="true"] {
        background-color: var(--hw-blue) !important;
    }

    /* 알림 박스 */
    div[data-testid="stAlert"] { border-radius: 12px; }

    /* 메트릭 카드 */
    div[data-testid="stMetric"] {
        background-color: var(--hw-gray-bg);
        border-radius: 20px;
        padding: 22px 26px;
    }
    div[data-testid="stMetricLabel"] p { color: var(--hw-text-sub); font-weight: 600; }
    div[data-testid="stMetricValue"] { color: var(--hw-blue); font-weight: 800; }

    /* 구분선 */
    .stApp hr { border-color: var(--hw-line); }

    /* 링크 */
    .stApp a { color: var(--hw-blue); }
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

# 주의: 마크다운이 코드 블록으로 오인하지 않도록 HTML은 빈 줄/들여쓰기 없이 한 줄로 구성
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
    '<div class="hw-side-desc">화천기공 해외출장 경비 산정 &amp; 자금팀 정산 자동화 시스템</div>'
    "</div>"
    '<div class="hw-side-caption">MENU</div>'
)

with st.sidebar:
    st.markdown(SIDEBAR_BRAND_HTML, unsafe_allow_html=True)
    menu = st.radio(
        "메뉴",
        [MENU_1, MENU_2, MENU_3],
        key="nav_menu",
        label_visibility="collapsed",
    )

# 페이지 상단 헤드
PAGE_HEAD = {
    MENU_1: ("STEP 01", "출장 정보 입력", "출장자 정보와 경비 항목을 입력하고 등록합니다."),
    MENU_2: ("STEP 02", "자금팀 연결 자료 생성", "등록된 출장 내역을 자금팀 제출용으로 집계합니다."),
    MENU_3: ("STEP 03", "출장비 산정 내역서 생성", "출장자별 해외출장비 산정 내역서를 확인합니다."),
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

        processed.append(ordered_d)

    df = pd.DataFrame(processed)
    return df


# ----------------------------------------------------
# 엑셀 생성 함수 (자금팀 연결 자료 / 출장비 산정 내역서)
# ----------------------------------------------------
import openpyxl.worksheet.properties
from openpyxl.utils import get_column_letter

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
    """지급처 표시명: '출장자'는 1번 메뉴에서 입력한 출장자 성명으로 대체"""
    if payer == "출장자":
        return (rec.get("출장자성명") or "").strip() or "출장자"
    return payer


def calc_item_rows(record):
    """출장 1건의 (구분, 항목, 금액, 지급처) 목록과 지급 기준값을 반환 (금액은 입력한 값 그대로 사용)"""
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
        ws.column_dimensions[get_column_letter(i)].width = w


def _xl_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_fund_excel(records):
    """2. 자금팀 연결 자료 엑셀 (정산 집계표 + 항목별 상세)"""
    df = process_travel_data(records)
    wb = openpyxl.Workbook()

    # ---- 시트 1: 정산 집계표 ----
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
        col = get_column_letter(j)
        c = ws.cell(row=tot, column=j, value=f"=SUM({col}{first}:{col}{last})")
        c.number_format = XL_NUM
    _xl_widths(ws, [7, 12, 16, 10, 14, 24, 22, 20, 18])
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)

    # ---- 시트 2: 항목별 상세 ----
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


# 출장비 산정 내역서 엑셀 양식 (제공해주신 '출장비산정내역' 양식을 코드에 내장)
# ※ 이 파이썬 파일과 같은 폴더에 아래 이름의 엑셀 파일을 두면, 내장 양식 대신 그 파일을 양식으로 사용합니다.
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
    "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"
    "bWxQSwUGAAAAAAwADAAmAwAA6CwAAAAA"
)


def load_statement_template():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATE_FILENAME)
    if os.path.exists(path):
        return openpyxl.load_workbook(path)
    return openpyxl.load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)))


def _find_cell(ws, text, startswith=False, min_row=1, max_row=None, col=None):
    """시트에서 텍스트가 일치하는 첫 셀을 찾는다"""
    for row in ws.iter_rows(min_row=min_row, max_row=max_row):
        for c in row:
            if c.value is None or (col is not None and c.column != col):
                continue
            v = str(c.value).strip()
            if (v.startswith(text) if startswith else v == text):
                return c
    return None


def _put(ws, row, col, value, number_format=None):
    """셀에 값을 입력 (병합된 셀의 비대표 칸이면 건너뜀)"""
    cell = ws.cell(row=row, column=col)
    try:
        cell.value = value
    except AttributeError:
        return
    if number_format:
        cell.number_format = number_format


def fill_statement_sheet(ws, rec):
    """양식 시트에 출장자 정보와 입력한 금액을 채워 넣는다 (양식의 서식/병합은 그대로 유지)
    - 양식의 라벨('소속', '성명', '숙박비' 등) 위치를 찾아 입력하므로 양식의 행/열이 이동해도 동작합니다."""
    items, std_daily, std_hotel, cur = calc_item_rows(rec)
    jpy = rec["지역구분"] == "특"
    days = rec["출장일수"]
    nights = rec["출장박수"]
    hotel_amt = sum(a for cat, n, a, p in items if cat == "출장비" and "숙박" in n)
    daily_amt = sum(a for cat, n, a, p in items if cat == "출장비" and "일당" in n)
    unit = "(100엔 기준)" if jpy else ""

    # 1) 출장자 정보: 라벨 오른쪽 칸에 입력 ('■ 적용 출장비 산정 기준' 윗부분에서만 탐색)
    sec = _find_cell(ws, "■ 적용 출장비 산정 기준")
    info_max = sec.row - 1 if sec else 10
    info = {
        "소속": rec["부서"],
        "성명": rec["출장자성명"],
        "직급": rec["직급"],
        "출장지": rec["출장지"],
        "지역구분": rec["지역구분"],
        "직급구분": rec["직급구분"],
        "출발일": datetime.date.fromisoformat(rec["출장시작일"]),
        "도착일": datetime.date.fromisoformat(rec["출장종료일"]),
        "출장기간": f"{nights}박 {days}일",
    }
    for label, value in info.items():
        c = _find_cell(ws, label, max_row=info_max)
        if c is not None:
            _put(ws, c.row, c.column + 1, value, "yyyy-mm-dd" if isinstance(value, datetime.date) else None)

    # 2) 적용 출장비 산정 기준 표
    hdr = _find_cell(ws, "기간 적용")
    if hdr is not None:
        cols = {str(c.value).strip(): c.column for c in ws[hdr.row] if c.value is not None}
        c_label = cols.get("구분", 1)
        c_basis = cols.get("산정 기준")
        c_period = cols.get("기간 적용")
        c_amt = cols.get("금액")
        c_calc = cols.get("원화 환산 산식")
        r_hotel = _find_cell(ws, "숙박비", min_row=hdr.row + 1, col=c_label)
        r_daily = _find_cell(ws, "일당", min_row=hdr.row + 1, col=c_label)
        r_total = _find_cell(ws, "지급 총액", min_row=hdr.row + 1, col=c_label)

        if r_hotel is not None:
            r = r_hotel.row
            if std_hotel == "실비":
                basis, calc = "실비", "실비 정산"
            else:
                basis, calc = f"{std_hotel:,}{cur} / 박", f"{std_hotel:,} {cur} * {nights}박 x 환율{unit}"
            if c_basis: _put(ws, r, c_basis, basis)
            if c_period: _put(ws, r, c_period, f"{nights}박")
            if c_amt: _put(ws, r, c_amt, hotel_amt, "#,##0")
            if c_calc: _put(ws, r, c_calc, calc)

        if r_daily is not None:
            r = r_daily.row
            if c_basis: _put(ws, r, c_basis, f"{std_daily:,}{cur} / 일")
            if c_period: _put(ws, r, c_period, f"{days}일")
            if c_amt: _put(ws, r, c_amt, daily_amt, "#,##0")
            if c_calc: _put(ws, r, c_calc, f"{std_daily:,} {cur} * {days}일 x 환율{unit}")

        if r_total is not None and c_amt and r_hotel is not None and r_daily is not None:
            L = get_column_letter(c_amt)
            _put(ws, r_total.row, c_amt, f"={L}{r_hotel.row}+{L}{r_daily.row}", "#,##0")

    # 3) 신청일
    d = _find_cell(ws, "신청일", startswith=True)
    if d is not None:
        d.value = f"신청일 : {datetime.date.today():%Y년 %m월 %d일}"


def build_statement_excel(records, indices):
    """3. 출장비 산정 내역서 엑셀 (제공된 양식 기준, 선택한 출장자별 시트)"""
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
        # 인쇄 시 F열(직급 등)이 다음 장으로 넘어가지 않도록 가로 1페이지에 맞춤
        ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    wb.remove(tpl)
    wb.calculation.fullCalcOnLoad = True
    return _xl_bytes(wb)


# ----------------------------------------------------
# 4. 페이지 구성 (사이드바 메뉴 선택에 따라 표시)
# ----------------------------------------------------
if menu == MENU_1:
    st.markdown("### 📋 출장 기본 정보 입력")
    if st.session_state.edit_target_index is not None:
        st.info(
            f"✏️ 현재 **[인덱스 {st.session_state.edit_target_index}]** 번 출장 내역 수정 중입니다."
        )

    target_edit_data = None
    if st.session_state.edit_target_index is not None and len(
        st.session_state.travel_list
    ) > st.session_state.edit_target_index:
        target_edit_data = st.session_state.travel_list[
            st.session_state.edit_target_index
        ]

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
            "사장", "부사장", "전무", "상무", "이사", "부장", "차장",
            "과장", "대리", "계장", "사원", "1급기능장", "2급기능장",
        ]
        default_pos = target_edit_data["직급"] if target_edit_data else "사원"
        pos_idx = (
            position_list.index(default_pos)
            if default_pos in position_list
            else 10
        )
        position = st.selectbox("직급", position_list, index=pos_idx)

        auto_pos_group = get_position_group(position)
        pos_group_options = ["임원(부사장이상)", "임원", "1급", "2급", "3급이하"]

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

    # ------------------------------------------------
    # 환율 입력 (지역 구분에 따라 통화 자동 전환)
    #   갑/을/병 → 미국 달러(USD, 1달러 기준)
    #   특       → 일본 엔화(JPY, 100엔 기준)
    # ------------------------------------------------
    is_jpy = region_group == "특"
    if is_jpy:
        rate_label = "적용 환율 입력 (일본 엔화 JPY · 100엔 기준)"
        rate_key = "exchange_rate_jpy"
        rate_fallback = 900.00
    else:
        rate_label = "적용 환율 입력 (미국 달러 USD · 1달러 기준)"
        rate_key = "exchange_rate_usd"
        rate_fallback = 1345.30

    default_exchange_rate = rate_fallback
    if target_edit_data and (target_edit_data.get("지역구분") == "특") == is_jpy:
        default_exchange_rate = target_edit_data["환율"]

    col_rate_info, col_rate_input = st.columns(2)
    with col_rate_info:
        st.markdown(
            "🔗 [서울외국환중개 환율 조회 링크](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
        )
        if is_jpy:
            st.caption("🌐 위 링크에서 조회한 일본 엔화(JPY 100엔) 환율을 우측 칸에 입력해주세요.")
        else:
            st.caption("🌐 위 링크에서 조회한 미국 달러(USD) 환율을 우측 칸에 입력해주세요.")
    with col_rate_input:
        exchange_rate = st.number_input(
            rate_label,
            min_value=0.0,
            value=float(default_exchange_rate),
            step=1.0,
            format="%.2f",
            key=rate_key,
        )

    raw_days = (end_date - start_date).days + 1
    if raw_days < 1:
        raw_days = 1

    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        is_flight_minus = st.checkbox("기내 박 적용(숙박 1박 차감)", value=False)
    with col_opt2:
        calculated_days = raw_days
        calculated_nights = calculated_days - 1 if calculated_days > 1 else 0
        if is_flight_minus:
            calculated_nights = max(0, calculated_nights - 1)

        st.info(f"📅 최종 산정된 출장 기간: **{calculated_nights}박 {calculated_days}일**")

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
    st.markdown("### 💵 금액 상세 입력")

    col_ratios = [1.2, 1.5, 2, 1.5]

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

    payer_sfx = name.strip()  # 성명이 바뀌면 지급처 위젯 key가 바뀌어 즉시 새로 그려짐

    def payer_fmt(v):
        # 저장값은 '출장자'로 유지하고, 화면에는 입력한 출장자 성명을 표시
        return (name.strip() or "출장자") if v == "출장자" else v

    updated_transport_rows = []
    transport_sum = 0
    for idx, row_data in enumerate(st.session_state.transport_rows):
        tc1, tc2, tc3, tc4 = st.columns(col_ratios)
        with tc1:
            st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>교통비</b></div>", unsafe_allow_html=True)
        with tc2:
            t_name = st.text_input(f"교통비 항목 {idx}", value=row_data["item"], key=f"t_item_{idx}", label_visibility="collapsed")
        with tc3:
            key_prefix = f"t_amt_str_{idx}"
            if key_prefix not in st.session_state:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"

            def make_on_change(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback

            amt_str = st.text_input(f"교통비 금액 {idx}", key=key_prefix, on_change=make_on_change(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            t_amt = int((int(digits or 0) // 1000) * 1000)
        with tc4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 0
            t_payer = st.selectbox(f"교통비 지급처 {idx}", payer_options, index=p_idx, key=f"t_payer_{idx}_{payer_sfx}", format_func=payer_fmt, label_visibility="collapsed")

        transport_sum += t_amt
        updated_transport_rows.append({"item": t_name, "amount": t_amt, "payer": t_payer})

    st.session_state.transport_rows = updated_transport_rows

    sc1, sc2, sc3, sc4 = st.columns(col_ratios)
    with sc2:
        st.markdown("<div style='text-align: center;'><b>교통비 소계</b></div>", unsafe_allow_html=True)
    with sc3:
        st.markdown(f"<div style='text-align: right; font-weight: bold;'>{transport_sum:,.0f} 원</div>", unsafe_allow_html=True)

    st.markdown("---")

    updated_travel_exp_rows = []
    travel_exp_sum = 0
    for idx, row_data in enumerate(st.session_state.travel_exp_rows):
        tec1, tec2, tec3, tec4 = st.columns(col_ratios)
        with tec1:
            st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>출장비</b></div>", unsafe_allow_html=True)
        with tec2:
            te_name = st.text_input(f"출장비 항목 {idx}", value=row_data["item"], key=f"te_item_{idx}", label_visibility="collapsed")
        with tec3:
            key_prefix = f"te_amt_str_{idx}"
            # 사이드바 메뉴 이동 후 돌아왔을 때 금액이 0으로 초기화되지 않도록
            # 위젯 상태가 없으면 저장된 값으로 복원
            if key_prefix not in st.session_state:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"

            def make_on_change_te(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback

            amt_str = st.text_input(f"출장비 금액 {idx}", key=key_prefix, on_change=make_on_change_te(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            te_amt = int((int(digits or 0) // 1000) * 1000)
        with tec4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 1
            te_payer = st.selectbox(f"출장비 지급처 {idx}", payer_options, index=p_idx, key=f"te_payer_{idx}_{payer_sfx}", format_func=payer_fmt, label_visibility="collapsed")

        travel_exp_sum += te_amt
        updated_travel_exp_rows.append({"item": te_name, "amount": te_amt, "payer": te_payer})

    st.session_state.travel_exp_rows = updated_travel_exp_rows

    tc1, tc2, tc3, tc4 = st.columns(col_ratios)
    with tc2:
        st.markdown("<div style='text-align: center;'><b>출장비 소계</b></div>", unsafe_allow_html=True)
    with tc3:
        st.markdown(f"<div style='text-align: right; font-weight: bold;'>{travel_exp_sum:,.0f} 원</div>", unsafe_allow_html=True)

    st.markdown("---")

    updated_other_rows = []
    other_sum = 0
    for idx, row_data in enumerate(st.session_state.other_rows):
        oc1, oc2, oc3, oc4 = st.columns(col_ratios)
        with oc1:
            st.markdown("<div style='display: flex; align-items: center; height: 45px; justify-content: center;'><b>기타</b></div>", unsafe_allow_html=True)
        with oc2:
            it_name = st.text_input(f"기타 항목명 {idx}", value=row_data["item"], placeholder="항목 입력", key=f"other_item_{idx}", label_visibility="collapsed")
        with oc3:
            key_prefix = f"other_amt_str_{idx}"
            if key_prefix not in st.session_state:
                st.session_state[key_prefix] = f"{int(row_data['amount']):,}"

            def make_on_change_other(k):
                def callback():
                    val = st.session_state[k]
                    digits = "".join(filter(str.isdigit, val))
                    st.session_state[k] = f"{int(digits):,}" if digits else "0"
                return callback

            amt_str = st.text_input(f"기타 금액 {idx}", key=key_prefix, on_change=make_on_change_other(key_prefix), label_visibility="collapsed")
            digits = "".join(filter(str.isdigit, amt_str))
            it_amt = int((int(digits or 0) // 1000) * 1000)
        with oc4:
            p_idx = payer_options.index(row_data["payer"]) if row_data["payer"] in payer_options else 0
            it_payer = st.selectbox(f"기타 지급처 {idx}", payer_options, index=p_idx, key=f"other_payer_{idx}_{payer_sfx}", format_func=payer_fmt, label_visibility="collapsed")

        other_sum += it_amt
        updated_other_rows.append({"item": it_name, "amount": it_amt, "payer": it_payer})

    st.session_state.other_rows = updated_other_rows

    b_c1, b_c2, b_c3 = st.columns([1.2, 2.5, 2.5])
    with b_c2:
        if st.button("➕ 기타 행 추가", use_container_width=True):
            st.session_state.other_rows.append({"item": "", "amount": 0, "payer": "여행사"})
            st.rerun()
    with b_c3:
        if len(st.session_state.other_rows) > 1:
            if st.button("➖ 기타 마지막 행 삭제", use_container_width=True):
                st.session_state.other_rows.pop()
                st.rerun()

    real_employee_total = sum(x["amount"] for x in st.session_state.travel_exp_rows + st.session_state.transport_rows + st.session_state.other_rows if x.get("payer") != "여행사")
    real_agency_total = sum(x["amount"] for x in st.session_state.travel_exp_rows + st.session_state.transport_rows + st.session_state.other_rows if x.get("payer") == "여행사")
    grand_total = real_employee_total + real_agency_total

    st.markdown("<br>", unsafe_allow_html=True)
    tot_c1, tot_c2, tot_c3, tot_c4 = st.columns(col_ratios)
    with tot_c2:
        st.markdown("<div class='row-highlight-blue' style='text-align: center;'><b>총 출장 경비 합계</b></div>", unsafe_allow_html=True)
    with tot_c3:
        st.markdown(f"<div class='row-highlight-blue' style='text-align: right;'><b>{grand_total:,.0f} 원</b></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    btn_label = "🔄 수정 사항 반영하기" if st.session_state.edit_target_index is not None else "➕ 입력 내역 저장 및 등록"
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
    st.markdown("### 📊 현재 등록된 출장 내역")

    if len(st.session_state.travel_list) > 0:
        raw_df = process_travel_data(st.session_state.travel_list)
        display_df = raw_df.drop(columns=["교통비항목리스트", "출장비항목리스트", "기타항목리스트", "지역구분", "직급구분", "환율"]).reset_index(drop=True)
        display_df.insert(0, "순번", range(1, len(display_df) + 1))

        gb = GridOptionsBuilder.from_dataframe(display_df)
        gb.configure_selection(selection_mode="single", use_checkbox=False)
        grid_options = gb.build()

        grid_response = AgGrid(
            display_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            fit_columns_on_grid_load=True,
            height=250,
            theme="alpine",
        )

        col_del1, col_del2 = st.columns([1, 1])
        with col_del1:
            if st.session_state.edit_target_index is not None:
                if st.button("❌ 수정 모드 취소"):
                    st.session_state.edit_target_index = None
                    st.rerun()
        with col_del2:
            if st.button("🗑️ 전체 데이터 초기화"):
                st.session_state.travel_list = []
                st.session_state.edit_target_index = None
                st.rerun()
    else:
        st.info("등록된 출장 내역이 없습니다.")

# ----------------------------------------------------
# 메뉴 2 & 3 (자금팀 / 산정내역서) - 기존 기능 유지
# ----------------------------------------------------
elif menu == MENU_2:
    st.markdown("### 💰 자금팀 제출용 정산 집계표")
    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric(label="총 출장 건수", value=f"{len(processed_df)} 건")
        with m2:
            st.metric(label="총 여행사 송금액", value=f"{processed_df['여행사_지급액'].sum():,.0f} 원")
        with m3:
            st.metric(label="총 직원 지급액", value=f"{processed_df['직원_지급액'].sum():,.0f} 원")

        fund_view = processed_df[["출장자성명", "부서", "직급", "출장지", "직원_지급액", "여행사_지급액", "총출장비"]].copy()
        fund_view.index = range(1, len(fund_view) + 1)
        fund_view.index.name = "순번"
        st.dataframe(
            fund_view.style.format({"직원_지급액": "{:,.0f}", "여행사_지급액": "{:,.0f}", "총출장비": "{:,.0f}"}),
            use_container_width=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            "📥 자금팀 연결 자료 엑셀 다운로드",
            data=build_fund_excel(st.session_state.travel_list),
            file_name=f"자금팀_연결자료_{datetime.date.today():%Y%m%d}.xlsx",
            mime=XL_MIME,
            use_container_width=True,
        )
    else:
        st.info("등록된 출장 내역이 없습니다. '1. 출장 정보 입력' 메뉴에서 먼저 등록해주세요.")

elif menu == MENU_3:
    st.markdown("### 📄 해외출장비 산정 내역서")
    if len(st.session_state.travel_list) > 0:
        processed_df = process_travel_data(st.session_state.travel_list)
        records = st.session_state.travel_list
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
        st.info("등록된 출장 내역이 없습니다. '1. 출장 정보 입력' 메뉴에서 먼저 등록해주세요.")
