import base64
import datetime
import io
import json
import os
import platform
import sqlite3
from contextlib import closing
from copy import copy
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
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

def _make_favicon():
    """브라우저 탭 아이콘: 화천 로고 모양(파란 막대)"""
    try:
        from PIL import Image, ImageDraw

        W = 128
        k = W / 906.0
        oy = (W - 264 * k) / 2 - 53 * k
        im = Image.new("RGBA", (W, W), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        for x0, y0, x1, y1 in ((0, 53, 88, 317), (193, 53, 712, 141), (193, 229, 712, 317), (818, 53, 906, 317)):
            d.rectangle([x0 * k, y0 * k + oy, x1 * k, y1 * k + oy], fill=(0, 92, 171, 255))
        return im
    except Exception:
        return None


try:
    st.set_page_config(
        page_title="HWACHEON - 해외출장비 정산 자동화 시스템",
        page_icon=_make_favicon(),
        layout="wide",
        initial_sidebar_state="expanded",
    )
except Exception:
    pass

# ----------------------------------------------------
# 2. 화천(HWACHEON) 홈페이지 스타일 커스텀 CSS
#    브랜드 컬러: PANTONE 2728C / C96 M69 Y0 K0 / R0 G92 B171 (#005CAB)
#    홈페이지 특징: 흰 배경 + 연회색(#F5F6F8) 라운드 카드/입력창 + 브랜드 블루 포인트 + Pretendard 폰트
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
        letter-spacing: -0.01em;
    }

    /* 메인 배경 및 레이아웃 */
    .stApp { background-color: #ffffff; color: var(--hw-text); }
    header[data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0.92);
        border-bottom: 1px solid var(--hw-line);
    }
    .block-container {
        max-width: 1280px !important;
        padding-top: 3.5rem;
        padding-bottom: 3rem;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
    }

    /* ---------- 사이드바 (항상 열림 고정) ---------- */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid var(--hw-line);
        transform: none !important;
        margin-left: 0 !important;
        visibility: visible !important;
        min-width: 290px !important;
        max-width: 290px !important;
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

    /* 사이드바 상단 브랜드 영역 (로고 + 시스템 설명 문구) */
    .hw-side-brand {
        padding: 18px 8px 22px 8px; text-align: center;
        border-bottom: 1px solid var(--hw-line);
        margin-bottom: 18px;
    }
    .hw-side-logo { display: flex; align-items: center; justify-content: center; gap: 12px; }
    .hw-side-logo svg { width: 100%; max-width: 168px; height: auto; margin: 0 auto; }
    .hw-side-logo-text {
        font-size: 20px; font-weight: 800; letter-spacing: 1.2px;
        color: var(--hw-blue); line-height: 1;
    }
    .hw-side-desc {
        font-size: 12.5px; color: var(--hw-text-sub); line-height: 1.6; margin-top: 14px; text-align: center;
    }
    .hw-side-prog {
        font-size: 21px; font-weight: 800; letter-spacing: -0.5px; color: var(--hw-text);
        padding: 22px 8px 8px 8px; margin-top: 12px; border-top: 1px solid var(--hw-line);
    }
    .hw-side-prog-on { color: var(--hw-blue); }
    .hw-side-caption {
        font-size: 12px; font-weight: 700; letter-spacing: 1.2px;
        color: var(--hw-text-sub); padding: 14px 8px 8px 8px;
    }

    /* 사이드바 메뉴 (버튼 → 홈페이지식 메뉴 리스트) */
    section[data-testid="stSidebar"] .stButton > button {
        width: 100%;
        justify-content: flex-start;
        text-align: left;
        background: transparent;
        color: var(--hw-text);
        border: none;
        border-radius: 12px;
        padding: 13px 18px;
        font-size: 15.5px;
        font-weight: 600;
        box-shadow: none;
    }
    section[data-testid="stSidebar"] .stButton > button > div { width: 100%; justify-content: flex-start; }
    section[data-testid="stSidebar"] .stButton > button p {
        text-align: left; width: 100%; margin: 0; color: inherit; font-weight: inherit;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: var(--hw-gray-bg); color: var(--hw-text);
    }
    section[data-testid="stSidebar"] .stButton > button[kind="primary"],
    section[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"] {
        background: var(--hw-blue); color: #ffffff;
    }
    section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover,
    section[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"]:hover {
        background: var(--hw-blue-dark); color: #ffffff;
    }
    section[data-testid="stSidebar"] [class*="st-key-sub_"] button {
        padding-left: 34px; font-size: 14.5px; font-weight: 500;
    }

    /* ---------- 페이지 헤드 (경로 + 제목) ---------- */
    .hw-page-head {
        padding-bottom: 20px;
        border-bottom: 1px solid var(--hw-line);
        margin-bottom: 28px;
    }
    .hw-crumb { font-size: 13px; color: var(--hw-text-sub); margin-bottom: 14px; }
    .hw-crumb b { color: var(--hw-blue); font-weight: 700; }
    .hw-crumb i { font-style: normal; margin: 0 8px; color: #B6BCC6; }
    .hw-page-eyebrow {
        font-size: 13px; font-weight: 700; letter-spacing: 1.2px;
        color: var(--hw-blue); margin-bottom: 6px;
    }
    .hw-page-title { font-size: 32px; font-weight: 800; color: var(--hw-text); letter-spacing: -0.6px; }
    .hw-page-sub { color: var(--hw-text-sub); font-size: 14.5px; margin-top: 8px; }

    /* ---------- 홈 (메인) ---------- */
    .hw-hero { padding: 8px 0 34px 0; }
    .hw-hero-eyebrow { font-size: 13px; font-weight: 700; letter-spacing: 2px; color: var(--hw-blue); margin-bottom: 14px; }
    .hw-hero-title { font-size: 44px; font-weight: 800; line-height: 1.25; letter-spacing: -1px; color: var(--hw-text); }
    .hw-hero-sub { font-size: 16px; color: var(--hw-text-sub); margin-top: 16px; line-height: 1.7; }

    .hw-prog-card {
        background: var(--hw-gray-bg);
        border-radius: 28px 28px 0 0;
        padding: 44px 40px 36px 40px;
        min-height: 300px;
        margin-bottom: -1rem;
    }
    .hw-prog-ico {
        width: 76px; height: 76px; border-radius: 22px; background: #ffffff;
        display: flex; align-items: center; justify-content: center; margin-bottom: 38px;
    }
    .hw-prog-ico svg { width: 38px; height: 38px; }
    .hw-prog-title { font-size: 30px; font-weight: 800; letter-spacing: -0.6px; color: var(--hw-text); margin-bottom: 14px; }
    .hw-prog-desc { font-size: 15.5px; color: #4B5563; line-height: 1.7; }
    .st-key-home_calc_btn button, .st-key-home_manage_btn button {
        border-radius: 0 0 28px 28px !important;
        padding: 20px 0 !important;
        font-size: 16px !important;
    }

    /* 제목 */
    .stApp h3 { font-weight: 800; letter-spacing: -0.3px; color: var(--hw-text); }

    /* 카드 스타일 섹션 */
    .hw-card { background: var(--hw-gray-bg); padding: 28px; border-radius: 20px; margin-bottom: 20px; }
    .hw-filter-title { font-size: 17px; font-weight: 800; margin: 4px 0 12px 0; }

    /* 하이라이트 박스 */
    .row-highlight-yellow {
        background-color: #FEE2E2; color: #991B1B; padding: 10px 14px;
        border-radius: 10px; font-weight: 600;
    }
    .row-highlight-blue {
        background-color: var(--hw-blue-soft); border: none; color: var(--hw-blue);
        padding: 12px 16px; border-radius: 10px; font-size: 16px;
    }

    /* 입력 필드 (홈페이지 검색창 스타일) */
    div[data-baseweb="input"],
    div[data-baseweb="base-input"],
    div[data-baseweb="select"] > div,
    div[data-testid="stDateInput"] div[data-baseweb="input"],
    div[data-baseweb="textarea"] {
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
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        background-color: var(--hw-blue);
        color: #ffffff;
        border-radius: 10px;
        border: none;
        font-weight: 600;
        padding: 10px 22px;
        transition: all 0.2s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
        background-color: var(--hw-blue-dark); color: #ffffff; border: none;
    }
    .stButton > button:focus:not(:active), .stDownloadButton > button:focus:not(:active) { color: #ffffff; border: none; }

    /* 탭 (홈페이지 하단 알약형 메뉴 스타일) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px; background-color: #ffffff; padding: 8px; border-radius: 14px;
        border: 1px solid var(--hw-line); box-shadow: 0 4px 16px rgba(0, 0, 0, 0.05);
        width: fit-content;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px; border-radius: 10px; padding: 0 22px; font-weight: 600;
        color: #374151; background-color: transparent;
    }
    .stTabs [aria-selected="true"] { background-color: var(--hw-blue) !important; color: #ffffff !important; }
    .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }

    /* 체크박스 포인트 컬러 */
    label[data-baseweb="checkbox"] > span:first-child[data-checked="true"],
    div[data-baseweb="checkbox"] > div[aria-checked="true"] { background-color: var(--hw-blue) !important; }

    /* 알림 박스 / 확장 패널 / 표 */
    div[data-testid="stAlert"] { border-radius: 12px; }
    div[data-testid="stExpander"] { border: 1px solid var(--hw-line); border-radius: 14px; background: #ffffff; }
    div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid var(--hw-line); }

    /* 메트릭 카드 */
    div[data-testid="stMetric"] { background-color: var(--hw-gray-bg); border-radius: 20px; padding: 22px 26px; }
    div[data-testid="stMetricLabel"] p { color: var(--hw-text-sub); font-weight: 600; }
    div[data-testid="stMetricValue"] { color: var(--hw-blue); font-weight: 800; font-size: 26px; }
    div[data-testid="stMetricValue"] > div { overflow: visible; text-overflow: clip; white-space: nowrap; }

    /* 구분선 / 링크 / 푸터 */
    .stApp hr { border-color: var(--hw-line); }
    .stApp a { color: var(--hw-blue); }
    .hw-footer {
        margin-top: 64px; padding: 26px 0 8px 0; border-top: 1px solid var(--hw-line);
        font-size: 12.5px; color: var(--hw-text-sub); line-height: 1.7;
    }
    .hw-footer b { color: var(--hw-text); }
    </style>
""",
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 3. 사이드바 (로고 + 프로그램/메뉴 내비게이션)
# ----------------------------------------------------
MENU_1 = "1. 출장 정보 입력"
MENU_2 = "2. 자금팀 연결 자료 생성"
MENU_3 = "3. 출장비 산정 내역서 생성"

MG_1 = "1. 출장 내역 등록"
MG_2 = "2. 출장 내역 조회 및 관리"

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
    + "</div>"
    '<div class="hw-side-desc">해외출장 업무 통합 시스템</div>'
    "</div>"
)

# 화면 상태: app_mode = home(메인) / calc(출장비 계산) / manage(해외출장 내역관리)
if "app_mode" not in st.session_state:
    st.session_state.app_mode = "home"
if "calc_menu" not in st.session_state:
    st.session_state.calc_menu = MENU_1
if "manage_menu" not in st.session_state:
    st.session_state.manage_menu = MG_1

if st.session_state.calc_menu not in (MENU_1, MENU_2, MENU_3):
    st.session_state.calc_menu = MENU_1
if st.session_state.manage_menu not in (MG_1, MG_2):
    st.session_state.manage_menu = MG_1


def _go(mode, sub=None):
    """버튼 클릭 시(화면 갱신 전에) 이동할 프로그램/메뉴를 저장"""
    st.session_state.app_mode = mode
    if sub is not None:
        st.session_state[f"{mode}_menu"] = sub


def _nav_button(label, key, active, mode, sub=None):
    st.button(
        label,
        key=key,
        type="primary" if active else "secondary",
        use_container_width=True,
        on_click=_go,
        args=(mode, sub),
    )


mode = st.session_state.app_mode

def _prog_heading(label, active):
    """프로그램 이름 (크게 강조) - 하위 메뉴는 항상 펼쳐서 표시"""
    cls = "hw-side-prog hw-side-prog-on" if active else "hw-side-prog"
    st.markdown(f'<div class="{cls}">{label}</div>', unsafe_allow_html=True)


with st.sidebar:
    st.markdown(SIDEBAR_BRAND_HTML, unsafe_allow_html=True)
    _nav_button("홈", "nav_home", mode == "home", "home")

    _prog_heading("해외출장비 계산", mode == "calc")
    for i, label in enumerate([MENU_1, MENU_2, MENU_3], start=1):
        _nav_button(label, f"sub_calc_{i}", mode == "calc" and st.session_state.calc_menu == label, "calc", label)

    _prog_heading("해외출장 내역관리", mode == "manage")
    for i, label in enumerate([MG_1, MG_2], start=1):
        _nav_button(label, f"sub_manage_{i}", mode == "manage" and st.session_state.manage_menu == label, "manage", label)

mode = st.session_state.app_mode
menu = st.session_state.calc_menu if mode == "calc" else None
mg_menu = st.session_state.manage_menu if mode == "manage" else None

# 페이지 상단 헤드 (경로 · 제목)
PAGE_HEAD = {
    MENU_1: ("해외출장비 계산", "STEP 01", "출장 정보 입력", "출장자 정보와 경비 항목을 입력하고 등록합니다."),
    MENU_2: ("해외출장비 계산", "STEP 02", "자금팀 연결 자료 생성", "등록된 출장 내역을 자금팀 제출용으로 집계합니다."),
    MENU_3: ("해외출장비 계산", "STEP 03", "출장비 산정 내역서 생성", "출장자별 해외출장비 산정 내역서를 확인합니다."),
    MG_1: ("해외출장 내역관리", "MENU 01", "출장 내역 등록", "등록 대기 중인 출장 내역을 확정 등록하여 DB에 기록합니다."),
    MG_2: ("해외출장 내역관리", "MENU 02", "출장 내역 조회 및 관리", "확정된 출장 내역을 검색·집계하고, 내역을 더블클릭하면 출장비 계산 당시 입력 내용을 확인합니다."),
}
_head_key = menu if mode == "calc" else mg_menu
if _head_key in PAGE_HEAD:
    _crumb, _eyebrow, _title, _sub = PAGE_HEAD[_head_key]
    st.markdown(
        '<div class="hw-page-head">'
        f'<div class="hw-crumb">HOME<i>&rsaquo;</i>{_crumb}<i>&rsaquo;</i><b>{_title}</b></div>'
        f'<div class="hw-page-eyebrow">{_eyebrow}</div>'
        f'<div class="hw-page-title">{_title}</div>'
        f'<div class="hw-page-sub">{_sub}</div>'
        "</div>",
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
            "출장목적": d.get("출장목적구분"),
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
            vals = [r2 - 1, rec["출장자성명"], rec["출장지"], cat, name, amt, payer]
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
    fx_txt = f"{float(rec['환율']):,.2f}"

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

    # 1-2) 적용환율: '적용환율' 라벨 오른쪽 칸에 1번 메뉴에서 입력한 환율 입력
    fx_label = _find_cell(ws, "적용환율")
    if fx_label is not None:
        _put(ws, fx_label.row, fx_label.column + 1, float(rec["환율"]), "#,##0.00")

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
                basis, calc = f"{std_hotel:,}{cur} / 박", f"{std_hotel:,} {cur} * {nights}박 x {fx_txt}{unit}"
            if c_basis: _put(ws, r, c_basis, basis)
            if c_period: _put(ws, r, c_period, f"{nights}박")
            if c_amt: _put(ws, r, c_amt, hotel_amt, "#,##0")
            if c_calc: _put(ws, r, c_calc, calc)

        if r_daily is not None:
            r = r_daily.row
            if c_basis: _put(ws, r, c_basis, f"{std_daily:,}{cur} / 일")
            if c_period: _put(ws, r, c_period, f"{days}일")
            if c_amt: _put(ws, r, c_amt, daily_amt, "#,##0")
            if c_calc: _put(ws, r, c_calc, f"{std_daily:,} {cur} * {days}일 x {fx_txt}{unit}")

        if r_total is not None and c_amt and r_hotel is not None and r_daily is not None:
            L = get_column_letter(c_amt)
            _put(ws, r_total.row, c_amt, f"={L}{r_hotel.row}+{L}{r_daily.row}", "#,##0")
            try:
                amt_cell = ws.cell(row=r_total.row, column=c_amt)
                amt_cell.fill = copy(r_total.fill)  # '지급 총액' 라벨과 같은 셀 배경
                lf = r_total.font
                amt_cell.font = Font(name=lf.name, size=lf.sz, bold=True, color=lf.color)  # 굵은 글씨
            except Exception:
                pass

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
# 해외출장 내역관리 (확정된 출장 내역 기록 · 검색 · 관리)
# ----------------------------------------------------
# 확정 등록된 내역은 이 파이썬 파일과 같은 폴더의 'travel_records.db'(SQLite) 파일에 저장됩니다.
# 파일을 복사해 두면 그대로 백업/이전할 수 있습니다.
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "travel_records.db")

DEPARTMENT_LIST = [
    "임원", "경영지원본부", "경영지원실", "인사지원팀", "관리팀", "재무전략실",
    "노동조합", "재무팀", "자금팀", "정보실", "정보팀", "IBU", "성장전략실",
    "프로젝트팀", "구매전략본부", "HTB 대만지사", "구매팀", "VI팀", "품질혁신본부",
    "QM팀", "보전팀", "생산본부", "생산관리팀", "생산기술팀", "가공팀", "F/S가공",
    "정밀가공", "가공지원", "UNIT팀", "UNIT준비", "UNIT조립", "UNIT서비스",
    "생산1팀", "생산2팀", "서비스센터", "서비스1팀", "서비스2팀", "서비스3팀",
    "서비스4팀", "기술개발연구소", "MC개발팀", "TC개발팀", "5축개발팀", "UNIT개발팀",
    "제어개발팀", "제어SW개발팀", "가공기술1팀", "가공기술2팀", "소재사업부문", "기타",
]

POSITION_LIST = [
    "사장", "부사장", "전무", "상무", "이사", "부장", "차장",
    "과장", "대리", "계장", "사원", "1급기능장", "2급기능장",
]

PURPOSE_OPTIONS = [
    "고객사 영업·상담",
    "전시회·박람회 참관",
    "설비 설치·시운전",
    "A/S·기술지원",
    "해외법인·지사 업무",
    "구매·협력사 미팅",
    "교육·연수",
    "계약·협상",
    "기타",
]

DATE_MIN = datetime.date(2000, 1, 1)
DATE_MAX = datetime.date(2100, 12, 31)

DB_FIELDS = [
    "traveler", "dept", "position", "country", "start_date", "end_date", "nights", "days",
    "purpose_cat", "purpose_detail", "emp_amount", "agency_amount", "total_amount",
    "approved_date", "paid_date", "memo", "items_json", "calc_json",
]

# 화면/엑셀에 표시할 열 (DB 열 이름, 표시 이름)
VIEW_COLS = [
    ("traveler", "출장자"), ("dept", "부서"), ("position", "직급"), ("country", "출장지"),
    ("start_date", "출장시작일"), ("end_date", "출장종료일"), ("period", "출장기간"),
    ("purpose_cat", "출장목적"), ("purpose_detail", "목적상세"),
    ("emp_amount", "직원_지급액"), ("agency_amount", "여행사_지급액"), ("total_amount", "총지급액"),
    ("approved_date", "결재승인일"), ("paid_date", "지급일"), ("memo", "비고"),
]
MONEY_COLS = ["직원_지급액", "여행사_지급액", "총지급액"]


def db_connect():
    return sqlite3.connect(DB_PATH)


def db_init():
    with closing(db_connect()) as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS trips (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                traveler TEXT NOT NULL, dept TEXT, position TEXT, country TEXT,
                start_date TEXT, end_date TEXT, nights INTEGER, days INTEGER,
                purpose_cat TEXT, purpose_detail TEXT,
                emp_amount INTEGER DEFAULT 0, agency_amount INTEGER DEFAULT 0, total_amount INTEGER DEFAULT 0,
                approved_date TEXT, paid_date TEXT, memo TEXT, items_json TEXT, created_at TEXT,
                calc_json TEXT
            )"""
        )
        cols = [r[1] for r in conn.execute("PRAGMA table_info(trips)").fetchall()]
        if "calc_json" not in cols:  # 이전 버전 DB 파일 보완
            conn.execute("ALTER TABLE trips ADD COLUMN calc_json TEXT")
        conn.commit()


def db_insert(rec):
    cols = DB_FIELDS + ["created_at"]
    vals = [rec.get(c) for c in DB_FIELDS] + [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    with closing(db_connect()) as conn:
        conn.execute(
            f"INSERT INTO trips ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", vals
        )
        conn.commit()


def db_update(rid, rec):
    sets = ",".join(f"{c}=?" for c in DB_FIELDS if c in rec)
    vals = [rec[c] for c in DB_FIELDS if c in rec] + [rid]
    with closing(db_connect()) as conn:
        conn.execute(f"UPDATE trips SET {sets} WHERE id=?", vals)
        conn.commit()


def db_delete(rid):
    with closing(db_connect()) as conn:
        conn.execute("DELETE FROM trips WHERE id=?", (rid,))
        conn.commit()


def db_exists(traveler, start_date, end_date, country):
    """같은 출장자·기간·출장지의 확정 내역이 이미 있는지 확인 (중복 등록 방지)"""
    with closing(db_connect()) as conn:
        row = conn.execute(
            "SELECT id FROM trips WHERE traveler=? AND start_date=? AND end_date=? AND country=?",
            (traveler, start_date, end_date, country),
        ).fetchone()
    return row[0] if row else None


def db_load():
    db_init()
    with closing(db_connect()) as conn:
        return pd.read_sql_query("SELECT * FROM trips ORDER BY start_date DESC, id DESC", conn)


def trip_from_calc(rec, approved, paid, memo):
    """해외출장비 계산의 출장 1건 → 확정 등록용 레코드"""
    row = process_travel_data([rec]).iloc[0]
    items, _, _, _ = calc_item_rows(rec)
    return {
        "traveler": rec["출장자성명"],
        "dept": rec["부서"],
        "position": rec["직급"],
        "country": rec["출장지"],
        "start_date": rec["출장시작일"],
        "end_date": rec["출장종료일"],
        "nights": int(rec["출장박수"]),
        "days": int(rec["출장일수"]),
        "purpose_cat": rec.get("출장목적구분") or "기타",
        "purpose_detail": rec.get("출장목적상세", ""),
        "emp_amount": int(row["직원_지급액"]),
        "agency_amount": int(row["여행사_지급액"]),
        "total_amount": int(row["총출장비"]),
        "approved_date": str(approved),
        "paid_date": str(paid) if paid else "",
        "memo": memo,
        "items_json": json.dumps(
            [{"구분": c, "항목": n, "금액": a, "지급처": p} for c, n, a, p in items], ensure_ascii=False
        ),
        # 출장비 계산 당시 입력한 데이터 전체 (더블클릭 상세 조회용)
        "calc_json": json.dumps(rec, ensure_ascii=False, default=str),
    }


def trips_view(df):
    """DB 조회 결과 → 화면/엑셀 표시용 표 (순번 1부터)"""
    v = df.copy()
    v["period"] = [f"{int(n)}박 {int(d)}일" for n, d in zip(v["nights"], v["days"])]
    v = v[[c for c, _ in VIEW_COLS]].rename(columns=dict(VIEW_COLS))
    v = v.fillna("")
    v.index = range(1, len(v) + 1)
    v.index.name = "순번"
    return v


def style_money(v, cols=None):
    cols = cols or MONEY_COLS
    return v.style.format({c: "{:,.0f}" for c in cols if c in v.columns})


def build_trips_excel(view, cond_text):
    """조회 결과 엑셀 (표 + 합계)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "해외출장 내역"
    heads = ["순번"] + list(view.columns)
    ncol = len(heads)
    ws["A1"] = "해외출장 내역 조회 결과"
    ws["A1"].font = Font(bold=True, size=16, color="005CAB")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws["A2"] = f"조회 조건: {cond_text}  |  출력일: {datetime.date.today():%Y-%m-%d}"
    ws["A2"].alignment = XL_LEFT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)

    hr = 4
    _xl_header(ws, hr, heads)
    money_idx = [i for i, h in enumerate(heads, start=1) if h in MONEY_COLS]
    for r, (no, row) in enumerate(view.iterrows(), start=hr + 1):
        vals = [no] + list(row.values)
        for j, v in enumerate(vals, start=1):
            if hasattr(v, "item"):
                v = v.item()
            c = ws.cell(row=r, column=j, value=v)
            c.border = XL_BORDER
            if j in money_idx:
                c.number_format = XL_NUM
                c.alignment = XL_RIGHT
            else:
                c.alignment = XL_CENTER

    first, last = hr + 1, hr + len(view)
    tot = last + 1
    ws.cell(row=tot, column=1, value="합 계")
    for j in range(1, ncol + 1):
        c = ws.cell(row=tot, column=j)
        c.fill = XL_SUB_FILL
        c.font = Font(bold=True)
        c.border = XL_BORDER
        c.alignment = XL_CENTER
    for j in money_idx:
        col = get_column_letter(j)
        c = ws.cell(row=tot, column=j, value=f"=SUM({col}{first}:{col}{last})")
        c.number_format = XL_NUM
        c.alignment = XL_RIGHT
    widths = {"순번": 7, "목적상세": 32, "비고": 24, "출장목적": 20, "부서": 16, "출장기간": 12}
    _xl_widths(ws, [widths.get(h, 14) for h in heads])
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    return _xl_bytes(wb)


def _flash():
    msg = st.session_state.pop("mg_flash", None)
    if msg:
        st.success(msg)


def _money_input(label, value=0, key=None):
    return st.number_input(label, min_value=0, value=int(value), step=1000, key=key)


def _group_table(f, col, label):
    g = (
        f.groupby(col)
        .agg(건수=("id", "count"), 직원_지급액=("emp_amount", "sum"), 여행사_지급액=("agency_amount", "sum"), 총지급액=("total_amount", "sum"))
        .sort_values("총지급액", ascending=False)
        .reset_index()
        .rename(columns={col: label})
    )
    g.index = range(1, len(g) + 1)
    g.index.name = "순번"
    return g


# ---------------- 출장 내역 조회 및 관리 ----------------
def _selected_rows(sel):
    if sel is None:
        return []
    if isinstance(sel, pd.DataFrame):
        return sel.to_dict("records")
    return list(sel)


def render_trip_detail(row):
    """DB에 저장된 출장 1건의 상세 (출장비 계산 당시 입력 데이터)"""
    st.markdown('<div class="hw-filter-title">출장 상세 내역 (출장비 계산 당시 입력 데이터)</div>', unsafe_allow_html=True)
    rec = None
    try:
        if row.get("calc_json"):
            rec = json.loads(row["calc_json"])
    except Exception:
        rec = None

    info = {
        "출장자": row["traveler"], "부서": row["dept"], "직급": row["position"], "출장지": row["country"],
        "출장 기간": f'{row["start_date"]} ~ {row["end_date"]} ({int(row["nights"])}박 {int(row["days"])}일)',
        "출장 목적": f'{row["purpose_cat"] or "-"}' + (f' / {row["purpose_detail"]}' if row["purpose_detail"] else ""),
        "결재승인일": row["approved_date"] or "-", "지급일": row["paid_date"] or "-", "비고": row["memo"] or "-",
    }
    if rec:
        info["직급 구분"] = rec.get("직급구분", "-")
        info["지역 구분"] = rec.get("지역구분", "-")
        info["적용 환율"] = f'{float(rec.get("환율", 0)):,.2f}' + (" (100엔 기준)" if rec.get("지역구분") == "특" else "")
    keys = list(info.keys())
    for i in range(0, len(keys), 3):
        cols = st.columns(3)
        for c, k in zip(cols, keys[i:i + 3]):
            with c:
                st.markdown(f"<div style='font-size:12px;color:#6B7280'>{k}</div><div style='font-weight:700;margin-bottom:10px'>{info[k]}</div>", unsafe_allow_html=True)

    if rec:
        items, _, _, _ = calc_item_rows(rec)
        if items:
            t = pd.DataFrame(items, columns=["구분", "항목", "금액(원)", "지급처"])
            t.index = range(1, len(t) + 1)
            t.index.name = "순번"
            st.dataframe(style_money(t, ["금액(원)"]), use_container_width=True)
    else:
        try:
            its = json.loads(row["items_json"]) if row["items_json"] else []
        except Exception:
            its = []
        if its:
            t = pd.DataFrame(its)
            t.index = range(1, len(t) + 1)
            t.index.name = "순번"
            st.dataframe(style_money(t, ["금액"]), use_container_width=True)
        else:
            st.caption("직접 입력으로 등록된 내역이라 항목별 상세 데이터가 없습니다.")
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("직원 지급액", f"{int(row['emp_amount']):,} 원")
    with m2:
        st.metric("여행사 지급액", f"{int(row['agency_amount']):,} 원")
    with m3:
        st.metric("총 지급액", f"{int(row['total_amount']):,} 원")


def render_manage_search():
    _flash()
    df = db_load()
    if df.empty:
        st.info("확정 등록된 출장 내역이 없습니다. '1. 출장 내역 등록' 메뉴에서 먼저 등록해 주세요.")
        return
    df = df.sort_values(["start_date", "id"], ascending=[False, False]).reset_index(drop=True)

    opts = {
        "mg_f_dept": sorted(df["dept"].dropna().unique()),
        "mg_f_purpose": sorted(df["purpose_cat"].dropna().unique()),
        "mg_f_name": sorted(df["traveler"].dropna().unique()),
        "mg_f_country": sorted(df["country"].dropna().unique()),
    }
    # 날짜는 기본적으로 비워 둠 / 목록에서 사라진 선택값 정리
    st.session_state.setdefault("mg_f_from", None)
    st.session_state.setdefault("mg_f_to", None)
    st.session_state.setdefault("mg_kw", "")
    for k, o in opts.items():
        st.session_state[k] = [v for v in st.session_state.get(k, []) if v in o]

    def _reset_all():
        st.session_state.mg_f_from = None
        st.session_state.mg_f_to = None
        st.session_state.mg_kw = ""
        for k in opts:
            st.session_state[k] = []
        st.session_state.mg_applied = None
        st.session_state.mg_detail_id = None

    st.markdown('<div class="hw-filter-title">검색 조건</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        f_from = st.date_input("출장 기간 (시작)", value=None, min_value=DATE_MIN, max_value=DATE_MAX, key="mg_f_from")
    with c2:
        f_to = st.date_input("출장 기간 (종료)", value=None, min_value=DATE_MIN, max_value=DATE_MAX, key="mg_f_to")
    with c3:
        kw = st.text_input("키워드 검색", key="mg_kw")

    c4, c5, c6, c7 = st.columns(4)
    with c4:
        f_dept = st.multiselect("부서", opts["mg_f_dept"], key="mg_f_dept", placeholder=" ")
    with c5:
        f_purpose = st.multiselect("출장 목적", opts["mg_f_purpose"], key="mg_f_purpose", placeholder=" ")
    with c6:
        f_name = st.multiselect("출장자", opts["mg_f_name"], key="mg_f_name", placeholder=" ")
    with c7:
        f_country = st.multiselect("출장지", opts["mg_f_country"], key="mg_f_country", placeholder=" ")
    st.caption("조건을 설정하지 않으면 전체 내역이 출장 시작일 최신순으로 표시됩니다. 조건 설정 후 '검색' 버튼을 눌러야 조건이 적용됩니다. 출장 기간은 일정이 하루라도 겹치는 내역을 조회합니다.")

    bc1, bc2, _sp = st.columns([1, 1, 4])
    with bc1:
        do_search = st.button("검색", key="mg_search_btn", type="primary", use_container_width=True)
    with bc2:
        st.button("초기화", key="mg_reset_btn", on_click=_reset_all, use_container_width=True)

    if do_search:
        st.session_state.mg_detail_id = None
        if f_from and f_to and f_from > f_to:
            st.session_state.mg_applied = None
            st.error("출장 기간의 시작일이 종료일보다 늦습니다.")
            return
        st.session_state.mg_applied = {
            "from": f_from, "to": f_to, "kw": kw,
            "dept": list(f_dept), "purpose": list(f_purpose), "name": list(f_name), "country": list(f_country),
        }
    ap = st.session_state.get("mg_applied") or {"from": None, "to": None, "kw": "", "dept": [], "purpose": [], "name": [], "country": []}
    f_from, f_to, kw = ap["from"], ap["to"], ap["kw"]
    f_dept, f_purpose, f_name, f_country = ap["dept"], ap["purpose"], ap["name"], ap["country"]

    f = df.copy()
    sd = pd.to_datetime(f["start_date"])
    ed = pd.to_datetime(f["end_date"])
    if f_to:
        f = f[sd <= pd.Timestamp(f_to)]
        ed = ed[f.index]
    if f_from:
        f = f[ed.loc[f.index] >= pd.Timestamp(f_from)]
    if f_dept:
        f = f[f["dept"].isin(f_dept)]
    if f_purpose:
        f = f[f["purpose_cat"].isin(f_purpose)]
    if f_name:
        f = f[f["traveler"].isin(f_name)]
    if f_country:
        f = f[f["country"].isin(f_country)]
    if kw.strip():
        hay = (
            f[["traveler", "dept", "position", "country", "purpose_cat", "purpose_detail", "memo"]]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
        )
        f = f[hay.str.contains(kw.strip(), case=False, regex=False)]

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("조회 건수", f"{len(f):,} 건")
    with m2:
        st.metric("총 지급액", f"{int(f['total_amount'].sum()):,} 원")
    with m3:
        st.metric("직원 지급액", f"{int(f['emp_amount'].sum()):,} 원")
    with m4:
        st.metric("여행사 지급액", f"{int(f['agency_amount'].sum()):,} 원")

    if f.empty:
        st.info("조건에 맞는 출장 내역이 없습니다.")
        return

    view = trips_view(f)
    grid_df = view.reset_index()
    grid_df["id"] = f["id"].astype(int).values
    st.caption("내역을 더블클릭하면 아래에 출장비 계산 당시 입력했던 데이터가 표시됩니다.")
    gb = GridOptionsBuilder.from_dataframe(grid_df)
    gb.configure_selection(selection_mode="single", use_checkbox=False, suppressRowClickSelection=True)
    gb.configure_default_column(minWidth=110, resizable=True)
    gb.configure_column("순번", minWidth=90, maxWidth=100)
    gb.configure_column("id", hide=True)
    num_fmt = JsCode(
        "function(p){ if (p.value === null || p.value === undefined || p.value === '') { return ''; }"
        " var n = Number(p.value); return isNaN(n) ? p.value : n.toLocaleString('ko-KR'); }"
    )
    gb.configure_columns(MONEY_COLS, valueFormatter=num_fmt, cellStyle={"textAlign": "right"}, minWidth=130)
    go = gb.build()
    go["onRowDoubleClicked"] = JsCode("function(e){ e.api.deselectAll(); e.node.setSelected(true); }")
    resp = AgGrid(
        grid_df,
        gridOptions=go,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        fit_columns_on_grid_load=True,
        height=380,
        theme="alpine",
        allow_unsafe_jscode=True,
        key="mg_grid",
    )
    picked = _selected_rows(resp.get("selected_rows") if hasattr(resp, "get") else resp["selected_rows"])
    if picked:
        st.session_state.mg_detail_id = int(picked[0]["id"])
    did = st.session_state.get("mg_detail_id")
    if did is not None and did in set(f["id"].astype(int)):
        st.markdown("---")
        render_trip_detail(f[f["id"] == did].iloc[0])
        st.button("상세 닫기", key="mg_detail_close", on_click=lambda: st.session_state.update(mg_detail_id=None))

    cond = f"{f_from or '전체'} ~ {f_to or '전체'}"
    for label, vals in (("부서", f_dept), ("목적", f_purpose), ("출장자", f_name), ("출장지", f_country)):
        if vals:
            cond += f" / {label}: {', '.join(vals)}"
    if kw.strip():
        cond += f" / 키워드: {kw.strip()}"
    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        "조회 결과 엑셀 다운로드",
        data=build_trips_excel(view, cond),
        file_name=f"해외출장내역_{datetime.date.today():%Y%m%d}.xlsx",
        mime=XL_MIME,
        use_container_width=True,
    )

    st.markdown("### 조회 결과 집계")
    t1, t2, t3, t4 = st.tabs(["출장 목적별", "부서별", "출장자별", "출장지별"])
    for tab, (col, label) in zip((t1, t2, t3, t4), (("purpose_cat", "출장 목적"), ("dept", "부서"), ("traveler", "출장자"), ("country", "출장지"))):
        with tab:
            g = _group_table(f, col, label)
            st.dataframe(style_money(g, ["직원_지급액", "여행사_지급액", "총지급액"]), use_container_width=True)


# ---------------- 확정 등록 ----------------
def render_manage_register():
    _flash()
    st.markdown("### 계산 결과 확정 등록")
    st.caption("해외출장비 계산에서 입력한 출장 중 아직 확정 등록하지 않은(등록 대기) 건만 표시됩니다. 확정 등록하면 '2. 출장 내역 조회 및 관리'에서 조회·관리됩니다.")

    recs = st.session_state.travel_list
    pending = [i for i, r in enumerate(recs) if not db_exists(r["출장자성명"], r["출장시작일"], r["출장종료일"], r["출장지"])]
    if not recs:
        st.info("계산된 출장 내역이 없습니다. '해외출장비 계산'에서 출장 정보를 먼저 등록해 주세요. (과거 내역은 아래 '직접 입력으로 등록'을 이용하세요.)")
    elif not pending:
        st.info("등록 대기 중인 출장 내역이 없습니다. 확정 등록된 내역은 '2. 출장 내역 조회 및 관리'에서 확인할 수 있습니다.")
    else:
        dfp = process_travel_data(recs)
        today = datetime.date.today()
        rows = []
        for i in pending:
            r = recs[i]
            overdue = datetime.date.fromisoformat(r["출장시작일"]) <= today
            rows.append(
                {
                    "출장자": r["출장자성명"],
                    "부서": r["부서"],
                    "출장지": r["출장지"],
                    "출장기간": f'{r["출장시작일"]} ~ {r["출장종료일"]}',
                    "출장목적": r.get("출장목적구분") or "-",
                    "총 출장비": int(dfp.loc[i, "총출장비"]),
                    "상태": "등록 지연(출발일 경과)" if overdue else "등록 대기",
                }
            )
        tbl = pd.DataFrame(rows)
        tbl.index = range(1, len(tbl) + 1)
        tbl.index.name = "순번"
        st.dataframe(style_money(tbl, ["총 출장비"]), use_container_width=True)

        sel = st.multiselect(
            "확정 등록할 출장 건",
            options=pending,
            default=pending,
            format_func=lambda i: f"{i + 1}. {recs[i]['출장자성명']} ({recs[i]['출장지']}, {recs[i]['출장시작일']})",
        )
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            approved = st.date_input("결재 승인일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX, key="mg_reg_approved")
        with c2:
            use_paid = st.checkbox("지급일 입력", value=False, key="mg_reg_use_paid")
        with c3:
            paid = st.date_input("지급일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX, key="mg_reg_paid", disabled=not use_paid)
        memo = st.text_input("비고 (선택)", key="mg_reg_memo")

        reg_clicked = st.button("선택한 출장 건 확정 등록", use_container_width=True, key="mg_reg_btn")
        if reg_clicked and not sel:
            st.warning("확정 등록할 출장 건을 선택해 주세요.")
        elif reg_clicked:
            ok, skipped = 0, 0
            for i in sel:
                r = recs[i]
                if db_exists(r["출장자성명"], r["출장시작일"], r["출장종료일"], r["출장지"]):
                    skipped += 1
                    continue
                db_insert(trip_from_calc(r, approved, paid if use_paid else None, memo))
                ok += 1
            msg = f"{ok}건을 확정 등록했습니다."
            if skipped:
                msg += f" (이미 등록된 {skipped}건은 제외)"
            st.session_state.mg_flash = msg
            st.rerun()

    with st.expander("직접 입력으로 등록 (해외출장비 계산을 거치지 않은 과거 내역)"):
        with st.form("mg_manual_form", clear_on_submit=True):
            a1, a2, a3 = st.columns(3)
            with a1:
                m_name = st.text_input("출장자 성명")
            with a2:
                m_dept = st.selectbox("부서", DEPARTMENT_LIST)
            with a3:
                m_pos = st.selectbox("직급", POSITION_LIST, index=POSITION_LIST.index("사원"))
            b1, b2, b3 = st.columns(3)
            with b1:
                m_country = st.text_input("출장지(국가)")
            with b2:
                m_start = st.date_input("출장 시작일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX)
            with b3:
                m_end = st.date_input("출장 종료일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX)
            p1, p2 = st.columns(2)
            with p1:
                m_pcat = st.selectbox("출장 목적 구분", PURPOSE_OPTIONS)
            with p2:
                m_pdetail = st.text_input("출장 목적 상세")
            d1, d2 = st.columns(2)
            with d1:
                m_emp = _money_input("직원 지급액(원)")
            with d2:
                m_agency = _money_input("여행사 지급액(원)")
            e1, e2 = st.columns(2)
            with e1:
                m_approved = st.date_input("결재 승인일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX)
            with e2:
                m_memo = st.text_input("비고")
            m_use_paid = st.checkbox("지급일 입력")
            m_paid = st.date_input("지급일", value=datetime.date.today(), min_value=DATE_MIN, max_value=DATE_MAX)
            submitted = st.form_submit_button("직접 등록")

        if submitted:
            if not m_name.strip() or not m_country.strip():
                st.error("출장자 성명과 출장지는 필수 입력 항목입니다.")
            elif m_end < m_start:
                st.error("출장 종료일이 시작일보다 빠릅니다.")
            elif db_exists(m_name.strip(), str(m_start), str(m_end), m_country.strip()):
                st.warning("같은 출장자·기간·출장지의 내역이 이미 등록되어 있습니다.")
            else:
                days = (m_end - m_start).days + 1
                db_insert(
                    {
                        "traveler": m_name.strip(), "dept": m_dept, "position": m_pos, "country": m_country.strip(),
                        "start_date": str(m_start), "end_date": str(m_end),
                        "nights": max(days - 1, 0), "days": days,
                        "purpose_cat": m_pcat, "purpose_detail": m_pdetail,
                        "emp_amount": int(m_emp), "agency_amount": int(m_agency), "total_amount": int(m_emp) + int(m_agency),
                        "approved_date": str(m_approved), "paid_date": str(m_paid) if m_use_paid else "",
                        "memo": m_memo, "items_json": "",
                    }
                )
                st.session_state.mg_flash = f"[{m_name.strip()}] 님의 출장 내역을 직접 등록했습니다."
                st.rerun()


def render_manage(page):
    try:
        db_init()
    except Exception as e:
        st.error(f"출장 내역 저장 파일(travel_records.db)을 열 수 없습니다: {e}")
        return
    if page == MG_1:
        render_manage_register()
    elif page == MG_2:
        render_manage_search()


# ---------------- 홈 (메인) ----------------
ICON_CALC = (
    '<svg viewBox="0 0 24 24" fill="none" stroke="#005CAB" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
    '<rect x="4" y="2" width="16" height="20" rx="2"/><line x1="8" y1="6" x2="16" y2="6"/>'
    '<line x1="16" y1="14" x2="16" y2="18"/>'
    '<path d="M16 10h.01M12 10h.01M8 10h.01M12 14h.01M8 14h.01M12 18h.01M8 18h.01"/></svg>'
)
ICON_MANAGE = (
    '<svg viewBox="0 0 24 24" fill="none" stroke="#005CAB" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>'
    '<line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>'
)


def _prog_card(icon, title, desc):
    return (
        '<div class="hw-prog-card">'
        f'<div class="hw-prog-ico">{icon}</div>'
        f'<div class="hw-prog-title">{title}</div>'
        f'<div class="hw-prog-desc">{desc}</div>'
        "</div>"
    )


def overdue_pending():
    """출발일 00시가 지났는데(출발일 <= 오늘) 아직 확정 등록하지 않은 출장 내역 목록"""
    try:
        db_init()
    except Exception:
        return []
    today = datetime.date.today()
    out = []
    for i, r in enumerate(st.session_state.travel_list):
        try:
            sd = datetime.date.fromisoformat(r["출장시작일"])
        except Exception:
            continue
        if sd <= today and not db_exists(r["출장자성명"], r["출장시작일"], r["출장종료일"], r["출장지"]):
            out.append((i, r, (today - sd).days))
    return sorted(out, key=lambda x: x[1]["출장시작일"])


def render_alerts():
    items = overdue_pending()
    _sp, bell = st.columns([5, 1.4])
    with bell:
        with st.popover(f"알림 {len(items)}" if items else "알림", icon=":material/notifications:", use_container_width=True):
            if not items:
                st.caption("확인할 알림이 없습니다.")
            else:
                st.markdown(f"**확정 등록이 필요한 출장 {len(items)}건**")
                st.caption("출발일이 지났지만 아직 확정 등록하지 않은 출장 내역입니다.")
                for _i, r, late in items:
                    st.markdown(
                        f"- **{r['출장자성명']}** · {r['출장지']} · 출발 {r['출장시작일']}"
                        + (f" (출발 {late}일 경과)" if late > 0 else " (오늘 출발)")
                    )
                st.button("출장 내역 등록으로 이동", key="alert_go_reg", use_container_width=True, on_click=_go, args=("manage", MG_1))


def render_home():
    render_alerts()
    st.markdown(
        '<div class="hw-hero">'
        '<div class="hw-hero-eyebrow">HWACHEON OVERSEAS BUSINESS TRIP</div>'
        '<div class="hw-hero-title">해외출장 업무 통합 시스템</div>'
        '<div class="hw-hero-sub">출장비 산정부터 자금팀 정산 자료 생성, 확정된 출장 내역의 기록·검색까지<br>한 곳에서 처리합니다.</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown(
            _prog_card(ICON_CALC, "해외출장비 계산", "출장 정보를 입력해 출장비를 산정하고, 자금팀 자료와 산정 내역서를 엑셀로 생성합니다."),
            unsafe_allow_html=True,
        )
        st.button("실행하기", key="home_calc_btn", use_container_width=True, on_click=_go, args=("calc",))
    with right:
        st.markdown(
            _prog_card(ICON_MANAGE, "해외출장 내역관리", "확정된 출장 내역을 기록하고, 목적·부서·출장자·기간별로 검색합니다."),
            unsafe_allow_html=True,
        )
        st.button("실행하기", key="home_manage_btn", use_container_width=True, on_click=_go, args=("manage",))


def render_footer():
    st.markdown(
        '<div class="hw-footer"><b>HWACHEON</b> 화천기공 해외출장 경비 산정 &amp; 자금팀 정산 자동화 시스템<br>'
        "본 시스템은 사내 업무용입니다. 출장비 산정 기준은 해외출장 지급규정을 따릅니다.</div>",
        unsafe_allow_html=True,
    )


# ----------------------------------------------------
# 4. 페이지 구성 (사이드바 메뉴 선택에 따라 표시)
# ----------------------------------------------------
if menu == MENU_1:
    st.markdown("### 출장 기본 정보 입력")
    if st.session_state.edit_target_index is not None:
        st.info(
            f"현재 **[인덱스 {st.session_state.edit_target_index}]** 번 출장 내역 수정 중입니다."
        )

    target_edit_data = None
    if st.session_state.edit_target_index is not None and len(
        st.session_state.travel_list
    ) > st.session_state.edit_target_index:
        target_edit_data = st.session_state.travel_list[
            st.session_state.edit_target_index
        ]

    department_list = DEPARTMENT_LIST

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

        position_list = POSITION_LIST
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
            "[서울외국환중개 환율 조회 링크](http://www.smbs.biz/ExRate/TodayExRate.jsp)"
        )
        if is_jpy:
            st.caption("위 링크에서 조회한 일본 엔화(JPY 100엔) 환율을 우측 칸에 입력해주세요.")
        else:
            st.caption("위 링크에서 조회한 미국 달러(USD) 환율을 우측 칸에 입력해주세요.")
    with col_rate_input:
        exchange_rate = st.number_input(
            rate_label,
            min_value=0.0,
            value=float(default_exchange_rate),
            step=1.0,
            format="%.2f",
            key=rate_key,
        )

    # 출장 목적 (확정 등록 및 해외출장 내역관리 검색에 사용)
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        _default_pc = (target_edit_data.get("출장목적구분") if target_edit_data else None) or PURPOSE_OPTIONS[0]
        purpose_cat = st.selectbox(
            "출장 목적 구분",
            PURPOSE_OPTIONS,
            index=PURPOSE_OPTIONS.index(_default_pc) if _default_pc in PURPOSE_OPTIONS else 0,
        )
    with col_p2:
        purpose_detail = st.text_input(
            "출장 목적 상세",
            value=(target_edit_data.get("출장목적상세", "") if target_edit_data else ""),
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

        st.info(f"최종 산정된 출장 기간: **{calculated_nights}박 {calculated_days}일**")

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
    st.markdown("### 금액 상세 입력")

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
        if st.button("기타 행 추가", use_container_width=True):
            st.session_state.other_rows.append({"item": "", "amount": 0, "payer": "여행사"})
            st.rerun()
    with b_c3:
        if len(st.session_state.other_rows) > 1:
            if st.button("기타 마지막 행 삭제", use_container_width=True):
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
    btn_label = "수정 사항 반영하기" if st.session_state.edit_target_index is not None else "입력 내역 저장 및 등록"
    submitted = st.button(btn_label, use_container_width=True)

    if submitted:
        if not name or not country:
            st.error("출장자 성명과 출장지는 필수 입력 항목입니다.")
        else:
            new_data = {
                "출장자성명": name,
                "부서": department,
                "직급": position,
                "직급구분": position_group,
                "출장지": country,
                "출장목적구분": purpose_cat,
                "출장목적상세": purpose_detail,
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
                st.success(f"[{name}] 님의 내역이 수정되었습니다.")
                st.session_state.edit_target_index = None
            else:
                st.session_state.travel_list.append(new_data)
                st.success(f"[{name}] 님의 내역이 등록되었습니다.")
            st.rerun()

    st.markdown("---")
    st.markdown("### 현재 등록된 출장 내역")

    if len(st.session_state.travel_list) > 0:
        raw_df = process_travel_data(st.session_state.travel_list)
        display_df = raw_df.drop(columns=["교통비항목리스트", "출장비항목리스트", "기타항목리스트", "지역구분", "직급구분", "환율"]).reset_index(drop=True)
        display_df.insert(0, "순번", range(1, len(display_df) + 1))

        gb = GridOptionsBuilder.from_dataframe(display_df)
        gb.configure_selection(selection_mode="single", use_checkbox=False)
        # 열 폭이 좁아 값이 잘리지 않도록 최소 폭 지정(넘치면 가로 스크롤)
        gb.configure_default_column(minWidth=110, resizable=True)
        gb.configure_column("순번", minWidth=90, maxWidth=100)
        # 숫자 열은 천단위 콤마 표시 + 오른쪽 정렬
        num_fmt = JsCode(
            "function(p){ if (p.value === null || p.value === undefined || p.value === '') { return ''; }"
            " var n = Number(p.value); return isNaN(n) ? p.value : n.toLocaleString('ko-KR'); }"
        )
        num_cols = [c for c in display_df.columns if pd.api.types.is_numeric_dtype(display_df[c])]
        if num_cols:
            gb.configure_columns(num_cols, valueFormatter=num_fmt, cellStyle={"textAlign": "right"}, minWidth=130)
        grid_options = gb.build()

        grid_response = AgGrid(
            display_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            fit_columns_on_grid_load=True,
            height=250,
            theme="alpine",
            allow_unsafe_jscode=True,
        )

        col_del1, col_del2 = st.columns([1, 1])
        with col_del1:
            if st.session_state.edit_target_index is not None:
                if st.button("수정 모드 취소"):
                    st.session_state.edit_target_index = None
                    st.rerun()
        with col_del2:
            if st.button("전체 데이터 초기화"):
                st.session_state.travel_list = []
                st.session_state.edit_target_index = None
                st.rerun()
    else:
        st.info("등록된 출장 내역이 없습니다.")

# ----------------------------------------------------
# 메뉴 2 & 3 (자금팀 / 산정내역서) - 기존 기능 유지
# ----------------------------------------------------
elif menu == MENU_2:
    st.markdown("### 자금팀 제출용 정산 집계표")
    if len(st.session_state.travel_list) > 0:
        _recs_all = st.session_state.travel_list
        _all_df = process_travel_data(_recs_all)
        fund_sel = st.multiselect(
            "자료를 생성할 출장 내역 선택",
            options=list(range(len(_recs_all))),
            default=list(range(len(_recs_all))),
            format_func=lambda i: f"{i + 1}. {_recs_all[i]['출장자성명']} ({_recs_all[i]['출장지']}, {_recs_all[i]['출장시작일']} ~ {_recs_all[i]['출장종료일']})",
            key=f"fund_sel_{len(_recs_all)}",
            placeholder=" ",
        )
        if not fund_sel:
            st.info("자료를 생성할 출장 내역을 한 건 이상 선택해 주세요.")
            render_footer()
            st.stop()
        fund_recs = [_recs_all[i] for i in fund_sel]
        processed_df = process_travel_data(fund_recs)
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
            "자금팀 연결 자료 엑셀 다운로드",
            data=build_fund_excel(fund_recs),
            file_name=f"자금팀_연결자료_{datetime.date.today():%Y%m%d}.xlsx",
            mime=XL_MIME,
            use_container_width=True,
        )
    else:
        st.info("등록된 출장 내역이 없습니다. '1. 출장 정보 입력' 메뉴에서 먼저 등록해주세요.")

elif menu == MENU_3:
    st.markdown("### 해외출장비 산정 내역서")
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
                f"[{selected_person}] 산정 내역서 엑셀 다운로드",
                data=build_statement_excel(records, [sel_idx]),
                file_name=f"출장비_산정내역서_{selected_person}_{datetime.date.today():%Y%m%d}.xlsx",
                mime=XL_MIME,
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "전체 출장자 일괄 다운로드 (출장자별 시트)",
                data=build_statement_excel(records, list(range(len(records)))),
                file_name=f"출장비_산정내역서_전체_{datetime.date.today():%Y%m%d}.xlsx",
                mime=XL_MIME,
                use_container_width=True,
            )
    else:
        st.info("등록된 출장 내역이 없습니다. '1. 출장 정보 입력' 메뉴에서 먼저 등록해주세요.")

# ----------------------------------------------------
# 홈 / 해외출장 내역관리 화면 및 공통 푸터
# ----------------------------------------------------
if mode == "home":
    render_home()
elif mode == "manage":
    render_manage(mg_menu)

render_footer()
