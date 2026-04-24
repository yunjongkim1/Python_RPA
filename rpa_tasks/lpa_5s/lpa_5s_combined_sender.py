# -*- coding: utf-8 -*-
"""
lpa_5s_combined_sender.py
────────────────────────────────────────────────────────────────────────────
G-MES에서 LPA + 5S 실적을 각각 두 플랜트(31111, 31311)씩 수집하여
달성률 100% 미만 항목을 하나의 이메일로 Outlook 자동 발송

수집 순서 (총 4회):
  1. LPA  – Plant 31111 (Alabama Plant #1)
  2. LPA  – Plant 31311 (Alabama Plant #2)
  3. 5S   – Plant 31111 (Alabama Plant #1)
  4. 5S   – Plant 31311 (Alabama Plant #2)

이메일 구조:
  [전체 요약]
  ── LPA ──────────────────────────
    🏭 Plant 31111
    🏭 Plant 31311
  ── 5S ───────────────────────────
    🏭 Plant 31111
    🏭 Plant 31311

사용법:
  python lpa_5s_combined_sender.py              # G-MES 자동수집 + 발송
  python lpa_5s_combined_sender.py --preview    # Outlook 미리보기 (발송 안 함)
  python lpa_5s_combined_sender.py --no-attach  # 첨부파일 없이 발송
────────────────────────────────────────────────────────────────────────────
"""

import os
import sys
import time
import shutil
import logging
import argparse
import textwrap
from datetime import date, datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

# .env 로드 (PROJECT_ROOT 등 경로 설정에 필요하므로 로거보다 먼저 로드)
_ENV_FILE = Path(__file__).resolve().parent / "lpa_5s_combined_sender.env"
_env_missing = not _ENV_FILE.exists()
if not _env_missing:
    load_dotenv(_ENV_FILE, override=True)

# ── 로거 설정 (PROJECT_ROOT 기반) ────────────────────────────────────
_PROJECT_ROOT = Path(os.getenv("PROJECT_ROOT", str(Path(__file__).resolve().parent)))
_LOG_DIR  = _PROJECT_ROOT / "Logs"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / f"lpa_5s_{datetime.now().strftime('%Y%m%d')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

if _env_missing:
    log.warning(f"[WARN] .env 파일 없음: {_ENV_FILE}")

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


# ══════════════════════════════════════════════════════════════════════
# ▌ 설정
# ══════════════════════════════════════════════════════════════════════

CHROMEDRIVER_EXE = os.getenv("CHROMEDRIVER_EXE", "")
DOWNLOAD_DIR     = Path(os.getenv("DOWNLOAD_DIR", str(Path.home() / "Downloads")))
OUTPUT_DIR       = Path(os.getenv("OUTPUT_DIR",   str(Path.home() / "Downloads")))

LOGIN_URL = os.getenv("GMES_LOGIN_URL", "https://gmes30-mcaal.hlmando.com/app/view.jsp?w2xPath=/home/login.xml")
GMES_ID   = os.getenv("GMES_ID", "")
GMES_PW   = os.getenv("GMES_PW", "")

LPA_URL = os.getenv("GMES_LPA_URL",
    "https://gmes30-mcaal.hlmando.com/app/view.jsp"
    "?w2xPath=/HPS/WHL/WHL0080_0.xml"
    "&menu_id=/v8ANAAzADg%3D"
    "&webParam=null&w2xHome=/home/&w2xDocumentRoot="
)
S5_URL = os.getenv("GMES_5S_URL",
    "https://gmes30-mcaal.hlmando.com/app/view.jsp"
    "?w2xPath=/HPS/WH5/WH50080_0.xml"
    "&menu_id=/v8ANAA0ADk%3D"
    "&webParam=null&w2xHome=/home/&w2xDocumentRoot="
)
KPI_URL = os.getenv("GMES_KPI_URL",
    "https://gmes30-mcaal.hlmando.com/app/view.jsp"
    "?w2xPath=/HPS/WHK/WHK0080_0.xml"
    "&menu_id=/v8AMQAyADAAOQ%3D%3D"
    "&webParam=null&w2xHome=/home/&w2xDocumentRoot="
)

# ── KPI 그래프 설정 ───────────────────────────────────────────────────
KPI_GRAPH_CANVAS_ID = "graphCanvas"   # Console에서 확인된 캔버스 ID
KPI_SUBPLANTS = [
    {"index": 0, "code": "31111", "name": "Alabama Plant #1"},
    {"index": 1, "code": "31311", "name": "Alabama Plant #2"},
]

# ── 수집 작업 목록 (순서대로 실행) ────────────────────────────────────
# index: SubPlant 콤보 드롭다운 아이템 번호 (0부터 시작)
COLLECT_TASKS = [
    {"type": "LPA", "url": LPA_URL, "index": 0, "code": "31111", "name": "Alabama Plant #1"},
    {"type": "LPA", "url": LPA_URL, "index": 1, "code": "31311", "name": "Alabama Plant #2"},
    {"type": "5S",  "url": S5_URL,  "index": 0, "code": "31111", "name": "Alabama Plant #1"},
    {"type": "5S",  "url": S5_URL,  "index": 1, "code": "31311", "name": "Alabama Plant #2"},
]

# ── G-MES 화면 요소 ID ────────────────────────────────────────────────
SUBPLANT_BTN_ID = "cboSearchFactory_button"
DATE_FROM_ID    = "txtSearchPeriodBeginDate_input"
DATE_TO_ID      = "txtSearchPeriodEndDate_input"
INQUIRY_BTN_ID  = "btnMainSearch"
DOWNLOAD_BTN_ID = "btnGrid1DownloadExcel"

# ── 이메일 설정 (lpa_5s_combined_sender.env 에서 로드) ───────────────
_MAIL_MODE = os.getenv("MAIL_MODE", "test").lower()
if _MAIL_MODE == "test":
    _dev     = os.getenv("DEVELOPER_EMAIL_TEST", "")
    EMAIL_TO = [e.strip() for e in _dev.split(",") if e.strip()]
    EMAIL_CC = []
else:
    EMAIL_TO = [v.split('#')[0].strip() for k, v in sorted(os.environ.items())
                if k.startswith("EMAIL_TO_") and v.strip()]
    EMAIL_CC = [v.split('#')[0].strip() for k, v in sorted(os.environ.items())
                if k.startswith("EMAIL_CC_") and v.strip()]
EMAIL_BCC = [v.split('#')[0].strip() for k, v in sorted(os.environ.items())
             if k.startswith("EMAIL_BCC_") and v.strip()]

# 메일 발송 전 설정 확인 출력
log.info(f"[MAIL] 모드: {_MAIL_MODE.upper()} | TO: {len(EMAIL_TO)}명 | CC: {len(EMAIL_CC)}명")
SUBJECT_TEMPLATE = "[HPS Audit Report] {date_range} LPA/5S Achievement Status"

# ── 필터링 임계값 ─────────────────────────────────────────────────────
THRESHOLD = 100.0

# ── 컬럼 매핑 (LPA / 5S 공통) ────────────────────────────────────────
COLUMN_MAP = {
    "W/C"           : ["W/C", "WC", "작업장"],
    "Layer"         : ["Layer", "레이어", "계층"],
    "점검자"        : ["Implement Person", "Implement  Person", "점검자", "담당자", "Auditor"],
    "Shift"         : ["Shift", "시프트", "근무조"],
    "점검일"        : ["Plan Date", "점검일", "계획일", "Check Date", "Date"],
    "상태"          : ["Inspection Status", "상태", "Status"],
    "계획"          : ["Plan Quantity", "계획", "Plan", "목표"],
    "실적"          : ["Inspection Number", "실적", "Actual"],
    "달성률"        : ["Implementation Rate", "달성률", "달성율",
                      "Achievement Rate", "Achievement", "Rate", "%"],
    "Action Result" : ["Action Result", "조치결과"],
    "NG(N)"         : ["NG Quantity(N)", "NG(N)"],
    "NG(NC)"        : ["NG Quantity(NC)", "NG(NC)"],
    "점검시간"      : ["Inspection Time", "점검시간"],
}

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════
# ▌ 1. 날짜 계산
# ══════════════════════════════════════════════════════════════════════

def compute_date_range():
    today  = date.today()
    to_d   = today - timedelta(days=1)
    from_d = date(to_d.year, to_d.month, 1)
    return from_d, to_d

def fmt_date(d: date) -> str:
    return d.strftime("%m/%d/%Y")


# ══════════════════════════════════════════════════════════════════════
# ▌ 2. Selenium 유틸
# ══════════════════════════════════════════════════════════════════════

def build_driver():
    options = Options()
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)
    if CHROMEDRIVER_EXE and Path(CHROMEDRIVER_EXE).is_file():
        try:
            driver = webdriver.Chrome(service=Service(CHROMEDRIVER_EXE), options=options)
        except Exception as e:
            if "version" in str(e).lower() or "session not created" in str(e).lower():
                log.info("[WARN] ChromeDriver mismatch → Selenium Manager fallback")
                driver = webdriver.Chrome(options=options)
            else:
                raise
    else:
        log.info("[INFO] ChromeDriver 경로 없음 → Selenium Manager 자동 감지")
        driver = webdriver.Chrome(options=options)
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow", "downloadPath": str(DOWNLOAD_DIR.resolve())
        })
    except:
        pass
    return driver


def dismiss_alert(driver, timeout=2):
    end = time.time() + timeout
    while time.time() < end:
        try:
            driver.switch_to.alert.accept()
            time.sleep(0.2)
            return True
        except:
            time.sleep(0.1)
    return False


def wait_no_overlay(driver, timeout=90):
    end = time.time() + timeout
    while time.time() < end:
        try:
            driver.switch_to.default_content()
            masks = driver.find_elements(
                By.XPATH, "//*[contains(@class,'blockUI') or contains(@class,'w2mask')]"
            )
            if not any(m.is_displayed() for m in masks):
                return True
        except:
            pass
        time.sleep(0.3)
    return False


def js_click(driver, el):
    driver.execute_script("""
        arguments[0].scrollIntoView({block:'center'});
        ['mousedown','mouseup','click'].forEach(t =>
          arguments[0].dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true}))
        );
    """, el)


def safe_click_id(driver, el_id, timeout=20):
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, el_id))
    )
    try:
        el.click()
    except:
        js_click(driver, el)
    return el


def accept_confirm(driver, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        if dismiss_alert(driver, 1):
            return
        for sel in ["button.ajs-ok", "button.ajs-button.ajs-ok"]:
            try:
                btns = [b for b in driver.find_elements(By.CSS_SELECTOR, sel)
                        if b.is_displayed()]
                if btns:
                    js_click(driver, btns[0])
                    return
            except:
                pass
        for xp in ["//*[normalize-space()='Yes']", "//*[normalize-space()='YES']"]:
            try:
                btns = [b for b in driver.find_elements(By.XPATH, xp)
                        if b.is_displayed()]
                if btns:
                    js_click(driver, btns[0])
                    return
            except:
                pass
        time.sleep(0.3)


def wait_download(start_ts: float, timeout=120) -> Path:
    """
    다운로드 완료 파일 감지.
    mtime 기반 + G-MES 고정 파일명(GMES_wh*.xls) 덮어쓰기 감지 병행.
    """
    gmes_patterns = ["GMES_wh*.xls", "GMES_wh*.xlsx", "Detailed Status*.xls", "Detailed Status*.xlsx"]
    snapshot = {}
    for pat in gmes_patterns:
        for p in DOWNLOAD_DIR.glob(pat):
            try:
                snapshot[p.name] = p.stat().st_size
            except:
                pass

    end = time.time() + timeout
    while time.time() < end:
        # 방법1: mtime 기반 (새 파일) — .xls/.xlsx만 감지
        for p in DOWNLOAD_DIR.glob("*"):
            if not p.is_file(): continue
            if not p.name.lower().endswith((".xls", ".xlsx")): continue
            if p.name.lower().endswith((".crdownload", ".tmp", ".part")): continue
            try:
                if p.stat().st_mtime >= start_ts:
                    s1 = p.stat().st_size
                    time.sleep(1.0)
                    s2 = p.stat().st_size
                    if s1 > 0 and s1 == s2:
                        return p
            except:
                pass
        # 방법2: G-MES 고정 파일명 덮어쓰기 감지
        for pat in gmes_patterns:
            for p in DOWNLOAD_DIR.glob(pat):
                try:
                    cur = p.stat().st_size
                    if cur != snapshot.get(p.name, -1) and cur > 0:
                        time.sleep(1.0)
                        if p.stat().st_size == cur:
                            log.info(f"[DL] G-MES 고정 파일명 감지: {p.name}")
                            return p
                except:
                    pass
        time.sleep(0.5)
    # 타임아웃 전 다운로드 폴더 현황 출력
    log.debug(f"[DEBUG] 다운로드 폴더 최근 파일 5개:")
    files = sorted(DOWNLOAD_DIR.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)
    for f in files[:5]:
        log.info(f"  {f.name}  ({f.stat().st_size}bytes, mtime={f.stat().st_mtime:.0f})")
    log.debug(f"[DEBUG] start_ts={start_ts:.0f}, now={time.time():.0f}")
    raise TimeoutError("다운로드 타임아웃")


# ══════════════════════════════════════════════════════════════════════
# ▌ 3. G-MES 로그인
# ══════════════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════════════
# ▌ 3. KPI 그래프 캡처
# ══════════════════════════════════════════════════════════════════════

def capture_kpi_graph(driver, code: str, index: int):
    """
    KPI > Total performance by plant 화면에 직접 URL로 진입,
    SubPlant 선택 후 graphCanvas를 base64 PNG로 캡처.
    """
    import base64
    log.info(f"[KPI] Plant {code} 그래프 캡처 시작...")
    try:
        # ── 직접 URL 접근 ────────────────────────────────────────────
        driver.get(KPI_URL)
        dismiss_alert(driver, 3)
        wait_no_overlay(driver, 120)
        time.sleep(3)

        # ── SubPlant 선택 ────────────────────────────────────────────
        current = driver.execute_script(
            "const el=document.getElementById('cboSearchFactory_label');"
            "return el ? (el.innerText||el.textContent||'').trim() : '';"
        ) or ""

        if code not in current:
            log.info(f"[KPI] SubPlant '{current}' → {code} 변경 중...")
            safe_click_id(driver, "cboSearchFactory_button", timeout=10)
            time.sleep(0.5)
            item_id = f"cboSearchFactory_itemTable_{index}"
            el = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, item_id))
            )
            js_click(driver, el)
            time.sleep(0.5)
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except:
                pass
            wait_no_overlay(driver, 60)
            time.sleep(1)
            log.info(f"[KPI] SubPlant {code} 선택 완료")
        else:
            log.info(f"[KPI] SubPlant 이미 선택됨: '{current}'")

        # ── Inquiry 클릭 (버튼 타입 무관하게 다양한 방법 시도) ──────
        clicked = False
        # 1) 텍스트로 찾기
        for xpath in [
            "//*[normalize-space(text())='Inquiry']",
            "//*[normalize-space(text())='inquiry']",
            "//*[contains(@onclick,'search') or contains(@onclick,'Search')]",
            "//*[contains(@id,'btnInquiry') or contains(@id,'btnSearch') or contains(@id,'btnMainSearch')]",
        ]:
            try:
                els = driver.find_elements(By.XPATH, xpath)
                for el in els:
                    if el.is_displayed():
                        js_click(driver, el)
                        clicked = True
                        log.info(f"[KPI] Inquiry 클릭 성공")
                        break
                if clicked:
                    break
            except:
                pass

        # 2) Inquiry 버튼 못 찾으면 SubPlant 변경만으로도 자동 조회될 수 있음
        if not clicked:
            log.info(f"[KPI] Inquiry 버튼 못 찾음 → 자동 조회 대기")

        wait_no_overlay(driver, 60)
        time.sleep(8)  # 차트 렌더링 충분히 대기

        # ── graphCanvas → base64 PNG (흰색 배경 합성) ──────────────
        canvas_b64 = driver.execute_script(f"""
            const c = document.getElementById('{KPI_GRAPH_CANVAS_ID}');
            if (!c) return null;
            try {{
                const offscreen = document.createElement('canvas');
                offscreen.width  = c.width;
                offscreen.height = c.height;
                const ctx = offscreen.getContext('2d');
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, offscreen.width, offscreen.height);
                ctx.drawImage(c, 0, 0);
                return offscreen.toDataURL('image/png').split(',')[1];
            }} catch(e) {{ return null; }}
        """)

        if canvas_b64 and len(canvas_b64) > 100:
            log.info(f"[KPI] Plant {code} 캡처 성공 ({len(canvas_b64):,} chars)")
            return canvas_b64

        # ── fallback: canvas 요소 스크린샷 ──────────────────────────
        log.info(f"[KPI] toDataURL 실패 → element screenshot 시도...")
        for selector in [
            f"#{KPI_GRAPH_CANVAS_ID}",
            "canvas",
            "[id*='graph']",
            "[id*='chart']",
            ".highcharts-container",
        ]:
            try:
                el = driver.find_element(By.CSS_SELECTOR, selector)
                if el.is_displayed() and el.size['width'] > 50:
                    b64 = base64.b64encode(el.screenshot_as_png).decode()
                    log.info(f"[KPI] element screenshot 성공 (selector='{selector}')")
                    return b64
            except:
                pass

        log.info(f"[KPI] Plant {code} 캡처 실패 — 이메일에 그래프 미포함")
        return None

    except Exception as e:
        log.info(f"[KPI] 캡처 오류: {e}")
        return None


def capture_all_kpi_graphs(driver) -> dict:
    graphs = {}
    for sp in KPI_SUBPLANTS:
        graphs[sp["code"]] = capture_kpi_graph(driver, sp["code"], sp["index"])
    return graphs

def login(driver):
    driver.get(LOGIN_URL)
    dismiss_alert(driver, 3)
    WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.ID, "txtUSER_ID"))
    ).clear()
    driver.find_element(By.ID, "txtUSER_ID").send_keys(GMES_ID)
    driver.find_element(By.ID, "txtPWD").clear()
    driver.find_element(By.ID, "txtPWD").send_keys(GMES_PW)
    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, "btnLogin"))
    ).click()
    dismiss_alert(driver, 3)
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.ID, "btn_nav07"))
    )
    wait_no_overlay(driver, 120)
    time.sleep(1)
    log.info("[LOGIN] 완료")


# ══════════════════════════════════════════════════════════════════════
# ▌ 4. 화면 진입
# ══════════════════════════════════════════════════════════════════════

def navigate_to(driver, url: str, label: str):
    log.info(f"[NAV] {label} 화면으로 이동...")
    driver.get(url)
    dismiss_alert(driver, 3)
    wait_no_overlay(driver, 120)
    time.sleep(2)
    end = time.time() + 30
    while time.time() < end:
        try:
            if driver.find_element(By.ID, INQUIRY_BTN_ID):
                log.info(f"[NAV] {label} 화면 로드 완료")
                return
        except:
            pass
        time.sleep(0.5)
    raise RuntimeError(f"{label} 화면 로드 실패. URL: {driver.current_url}")


# ══════════════════════════════════════════════════════════════════════
# ▌ 5. SubPlant 선택
# ══════════════════════════════════════════════════════════════════════

def select_subplant(driver, item_index: int, code: str):
    current = driver.execute_script(
        "const el=document.getElementById('cboSearchFactory_label');"
        "return el ? (el.innerText||el.textContent||'').trim() : '';"
    ) or ""

    if code in current:
        log.info(f"[SUBPLANT] 이미 선택됨: '{current}'")
        return

    log.info(f"[SUBPLANT] '{current}' → {code} 선택 중...")
    safe_click_id(driver, SUBPLANT_BTN_ID, timeout=10)
    time.sleep(0.4)

    item_id = f"cboSearchFactory_itemTable_{item_index}"
    el = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, item_id))
    )
    js_click(driver, el)
    time.sleep(0.4)

    driver.execute_script("""
        const el = document.getElementById('cboSearchFactory_label');
        if (el) ['change','blur'].forEach(t =>
            el.dispatchEvent(new Event(t, {bubbles:true}))
        );
    """)
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except:
        pass
    time.sleep(0.5)

    after = driver.execute_script(
        "const el=document.getElementById('cboSearchFactory_label');"
        "return el ? (el.innerText||el.textContent||'').trim() : '';"
    ) or ""
    log.info(f"[SUBPLANT] 완료: '{after}'")
    if code not in after:
        raise RuntimeError(f"SubPlant {code} 선택 실패. 현재: '{after}'")


# ══════════════════════════════════════════════════════════════════════
# ▌ 6. 날짜 입력
# ══════════════════════════════════════════════════════════════════════

def set_dates(driver, from_str: str, to_str: str):
    def _set(el_id, val):
        driver.execute_script(f"""
            var el = document.getElementById('{el_id}');
            if (el) {{
                el.value = '{val}';
                ['input','change','blur'].forEach(function(t) {{
                    el.dispatchEvent(new Event(t, {{bubbles:true}}));
                }});
            }}
        """)
        time.sleep(0.3)
        dismiss_alert(driver, 2)
    _set(DATE_FROM_ID, from_str)
    _set(DATE_TO_ID,   to_str)
    log.info(f"[DATE] From={from_str}  To={to_str}")


# ══════════════════════════════════════════════════════════════════════
# ▌ 7. 다운로드
# ══════════════════════════════════════════════════════════════════════

def download_excel(driver, label: str) -> Path:
    # 기존 G-MES 고정 파일명 삭제 (덮어쓰기 충돌 방지)
    for pat in ["GMES_wh*.xls", "GMES_wh*.xlsx", "Detailed Status*.xls", "Detailed Status*.xlsx"]:
        for old in DOWNLOAD_DIR.glob(pat):
            try:
                old.unlink()
                log.info(f"[DL] 이전 파일 삭제: {old.name}")
            except:
                pass

    start_ts = time.time()
    safe_click_id(driver, DOWNLOAD_BTN_ID, timeout=20)
    log.info(f"[DL] 다운로드 버튼 클릭")
    time.sleep(1.5)
    accept_confirm(driver, timeout=5)
    time.sleep(1.0)
    accept_confirm(driver, timeout=3)

    log.info(f"[DL] 파일 대기 중...")
    raw = wait_download(start_ts=start_ts, timeout=120)

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = OUTPUT_DIR / f"{label}_{ts}{raw.suffix}"
    shutil.move(str(raw), str(dst))
    log.info(f"[DL] 저장 완료: {dst.name}")
    return dst


# ══════════════════════════════════════════════════════════════════════
# ▌ 8. G-MES 자동수집 (4개 작업 순차 실행)
# ══════════════════════════════════════════════════════════════════════

def collect_all() -> tuple[list[dict], date, date]:
    """
    COLLECT_TASKS를 순서대로 실행하여 각 파일 경로를 반환.
    Returns: (results, from_d, to_d)
      results: [{"type","code","name","path"}, ...]
    """
    if not GMES_PW:
        raise RuntimeError("환경변수 GMES_PW 가 설정되지 않았습니다.")

    from_d, to_d = compute_date_range()
    from_str = fmt_date(from_d)
    to_str   = fmt_date(to_d)
    log.info(f"[DATE] 조회기간: {from_str} ~ {to_str}")

    driver  = build_driver()
    results = []

    kpi_graphs = {}
    try:
        login(driver)

        for task in COLLECT_TASKS:
            label = f"{task['type']}_{task['code']}"
            log.info(f"\n{'='*52}")
            log.info(f"[COLLECT] {task['type']} – Plant {task['code']} ({task['name']})")
            log.info(f"{'='*52}")

            navigate_to(driver, task["url"], task["type"])
            select_subplant(driver, task["index"], task["code"])
            set_dates(driver, from_str, to_str)

            log.info("[INQUIRY] 조회...")
            safe_click_id(driver, INQUIRY_BTN_ID, timeout=20)
            wait_no_overlay(driver, 120)
            time.sleep(2)
            log.info("[INQUIRY] 완료")

            # "No data." 텍스트로 데이터 없음 감지
            no_data = driver.execute_script("""
                return document.body.innerText.indexOf('No data.') !== -1;
            """)
            log.info(f"[GRID] No data: {no_data}")

            if no_data:
                log.warning(f"[WARN] No data → 다운로드 건너뜀")
                results.append({
                    "type": task["type"], "code": task["code"],
                    "name": task["name"], "path": None
                })
                continue
            log.info(f"[GRID] 데이터 있음 → 다운로드 시도")

            path = download_excel(driver, label)
            results.append({
                "type": task["type"], "code": task["code"],
                "name": task["name"], "path": path
            })

        # KPI 그래프 캡처 (LPA/5S 수집 완료 후, driver 종료 전)
        log.info("[KPI] 그래프 캡처 시작...")
        kpi_graphs = capture_all_kpi_graphs(driver)

    except Exception as e:
        log.error(f"[ERROR] 수집 중 오류: {e}")
        raise
    finally:
        driver.quit()

    return results, from_d, to_d, kpi_graphs


# ══════════════════════════════════════════════════════════════════════
# ▌ 9. Excel 로드 및 필터링
# ══════════════════════════════════════════════════════════════════════

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for std_name, candidates in COLUMN_MAP.items():
        for col in df.columns:
            if str(col).strip() in candidates:
                rename_map[col] = std_name
                break
    if rename_map:
        log.info(f"[NORMALIZE] {rename_map}")
    return df.rename(columns=rename_map)


def load_excel(filepath: Path) -> pd.DataFrame:
    log.info(f"[LOAD] {filepath.name}")
    rate_cands = COLUMN_MAP["달성률"]
    for header_row in range(10):
        df = pd.read_excel(filepath, header=header_row, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        for cand in rate_cands:
            if cand in df.columns:
                log.info(f"[LOAD] 헤더 행 {header_row} 확정 ('{cand}' 발견)")
                return normalize_columns(df)
    log.info("[WARN] 달성률 컬럼 자동 탐색 실패 → header=0 fallback")
    df = pd.read_excel(filepath, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return normalize_columns(df.dropna(how="all"))


def parse_rate(val) -> float | None:
    if pd.isna(val) or str(val).strip() in ("", "-", "N/A", "nan"):
        return None
    s = str(val).strip().replace(",", "").replace(" ", "")
    try:
        if s.endswith("%"):
            return float(s[:-1])
        v = float(s)
        return v * 100 if 0 < v <= 1.0 else v
    except ValueError:
        return None


def filter_below(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    total = len(df)
    if "달성률" not in df.columns:
        log.warning(f"[WARN] '달성률' 컬럼 없음: {list(df.columns)}")
        return df, total
    df = df.copy()
    df["_r"] = df["달성률"].apply(parse_rate)
    below = df[df["_r"].notna() & (df["_r"] < THRESHOLD)].drop(columns=["_r"])
    df    = df.drop(columns=["_r"])
    log.info(f"[FILTER] 전체: {total}행 → 100% 미만: {len(below)}행")
    return below, total


# ══════════════════════════════════════════════════════════════════════
# ▌ 10. 이메일 HTML 본문 생성
# ══════════════════════════════════════════════════════════════════════

def _rate_color(val: str) -> str:
    r = parse_rate(val)
    if r is None: return "#333"
    if r == 0:    return "#cc0000"
    if r < 50:    return "#e65c00"
    if r < 80:    return "#e6b800"
    return "#cc7700"


def df_to_html_table(df: pd.DataFrame) -> str:
    TH = ('style="background:#1F4E79;color:#fff;padding:8px 12px;'
          'text-align:center;border:1px solid #ccc;white-space:nowrap;"')
    TD = ('style="padding:6px 10px;border:1px solid #ddd;'
          'text-align:center;vertical-align:middle;"')
    headers = "".join(f"<th {TH}>{col}</th>" for col in df.columns)
    rows = ""
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#f5f8fc" if i % 2 == 0 else "#fff"
        cells = ""
        for col in df.columns:
            v = "" if pd.isna(row[col]) else str(row[col]).strip()
            if col == "달성률" and v:
                v = f'<span style="color:{_rate_color(v)};font-weight:bold">{v}</span>'
            cells += f"<td {TD}>{v}</td>"
        rows += f'<tr style="background:{bg}">{cells}</tr>\n'
    return (
        '<table style="border-collapse:collapse;width:100%;'
        'font-family:맑은 고딕,Arial,sans-serif;font-size:12px;">'
        f"<thead><tr>{headers}</tr></thead><tbody>{rows}</tbody></table>"
    )


def plant_section(ptype: str, code: str, name: str,
                  df_below: pd.DataFrame, total: int) -> str:
    n   = len(df_below)
    pct = round(n / total * 100, 1) if total > 0 else 0

    badge = (
        '<span style="background:#2e7d32;color:#fff;padding:3px 10px;'
        'border-radius:12px;font-size:12px;">✅ 전체 달성</span>'
        if n == 0 else
        f'<span style="background:#cc0000;color:#fff;padding:3px 10px;'
        f'border-radius:12px;font-size:12px;">⚠️ 미달성 {n}건</span>'
    )
    tbl = (
        df_to_html_table(df_below) if n > 0 else
        f'<div style="background:#e8f5e9;border:1px solid #a5d6a7;border-radius:6px;'
        f'padding:14px;text-align:center;">'
        f'<strong style="color:#2e7d32;">전체 {total}건 모두 달성률 100% 이상입니다. 🎉</strong>'
        f'</div>'
    )
    return f"""
  <div style="margin-bottom:16px;">
    <div style="background:#2E4057;color:#fff;padding:9px 14px;border-radius:6px 6px 0 0;
                display:flex;justify-content:space-between;align-items:center;">
      <strong style="font-size:13px;">🏭 [{ptype}] Plant {code} – {name}</strong>
      {badge}
    </div>
    <div style="background:#f8f9fa;border:1px solid #dee2e6;border-top:none;
                padding:8px 14px;">
      <span style="font-size:12px;color:#555;">
        전체 <strong>{total}</strong>건 &nbsp;|&nbsp;
        100% 미만 <strong style="color:#cc0000;">{n}건 ({pct}%)</strong>
      </span>
    </div>
    <div style="overflow-x:auto;border:1px solid #dee2e6;border-top:none;
                border-radius:0 0 6px 6px;padding:8px;">
      {tbl}
    </div>
  </div>
"""


def category_block(ptype: str, sections_html: str,
                   total: int, below: int) -> str:
    """LPA 또는 5S 카테고리 전체 블록"""
    color    = "#1F4E79" if ptype == "LPA" else "#5B4A8A"
    icon     = "📋" if ptype == "LPA" else "🧹"
    pct      = round(below / total * 100, 1) if total > 0 else 0
    status   = (
        '<span style="color:#2e7d32;font-weight:bold;">✅ 전체 달성</span>'
        if below == 0 else
        f'<span style="color:#cc0000;font-weight:bold;">⚠️ 미달성 {below}건 ({pct}%)</span>'
    )
    return f"""
  <!-- {ptype} 카테고리 -->
  <div style="margin-bottom:28px;">
    <div style="background:{color};color:#fff;padding:12px 18px;
                border-radius:8px 8px 0 0;">
      <h3 style="margin:0;font-size:15px;">{icon} {ptype} 실적 현황</h3>
    </div>
    <div style="background:#eef2f7;border:1px solid #c5d0de;border-top:none;
                padding:8px 18px;border-radius:0 0 0 0;">
      <span style="font-size:13px;">
        전체 <strong>{total}</strong>건 &nbsp;|&nbsp; {status}
      </span>
    </div>
    <div style="border:1px solid #c5d0de;border-top:none;
                border-radius:0 0 8px 8px;padding:12px;">
      {sections_html}
    </div>
  </div>
"""


def build_email_body(processed: list[dict],
                     from_d: date, to_d: date,
                     kpi_graphs: dict = None) -> str:
    """
    processed: [{"type","code","name","df_below","total"}, ...]
    """
    dr = f"{from_d.strftime('%Y-%m-%d')} ~ {to_d.strftime('%Y-%m-%d')}"

    # KPI 그래프 이미지 섹션 생성
    kpi_graphs = kpi_graphs or {}
    kpi_imgs = []
    for sp in KPI_SUBPLANTS:
        code = sp["code"]
        name = sp["name"]
        b64  = kpi_graphs.get(code)
        if b64:
            img_tag = (
                f'<td width="50%" style="text-align:center;padding:8px;">'
                f'<div style="font-size:13px;font-weight:bold;color:#1F4E79;'
                f'margin-bottom:6px;">Plant {code} &#8211; {name}</div>'
                f'<img src="data:image/png;base64,{b64}" '
                f'style="width:100%;max-width:480px;border:1px solid #ddd;border-radius:6px;" />'
                f'</td>'
            )
        else:
            img_tag = (
                f'<td width="50%" style="text-align:center;color:#aaa;'
                f'font-size:12px;padding:20px;">'
                f'Plant {code} 그래프 캡처 실패</td>'
            )
        kpi_imgs.append(img_tag)

    if any(kpi_graphs.get(sp["code"]) for sp in KPI_SUBPLANTS):
        kpi_section = (
            '<div style="margin-bottom:20px;">'
            '<div style="background:#2E4057;color:#fff;padding:10px 16px;'
            'border-radius:6px 6px 0 0;font-size:14px;font-weight:bold;">'
            '&#128202; Total Performance by Plant (KPI 현황)</div>'
            '<div style="border:1px solid #dee2e6;border-top:none;'
            'border-radius:0 0 6px 6px;padding:16px;">'
            '<table width="100%" cellpadding="8" cellspacing="0" border="0">'
            '<tr valign="top">'
            + "".join(kpi_imgs) +
            '</tr></table>'
            '</div></div>'
        )
    else:
        kpi_section = ""

    # LPA / 5S 분리
    lpa_items = [r for r in processed if r["type"] == "LPA"]
    s5_items  = [r for r in processed if r["type"] == "5S"]

    def _build_category(ptype, items):
        secs  = "".join(
            plant_section(ptype, r["code"], r["name"], r["df_below"], r["total"])
            for r in items
        )
        total = sum(r["total"] for r in items)
        below = sum(len(r["df_below"]) for r in items)
        return category_block(ptype, secs, total, below), total, below

    lpa_block, lpa_total, lpa_below = _build_category("LPA", lpa_items)
    s5_block,  s5_total,  s5_below  = _build_category("5S",  s5_items)

    grand_total = lpa_total + s5_total
    grand_below = lpa_below + s5_below
    grand_pct   = round(grand_below / grand_total * 100, 1) if grand_total > 0 else 0

    guide = (
        '<strong style="color:#2e7d32;">LPA·5S 전체 항목 모두 100% 달성되었습니다. 수고하셨습니다! 🎉</strong>'
        if grand_below == 0 else
        '아래 <strong style="color:#cc0000;">달성률 100% 미만</strong> 항목에 대한 원인 파악 및 조치를 부탁드립니다.'
    )

    return f"""<html><head><meta charset="utf-8"></head>
<body style="font-family:맑은 고딕,Arial,sans-serif;font-size:14px;color:#333;padding:20px;max-width:1100px;">

  <!-- 안내 문구 (최상단) -->
  <p style="margin-bottom:20px;">
    안녕하세요,<br><br>
    G-MES HPS 실적을 공유드립니다. 아래 달성률 100% 미만 항목에 대한 원인 파악 및 조치를 부탁드립니다.
  </p>

  <!-- KPI 그래프 섹션 -->
  {kpi_section}

  <!-- 메인 헤더 -->
  <div style="background:#1F4E79;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0;">
    <h2 style="margin:0;font-size:17px;">
      📊 HPS 실적 현황 알림 (LPA · 5S) &nbsp;|&nbsp;
      <span style="font-weight:normal;font-size:14px;">{dr}</span>
    </h2>
  </div>

  <!-- 전체 요약 -->
  <div style="background:#f0f4f8;border:1px solid #c0d0e0;border-top:none;
              padding:14px 20px;border-radius:0 0 8px 8px;margin-bottom:24px;">
    <table style="border:none;width:auto;border-spacing:0;">
      <tr>
        <td style="padding:4px 20px 4px 0;border:none;"><strong>🗓 조회기간</strong></td>
        <td style="padding:4px 0;border:none;">{dr}</td>
      </tr>
      <tr>
        <td style="padding:4px 20px 4px 0;border:none;"><strong>📋 LPA 전체</strong></td>
        <td style="padding:4px 0;border:none;">
          {lpa_total}건 &nbsp;|&nbsp;
          100% 미만 <span style="color:#cc0000;font-weight:bold;">{lpa_below}건</span>
        </td>
      </tr>
      <tr>
        <td style="padding:4px 20px 4px 0;border:none;"><strong>🧹 5S 전체</strong></td>
        <td style="padding:4px 0;border:none;">
          {s5_total}건 &nbsp;|&nbsp;
          100% 미만 <span style="color:#cc0000;font-weight:bold;">{s5_below}건</span>
        </td>
      </tr>
      <tr>
        <td style="padding:4px 20px 4px 0;border:none;"><strong>⚠️ 합산 미달성</strong></td>
        <td style="padding:4px 0;border:none;">
          <span style="color:#cc0000;font-weight:bold;">{grand_below}건 ({grand_pct}%)</span>
        </td>
      </tr>
    </table>
  </div>

  <!-- LPA 섹션 -->
  {lpa_block}

  <!-- 5S 섹션 -->
  {s5_block}

  <!-- 주석 -->
  <p style="font-size:12px;color:#888;border-top:1px solid #eee;padding-top:12px;margin-top:8px;">
    ※ 본 메일은 G-MES HPS RPA 자동 발송 메일입니다.<br>
    ※ 수집 기준: Plant 31111 / 31311 / 조회기간 {dr}<br>
    ※ 문의: jakyung.koo@hlcompany.com (OE Team)
  </p>

</body></html>"""


# ══════════════════════════════════════════════════════════════════════
# ▌ 11. Outlook 발송
# ══════════════════════════════════════════════════════════════════════

def send_via_outlook(subject: str, body_html: str,
                     to: list, cc: list = None, bcc: list = None,
                     attachments: list = None,
                     preview: bool = False) -> bool:
    try:
        import win32com.client as win32
    except ImportError:
        log.error("[ERROR] pip install pywin32")
        return False
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)
        mail.Subject  = subject
        mail.HTMLBody = body_html
        mail.To       = "; ".join(to)
        if cc:  mail.CC  = "; ".join(cc)
        if bcc: mail.BCC = "; ".join(bcc)
        for att in (attachments or []):
            att = Path(att)
            if att.exists():
                mail.Attachments.Add(str(att.resolve()))
                log.info(f"[ATTACH] 첨부: {att.name}")
        if preview:
            mail.Display()
            log.info("[PREVIEW] Outlook 창을 열었습니다. 확인 후 수동으로 발송하세요.")
        else:
            mail.Send()
            log.info(f"[SEND] ✅ 발송 완료 → {'; '.join(to)}")
        return True
    except Exception as e:
        log.error(f"[ERROR] Outlook 발송 실패: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════
# ▌ 12. 메인 파이프라인
# ══════════════════════════════════════════════════════════════════════

def run(preview=False):
    # ── Step 0: 오래된 산출물 정리 (7일 초과 LPA/5S 파일 삭제) ──────
    cutoff = datetime.now() - timedelta(days=7)
    for f in OUTPUT_DIR.glob("*.xls*"):
        if any(f.name.startswith(p) for p in ("LPA_", "5S_", "KPI_")):
            if datetime.fromtimestamp(f.stat().st_mtime) < cutoff:
                try:
                    f.unlink()
                    log.info(f"[CLEAN] 삭제: {f.name}")
                except Exception as e:
                    log.info(f"[CLEAN] 삭제 실패: {f.name} ({e})")

    # ── Step 1: G-MES 자동수집 ───────────────────────────────────────
    raw_results, from_d, to_d, kpi_graphs = collect_all()

    dr = f"{from_d.strftime('%Y.%m.%d')}~{to_d.strftime('%Y.%m.%d')}"

    # ── Step 2: Plant Mobile Audit Report 업데이트 ───────────────────
    lpa_path = next((r["path"] for r in raw_results
                     if r["type"] == "LPA" and r["code"] == "31111" and r["path"]), None)
    s5_path  = next((r["path"] for r in raw_results
                     if r["type"] == "5S"  and r["code"] == "31111" and r["path"]), None)

    report_path    = None
    rates_img_b64  = None

    if lpa_path and s5_path:
        try:
            import importlib.util as _ilu
            _mod_path = Path(__file__).resolve().parent / "plant_report_updater.py"
            _spec = _ilu.spec_from_file_location("plant_report_updater", str(_mod_path))
            _mod  = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_mod)
            update_report = _mod.update_report
            log.info("\n[REPORT] Plant Mobile Audit Report 업데이트 시작...")
            report_path, rates_img_b64 = update_report(lpa_path, s5_path, from_d=from_d, to_d=to_d)
            log.info(f"[REPORT] ✅ 완료: {report_path.name}")
        except Exception as e:
            import traceback
            log.info(f"[REPORT] ⚠️ 업데이트 실패: {e}")
            traceback.print_exc()
    else:
        log.info("[REPORT] LPA 또는 5S 파일 없음 → 리포트 업데이트 건너뜀")

    # ── Step 3: 이메일 본문 구성 ─────────────────────────────────────
    # KPI 그래프 좌우 배치 HTML
    kpi_html = ""
    if kpi_graphs:
        kpi_imgs_html = ""
        for sp in [{"code": "31111", "name": "Alabama Plant #1"},
                   {"code": "31311", "name": "Alabama Plant #2"}]:
            b64 = kpi_graphs.get(sp["code"])
            if b64:
                # KPI 이미지 압축
                try:
                    import io
                    from PIL import Image as _PILImg
                    kpi_bytes = base64.b64decode(b64)
                    kpi_img = _PILImg.open(io.BytesIO(kpi_bytes)).convert("RGB")
                    if kpi_img.width > 400:
                        ratio = 400 / kpi_img.width
                        kpi_img = kpi_img.resize((400, int(kpi_img.height * ratio)), _PILImg.LANCZOS)
                    buf = io.BytesIO()
                    kpi_img.save(buf, "JPEG", quality=60, optimize=True)
                    b64 = base64.b64encode(buf.getvalue()).decode()
                except Exception:
                    pass
                kpi_imgs_html += (
                    f'<td width="30%" style="text-align:center;padding:4px;">'
                    f'<div style="font-size:13px;font-weight:bold;color:#1F4E79;margin-bottom:6px;">'
                    f'Plant {sp["code"]} &ndash; {sp["name"]}</div>'
                    f'<img src="data:image/jpeg;base64,{b64}" '
                    f'style="width:100%;border:1px solid #ddd;border-radius:6px;" /></td>'
                )
        if kpi_imgs_html:
            kpi_html = (
                '<div style="margin-bottom:20px;max-width:800px;">'
                '<div style="background:#2E4057;color:#fff;padding:8px 12px;'
                'font-size:13px;font-weight:bold;">'
                '&#128202; Total Performance by Plant (KPI 현황)</div>'
                '<div style="border:1px solid #dee2e6;border-top:none;padding:8px;">'
                '<table width="100%" cellpadding="0" cellspacing="4" border="0"><tr>'
                + kpi_imgs_html +
                '</tr></table></div></div>'
            )

    # rates_img_b64는 이제 HTML 테이블 문자열
    rates_section = f'<div style="overflow-x:auto;margin-bottom:20px;">{rates_img_b64}</div>' if rates_img_b64 else '<p style="color:#888;">(Rates sheet data unavailable — please refer to the attached Excel file.)</p>'

    body = f"""<html><head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;color:#333;padding:20px;max-width:1400px;">

  <p>Hello All,</p>
  <p>Please refer to the list below to check for any delayed LPA/5S audits.
  Starting April 2025, for any audits exceeding the 3-day limit (Past Due Cases),
  the HPS team requires a statement of reason. Timely completion is critical,
  as it directly impacts our Plant Health Chart and overall operational efficiency.
  Please complete it on time. If you&#39;ve already done so, thank you.</p>
  <p><em>Note: Only 1st shift team members are used for notification purposes.
  If you are not the actual person in charge, please notify the relevant team member accordingly.</em></p>

  {kpi_html}

  {rates_section}

  <p style="font-size:12px;color:#888;border-top:1px solid #eee;padding-top:12px;">
    * This email is automatically generated by G-MES HPS RPA.<br>
    * Report period: {dr}<br>
    * Contact: jakyung.koo@hlcompany.com (OE Team)
  </p>
</body></html>"""

    subject = f"[HPS Audit Report] {dr} Plant Mobile Audit Report"

    # ── Step 4: 첨부파일 구성 (xlsx 1개) ─────────────────────────────
    attachments = []
    if report_path and report_path.exists():
        attachments.append(report_path)

    # ── Step 5: 이메일 발송 ──────────────────────────────────────────
    log.info(f"\n[EMAIL] 제목 : {subject}")
    log.info(f"[EMAIL] 수신 : {', '.join(EMAIL_TO)}")

    ok = send_via_outlook(
        subject     = subject,
        body_html   = body,
        to          = EMAIL_TO,
        cc          = EMAIL_CC if EMAIL_CC else None,
        bcc         = EMAIL_BCC if EMAIL_BCC else None,
        attachments = attachments or None,
        preview     = preview,
    )

    if ok and not preview:
        log.info("\n✅ 발송 완료!")
    elif not ok:
        log.info("\n❌ 이메일 발송 실패.")


# ══════════════════════════════════════════════════════════════════════
# ▌ 12. 오류 알림 메일
# ══════════════════════════════════════════════════════════════════════

def send_error_notification(exc: Exception):
    """오류 발생 시 로그 파일 내용을 DEVELOPER_EMAIL_TEST 로 발송"""
    dev = os.getenv("DEVELOPER_EMAIL_TEST", "")
    if not dev:
        return
    try:
        import win32com.client as win32
        import traceback
        # 로그 파일 읽기
        log_text = ""
        try:
            log_text = _LOG_FILE.read_text(encoding="utf-8", errors="replace")
        except Exception:
            log_text = "(로그 파일을 읽을 수 없습니다)"
        tb = traceback.format_exc()
        body = (
            "<html><body style='font-family:monospace;font-size:13px;'>"
            f"<h3 style='color:#cc0000;'>LPA/5S RPA 오류 발생</h3>"
            f"<p><b>오류:</b> {exc}</p>"
            f"<pre style='background:#f5f5f5;padding:12px;border:1px solid #ddd;'>{tb}</pre>"
            f"<hr/><h4>로그 전문</h4>"
            f"<pre style='background:#f0f0f0;padding:12px;border:1px solid #ddd;'>{log_text}</pre>"
            "</body></html>"
        )
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject  = f"[RPA ERROR] lpa_5s_combined_sender – {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        mail.HTMLBody = body
        mail.To       = dev
        mail.Send()
        log.info(f"[ERROR-NOTIFY] 오류 알림 발송 → {dev}")
    except Exception as e:
        log.warning(f"[ERROR-NOTIFY] 알림 발송 실패: {e}")


# ══════════════════════════════════════════════════════════════════════
# ▌ CLI
# ══════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="LPA + 5S 실적 100% 미만 필터링 후 Outlook 통합 이메일 발송",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            사용 예시:
              python lpa_5s_combined_sender.py              # G-MES 자동수집 + 발송
              python lpa_5s_combined_sender.py --preview    # Outlook 미리보기
              python lpa_5s_combined_sender.py --no-attach  # 첨부파일 없이 발송
        """)
    )
    parser.add_argument("--preview",   action="store_true",
                        help="Outlook 미리보기 창 열기 (발송 안 함)")

    args = parser.parse_args()
    try:
        run(preview=args.preview)
    except Exception as exc:
        log.error(f"[FATAL] 실행 중 오류 발생: {exc}", exc_info=True)
        send_error_notification(exc)
        sys.exit(1)


if __name__ == "__main__":
    main()

