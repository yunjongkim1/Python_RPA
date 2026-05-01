# daily_printout_automail.py

import os
import time
import sys  # 명령줄 인자를 받기 위해 추가.

from datetime import datetime, timedelta
from dotenv import load_dotenv  # 추가
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# dailyprintout.env 로드
if getattr(sys, 'frozen', False):
    _env_path = Path(sys.executable).parent / 'dailyprintout.env'
    _project_root = str(Path(sys.executable).parent.parent.parent)
else:
    _env_path = Path(__file__).parent / 'dailyprintout.env'
    _project_root = os.getenv("PROJECT_ROOT", str(Path(__file__).parent.parent.parent))
load_dotenv(_env_path)


# 프로젝트 루트 경로를 sys.path에 추가으켜서 core/ 모듈 임포트 가능게 설정
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)


# core 모듈에서 공통 함수와 로깅 기능 가져오기
from core.browser_config import win_open
from core.common_fn import (log, find_in_any_frame, close_alert_if_exists, click_pdf_print_button, 
                        set_calendar_date, send_mail_with_attachments, get_log_for_mail, safe_filename)

has_error = False


def main():
    start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log("----------------------------------------------------------------------------------")
    log(f"🚀 Daily Printout Automail 프로세스 시작 - [{start_time}]")
    log("----------------------------------------------------------------------------------")
    
    global has_error
    current_hour = datetime.now().hour
    down_dir = os.getenv("DOWNLOAD_DIR", r"C:\\RPA\\Download")
    pdf_down_dir = os.getenv("PDF_DOWNLOAD_DIR", r"C:\Users\yunjong.kim\Downloads")

    #--- 설정 영역 ---
#    url = os.getenv("GMES_ALDEV_URL")
    url = os.getenv("GMES_ALPROD_URL")  # 운영 URL로 변경 (2026-03-10)

    menus = [
        {"eng": "POP", "kor": "POP"},
        {"eng": "Effort Mngt.", "kor": "공수관리"},
        {"eng": "Daily Production Report", "kor": "일일공정보고"}
    ]
    work_date = (datetime.now() - timedelta(days=1 if current_hour < 12 else 0)).strftime('%Y%m%d')      # 12시 이전이면 1일 차감, 아니면 0일 차감

    # 시프트 리스트 고정 (A, B, C 순서대로)
    if    7 <= current_hour < 15: s_idx = 2; s_name = "Shift C"
    elif 15 <= current_hour < 23: s_idx = 0; s_name = "Shift A"
    else                        : s_idx = 1; s_name = "Shift B"

    # 메일 모드(test/prod)에 따라 수신자 로드 (.env의 PRINT_MAIL_MODE로 토글)
    _mode = os.getenv("PRINT_MAIL_MODE", "test").upper()
    mail_to = [v.split('#')[0].strip() for k, v in sorted(os.environ.items()) if k.startswith(f"MAIL_{_mode}_TO_") and v.strip()]
    mail_cc = [v.split('#')[0].strip() for k, v in sorted(os.environ.items()) if k.startswith(f"MAIL_{_mode}_CC_") and v.strip()]
    # DEVELOPER_EMAIL_1, DEVELOPER_EMAIL_2 ... 모두 읽기
    developer_email = [v.split('#')[0].strip() for k, v in sorted(os.environ.items()) if k.startswith("DEVELOPER_EMAIL") and v.strip()]
    
    log(f"📧 메일 모드: {_mode} | 수신자: {len(mail_to)}명 | 참조: {len(mail_cc)}명")
    final_pdf_files = []
    subject = f"[{work_date}] GMES Daily Production Report Automail"

    # 2. 외부 인자(add_email) 처리 로직 추가 ★
    if len(sys.argv) > 1:
        log(f"sys.argv: {sys.argv}")
        extra_email = sys.argv[1].strip()
        # 이메일 형식(@ 포함)인 경우에만 리스트에 추가
        if "@" in extra_email:
            # 중복 방지를 위해 리스트에 없을 때만 추가
            if extra_email not in mail_to:
                mail_to.append(extra_email)
                log(f"➕ 추가 수신자 포함됨: {extra_email}")

    try:
        #1. 브라우저 실행 및 메뉴 이동
        driver = win_open(url, menus)
        close_alert_if_exists(driver, timeout=3)

        #2. 날짜 설정 (WebSquare 객체 대응)
        set_calendar_date(driver, work_date)

        #3. 공장 목록 추출
        log("공장 목록 추출 중...")
        combo = find_in_any_frame(driver, By.ID, "cboSearchFactory")
        driver.execute_script("arguments[0].click();", combo)
        time.sleep(1)
        
        options = driver.find_elements(By.XPATH, "//td[starts-with(@id, 'cboSearchFactory_itemTable_')]")
        # 중요: 요소 객체가 아닌 '이름'만 리스트로 먼저 저장 (에러 방지)
        factory_names = [opt.get_attribute("innerText").strip() for opt in options]
    #    log(f"발견된 공장: {len(factory_names)}개 -> {factory_names}")

        # 요청에 의해 당분간 1공장 Plant1 만 처리하도록 변경
        factory_names = [name for name in factory_names if "31111" in name or "Alabama Plant 1" in name]
        log(f"처리 대상 공장: {len(factory_names)}개 -> {factory_names}")

        #3. 공장별 루프 실행
        btn_print = find_in_any_frame(driver, By.ID, "btnPrint1")
        
        for i, name in enumerate(factory_names):
            log(f"공장 loop 시작 : {name} ({i+1}/{len(factory_names)})")

            try:
                combo = find_in_any_frame(driver, By.ID, "cboSearchFactory")
                driver.execute_script("arguments[0].click();", combo)
                time.sleep(0.3) # 목록이 렌더링될 아주 짧은 시간

                target_xpath = f"//td[contains(@id, 'cboSearchFactory_itemTable') and contains(text(), '{name}')]"
                current_opt = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, target_xpath)))
                
                driver.execute_script("arguments[0].click();", current_opt)
                time.sleep(0.3) # 선택 후 화면 갱신 대기

            except Exception as e:
                log(f"공장 선택 중 오류 ({name}): {e}")
                has_error = True
                continue # 다음 공장으로 진행


            #4. 시프트 라디오 버튼 클릭 (ID 숫자를 직접 대입)
            shift_radio = find_in_any_frame(driver, By.ID, f"rdoSearchShift_input_{s_idx}")
            
            if shift_radio:
                driver.execute_script("arguments[0].click();", shift_radio)
                log(f"시프트 라디오 버튼을 선택합니다. {s_name} (공장: {name})")
                time.sleep(0.3) # 시프트 선택 후 화면 갱신 대기
            else:
                log(f"⚠️ 시프트 라디오 버튼을 찾을 수 없습니다: {s_name} (공장: {name})")
                has_error = True
                
            # 출력 및 이름 변경
            driver.execute_script("arguments[0].click();", btn_print)
            raw_file = click_pdf_print_button(driver, pdf_down_dir)
            # 프린트 파일에서 자동 생성된 파일명을 일자와 공장코드로 변경하면서 RPA 다운로드 폴더로 이동
            if raw_file:
                from shutil import move
                from pathlib import Path
                safe_name = safe_filename(work_date + "_" + name + "_" + s_name)
                new_file = Path(down_dir) / (safe_name + Path(raw_file).suffix)
                log(f"[DEBUG] move 시도: 원본={raw_file} (존재={os.path.exists(raw_file)}), 대상={new_file} (존재={os.path.exists(new_file)})")
                try:
                    move(raw_file, new_file)
                    final_pdf_files.append(str(new_file))
                    log(f"성공: {os.path.basename(new_file)}")
                except Exception as move_err:
                    log(f"[ERROR] 파일 이동 실패: {raw_file} → {new_file} | {move_err}")
                    has_error = True
            else:
                log(f"실패: {name} 리포트를 받지 못했습니다.")
                has_error = True

        #4. 메일 발송
        if final_pdf_files:
            file_names = "\n".join([os.path.basename(f) for f in final_pdf_files])
            body = (
                f"{subject}\n\n"f"--------------------------------------------------\n"
                f"✅ [Official Operation Notice]\n"
                f"The GMES RPA reporting system is now in official operation.\n"
                f"Status: Production Environment (Post-Pilot)\n"
                f"--------------------------------------------------\n\n"
                f"📍 Report Summary\n"
                f"- Target Factory: {name}\n"
                f"- Target Shift: {s_name}\n\n"
                f"📎 Attached Files:\n{file_names}\n\n"
                f"If you encounter any issues, please contact the administrator.\n\n"
                f"Best Regards,\n"
                f"Automated by RPA"
            )
            
            if _mode == "TEST":
                send_mail_with_attachments(final_pdf_files, developer_email, [], subject, body)
            else:
                send_mail_with_attachments(final_pdf_files, mail_to, mail_cc, subject, body)
    
    except Exception as e:
        log(f"❌ 치명적 오류: {e}")
        has_error = True
    finally:
        if has_error:
            error_log_content = get_log_for_mail()  # 메모리에 쌓인 로그 가져오기
            error_subject = f"🚨 [RPA 오류 알림] {subject} error"
            error_body = f"자동화 프로세스 중 오류가 발생했습니다.\n\n[실행 로그 요약]\n{error_log_content}"
        
            send_mail_with_attachments([], developer_email, [], error_subject, error_body)

        end_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log("----------------------------------------------------------------------------------")
        log(f"🏁 Daily Printout Automail 프로세스 종료 - [{end_time}]")
        log("----------------------------------------------------------------------------------")
if __name__ == "__main__":
    main()