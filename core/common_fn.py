# common_fn.py

import os, time
import win32com.client as win32
import sys
import logging

import smtplib
from email.message import EmailMessage

from logging.handlers import RotatingFileHandler
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

_mail_log_buffer = []   # 전역 변수로 로그 메시지를 담을 리스트 선언

# 변수 가져오기
my_id = os.getenv("GMES_ID")
my_pw = os.getenv("GMES_PW")
ou_id = os.getenv("OUTLOOK_ID")
ou_pw = os.getenv("OUTLOOK_PW")


def getID():
    return my_id


def getPW():
    return my_pw


def log(message):
    """현재 사용자 다운로드 폴더 내 logs 폴더에 누적 기록"""
    logger = logging.getLogger("RPA")
    
    if not logger.hasHandlers():
        # 1. 로그 경로 설정 (.env의 LOG_DIR 우선, 없으면 USERPROFILE\Downloads\logs)
        log_dir = os.getenv("LOG_DIR") or os.path.join(os.environ['USERPROFILE'], 'Downloads', 'logs')
        os.makedirs(log_dir, exist_ok=True)
        
        # 2. 실행파일명_날짜.log 생성
        file_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
        log_path = os.path.join(log_dir, f"{file_name}_{datetime.now().strftime('%Y%m%d')}.log")

        # 3. 로거 설정
        handler = logging.FileHandler(log_path, encoding='utf-8')
        handler.setFormatter(logging.Formatter('[%(asctime)s] %(message)s', '%H:%M:%S'))
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)

    # 2. 콘솔 출력 및 파일 기록
    current_time = datetime.now().strftime('%H:%M:%S')
    print(f"[{current_time}] {message}")
    logger.info(message)
    _mail_log_buffer.append(f"[{current_time}] {message}")


def get_log_for_mail():
    """지금까지 쌓인 로그를 하나의 문자열로 반환"""
    return "\n".join(_mail_log_buffer)


def find_in_any_frame(driver, by, locator, timeout=5):
    """ 모든 iframe을 뒤져서 요소를 찾고 해당 프레임으로 전환 """
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    wait = WebDriverWait(driver, timeout)

    for frame in iframes:
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(frame)
            elem = wait.until(EC.presence_of_element_located((by, locator)))
            return elem
        except:
            continue
    
    try:
        driver.switch_to.default_content()
        return wait.until(EC.presence_of_element_located((by, locator)))
    except:
        raise Exception(f"요소 {locator}를 찾을 수 없습니다.")
    

def close_alert_if_exists(driver, timeout=5):
    """ 알림창(ajs-ok 클래스) 확인 후 클릭 """
    log("알림창(Confirm) 탐색 중...")
    xpath_confirm = "//button[contains(@class, 'ajs-ok')]"
    try:
        first_match = find_in_any_frame(driver, By.XPATH, xpath_confirm, timeout=timeout)
        if first_match:
            all_btns = driver.find_elements(By.XPATH, xpath_confirm)
            for btn in all_btns:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
                    return True
    except:
        return False


def check_and_close_system_alert(driver, target_text, timeout=5):
    """ 특정 텍스트가 포함된 알림창을 찾아서 로그를 남기고 '확인' 버튼을 클릭합니다. """
    try:
        # 1. 특정 텍스트를 포함한 요소 탐색 (데이터 없음, 타임아웃 등)
        xpath_msg = f"//*[contains(text(), '{target_text}')]"
        alert_msg = find_in_any_frame(driver, By.XPATH, xpath_msg, timeout=timeout)
        
        if alert_msg:
            msg_text = alert_msg.text.replace('\n', ' ').strip()
            log(f"⚠️ 시스템 알림 감지: [{msg_text}] -> 스킵합니다.")
            
            # 2. '확인' 버튼(ajs-ok 클래스 또는 '확인' 텍스트) 탐색 및 클릭
            xpath_confirm = "//button[contains(@class, 'ajs-ok') or contains(text(), '확인')]"
            confirm_btn = driver.find_element(By.XPATH, xpath_confirm)
            
            if confirm_btn.is_displayed():
                driver.execute_script("arguments[0].click();", confirm_btn)
                time.sleep(1)
                return True
    except:
        pass
    return False


def set_calendar_date(driver, date_str):
    """ 날짜 설정 (WebSquare 객체 대응) """
    log(f"날짜 설정 시도: {date_str}")
    try:
        calendar_input = find_in_any_frame(driver, By.ID, "CminCalendar_input")
        if calendar_input:
            script = f"""
            var el = document.getElementById('CminCalendar_input');
            if (window.$p && $p.getComponentById('CminCalendar')) {{
                $p.getComponentById('CminCalendar').setValue('{date_str}');
            }} else {{
                el.value = '{date_str}';
            }}
            """
            driver.execute_script(script)
            log(f"날짜 설정 완료")
            time.sleep(0.5)
    except Exception as e:
        log(f"날짜 설정 중 오류: {e}")


def wait_for_new_file(download_dir, timeout=60):
    """함수 호출 시점 이후에 생성된 새로운 PDF 파일만 대기"""

    start_time = time.time()
    while time.time() - start_time < timeout:
        # PDF 파일 목록 가져오기
        files = [os.path.join(download_dir, f) for f in os.listdir(download_dir) 
                 if f.endswith('.pdf')]
        
        if files:
            # 가장 최근 파일 찾기
            latest_file = max(files, key=os.path.getctime)
            ctime = os.path.getctime(latest_file)

            # 파일 생성 시간이 함수 호출 시간(start_time) 이후인 경우만 인정
            if ctime > start_time:
                # 다운로드 중인 임시 파일 확인
                if any(f.endswith(('.tmp', '.crdownload')) for f in os.listdir(download_dir)):
                    time.sleep(1)
                    continue
                
                # 파일 크기가 안정화될 때까지 아주 잠시 대기 (파일 쓰기 완료 확인)
                time.sleep(1) 
                return latest_file
                
        time.sleep(1)
    
    log("❌ 타임아웃: 새로운 파일을 찾을 수 없습니다.")
    return None


def file_rename(old_file_path, target_date, subplant):
    """ 파일명을 [날짜]_[subplant].pdf 형식으로 변경하되, 중복 시 seq 순으로 생성합니다. """
    try:
        old_path = Path(old_file_path)
        download_dir = old_path.parent
        
        # 파일명에서 특수문자 제거 (파일명으로 사용할 수 없는 문자 처리)
        clean_subplant = "".join(c for c in subplant if c.isalnum() or c in (' ', '_', '-')).strip()
        
        # 기본 파일명 설정
        base_name = f"{target_date}_{clean_subplant}"
        extension = ".pdf"
        
        # 중복 체크 및 순차 번호 부여
        final_path = download_dir / f"{base_name}{extension}"
        counter = 1
        
        while final_path.exists():
            # 중복될 경우: 날짜_공장(1).pdf 형식으로 생성
            final_path = download_dir / f"{base_name}({counter}){extension}"
            counter += 1
            
        # 실제 파일 이름 변경
        os.rename(old_path, final_path)
        return str(final_path)
        
    except Exception as e:
        print(f"파일명 변경 중 오류 발생: {e}")
        return old_file_path
 

def send_mail_with_attachments(attachment_paths, mail_to, mail_cc, subject, mail_body=None):
    """  mail_to, mail_cc가 리스트(Array) 형태일 경우 세미콜론으로 연결하여 메일을 발송합니다. """
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # 리스트인 경우 세미콜론(;)으로 연결, 문자열인 경우 그대로 사용
    final_to = "; ".join(mail_to) if isinstance(mail_to, list) else mail_to
    final_cc = "; ".join(mail_cc) if isinstance(mail_cc, list) else mail_cc

    mail.To = final_to
    mail.CC = final_cc
    mail.Subject = subject
    
    # 이메일 본문 구성 (첨부파일명 목록 추가 기능 포함)
    file_list_str = "\n".join([os.path.basename(p) for p in attachment_paths])
    mail.Body = mail_body
    
    # 파일 첨부
    for path in attachment_paths:
        if os.path.exists(path):
            mail.Attachments.Add(path)
        else:
            log(f"경고: 첨부할 파일을 찾을 수 없습니다: {path}")

    mail.Send()
    log("메일 발송 완료")


def send_smtpmail_with_attachments(attachment_paths, mail_to, mail_cc, subject, mail_body=None):
    """ 컨테이너 환경에서 SMTP를 사용하여 메일을 발송합니다. """
    
    # 1. 메일 서버 설정 (회사 IT팀에 확인 필요)
    SMTP_SERVER = "smtp.office365.com"  # Outlook/Office365 기준
    SMTP_PORT = 587
    SMTP_USER = ou_id   # 발송할 계정 이메일
    SMTP_PASSWORD = ou_pw      # 비밀번호 (혹은 앱 비밀번호)

    # 2. 메일 메시지 구성
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SMTP_USER
    msg['To'] = ", ".join(mail_to) if isinstance(mail_to, list) else mail_to
    msg['Cc'] = ", ".join(mail_cc) if isinstance(mail_cc, list) else mail_cc
    msg.set_content(mail_body)

    # 3. 파일 첨부
    for path in attachment_paths:
        if os.path.exists(path):
            with open(path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(path)
            msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)
        else:
            print(f"경고: 첨부할 파일을 찾을 수 없습니다: {path}")

    # 4. 실제 메일 발송
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # 보안 연결
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        print("메일 발송 완료 (SMTP)")
    except Exception as e:
        print(f"메일 발송 실패: {e}")
