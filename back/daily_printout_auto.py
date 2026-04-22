import os
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
from datetime import datetime, timedelta

# .env 파일 로드
load_dotenv()

# 변수 가져오기
my_id = os.getenv("GMES_ID")
my_pw = os.getenv("GMES_PW")


def log(message):
        now = datetime.now().strftime('%H:%M:%S')
        print(f"[{now}] {message}")

def click_menu(driver, wait, menu_en, menu_ko):
    """ 영문 메뉴명을 우선으로 찾고, 없으면 한글 메뉴명을 찾아 클릭하는 함수. a 태그와 div 태그 모두 대응 가능하도록 설계함."""

    # XPATH 설명: //*[...] -> 모든 태그 중에서 텍스트가 menu_en 혹은 menu_ko를 포함하는 요소를 찾음
    xpath = f"//*[(self::a or self::div) and (contains(text(), '{menu_en}') or contains(text(), '{menu_ko}'))]"
    
    try:
        # 1. 통합 XPATH로 요소 기다리기
        menu_element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        
        # 실제 어떤 이름으로 찾았는지 로그 출력
        found_text = menu_element.text
        menu_element.click()
        log(f"메뉴 클릭 성공: {found_text} (입력값: {menu_en} / {menu_ko})")
        
    except Exception as e:
        log(f"메뉴 클릭 실패: '{menu_en}' 또는 '{menu_ko}'를 찾을 수 없습니다.")
        raise e
    
def find_in_any_frame(driver, by, locator, timeout=10):
    """여러 iframe 중 해당 요소를 포함한 프레임을 찾아 진입 후 요소를 반환합니다."""
    # 1. 기본 페이지 내용으로 복귀
    driver.switch_to.default_content()
    
    # 2. 현재 페이지의 모든 iframe 찾기
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    wait = WebDriverWait(driver, timeout)

    # 3. 각 프레임을 하나씩 돌며 요소 찾기
    for frame in iframes:
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(frame)
            # 요소가 나타날 때까지 대기
            elem = wait.until(EC.presence_of_element_located((by, locator)))
            print(f"[{locator}] 요소를 프레임 내부에서 찾았습니다.")
            return elem
        except Exception:
            # 현재 프레임에 없으면 다음 프레임으로 이동
            continue

    # 4. 모든 프레임 탐색 후에도 못 찾은 경우 기본 페이지에서 재시도
    try:
        driver.switch_to.default_content()
        return wait.until(EC.presence_of_element_located((by, locator)))
    except:
        raise Exception(f"어느 iframe에서도 요소 {locator} 를 찾지 못했습니다.")

def close_alert_if_exists(driver, timeout=5):
    """ find_in_any_frame을 사용하여 특정 클래스와 텍스트를 가진 Confirm 버튼을 찾아 클릭합니다."""
    log("알림창(Confirm) 탐색 및 닫기 시도...")
    
    # 클래스가 ajs-ok인 버튼을 타겟으로 함
    xpath_confirm = "//button[contains(@class, 'ajs-ok')]"

    try:
        # 1. 일단 find_in_any_frame으로 프레임 위치를 잡습니다.
        # 찾기에 성공하면 드라이버는 이미 해당 프레임 안에 위치하게 됩니다.
        first_match = find_in_any_frame(driver, By.XPATH, xpath_confirm, timeout=timeout)
        
        if first_match:
            # 2. [핵심] 현재 위치한 프레임 내에서 동일한 XPATH를 가진 모든 버튼을 가져옵니다.
            all_btns = driver.find_elements(By.XPATH, xpath_confirm)
            target_btn = None

            for btn in all_btns:
                # 3. 실제로 눈에 보이는(Visible) 버튼을 찾습니다.
                if btn.is_displayed():
                    target_btn = btn
                    break

            # log(f"최초 버튼 확인 -> 가시성: {first_match.is_displayed()}, 텍스트: [{first_match.text}]")
            # log(f"최종 버튼 확인 -> 가시성: {target_btn.is_displayed()},  텍스트: [{target_btn.text}]")

            # 창닫기 실행.
            driver.execute_script("arguments[0].click();", target_btn)
            
            # 창이 사라지는 시간을 고려해 잠시 대기
            time.sleep(1)
            return True
            
    except Exception as e:
        log(f"알림창 처리 실패: {e}")
        return False

def set_calendar_date(driver, date_str):
    """ CminCalendar_input에 특정 날짜(YYYYMMDD)를 설정합니다. """

    log(f"날짜 설정 시작: {date_str}")
    try:
        # find_in_any_frame을 이용해 해당 요소가 있는 프레임으로 먼저 진입합니다.
        calendar_input = find_in_any_frame(driver, By.ID, "CminCalendar_input")
        
        if calendar_input:
            # 자바스크립트로 직접 값을 주입합니다. 
            # WebSquare 객체가 있을 경우 .setValue()를, 없을 경우 .value를 사용합니다.
            script = f"""
            var el = document.getElementById('CminCalendar_input');
            if (window.$p && $p.getComponentById('CminCalendar')) {{
                $p.getComponentById('CminCalendar').setValue('{date_str}');
            }} else {{
                el.value = '{date_str}';
            }}
            """
            driver.execute_script(script)
            log(f"날짜 설정 완료: {date_str}")
            time.sleep(0.5)
    except Exception as e:
        log(f"날짜 설정 중 오류 발생: {e}")

def click_pdf_print_button(driver, download_path=None):
    """ 새로 뜬 보고서 뷰어 창으로 전환하여 PDF 인쇄 버튼을 클릭하고 다운로드를 대기합니다. """
    
    # 0. 경로 설정 (None일 경우 기본 다운로드 폴더 지정)
    if download_path is None:
        download_path = os.path.join(os.path.expanduser("~"), "Downloads")
    
    log("보고서 뷰어 창 탐색 및 전환 시도...")
    
    try:
        # 1. 새 창(팝업)이 뜰 때까지 대기
        WebDriverWait(driver, 15).until(lambda d: len(d.window_handles) > 1)
        main_window = driver.current_window_handle  # 현재(메인) 창 저장
        
        # 뷰어 창으로 전환
        viewer_window = [w for w in driver.window_handles if w != main_window][0]
        driver.switch_to.window(viewer_window)
        
        log("뷰어 내부 데이터 로딩 대기 중...")
        wait = WebDriverWait(driver, 30)

        # 2. 로딩 레이어 대기
        try:
            # Crownix 로딩 레이어가 사라질 때까지 대기
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "crownix-viewer-wait-layer")))
            log("데이터 로딩 완료.")
        except:
            log("로딩 레이어 대기 중 타임아웃 혹은 이미 완료됨.")

        # 3. PDF 버튼 대기 및 클릭
        pdf_btn_xpath = "//li[@id='crownix-toolbar-print_pdf']//button"
        pdf_button = wait.until(EC.element_to_be_clickable((By.XPATH, pdf_btn_xpath)))
        
        # 다운로드 시작 전 파일 목록 스냅샷
        before_files = set(os.listdir(download_path))
        
        log("PDF 인쇄 버튼 클릭!")
        driver.execute_script("arguments[0].click();", pdf_button)

        # 4. 파일 다운로드 완료 모니터링
        log("파일 다운로드 감지 시작...")
        start_time = time.time()
        timeout = 60  # 최대 60초 대기
        target_file = None

        while time.time() - start_time < timeout:
            after_files = set(os.listdir(download_path))
            new_files = after_files - before_files
            
            if new_files:
                # 새로 생성된 파일 중 임시파일(.crdownload 등)이 아닌 .pdf 찾기
                pdf_files = [f for f in new_files if f.lower().endswith(".pdf")]
                if pdf_files:
                    target_file = pdf_files[0]
                    log(f"다운로드 감지됨: {target_file}")
                    break
            time.sleep(1)
        
        if not target_file:
            log("경고: 60초 내에 다운로드된 PDF 파일을 찾지 못했습니다.")

        # 5. 창 닫기 및 메인 창 복귀
        time.sleep(2) # 파일 저장 마무리 시간
        driver.close() # 현재 뷰어 창만 닫기
        driver.switch_to.window(main_window) # 반드시 메인으로 복귀해야 다음 루프 가능
        log("출력 프로세스 완료 및 메인 창 복귀.")
        
        return True

    except Exception as e:
        log(f"클릭/다운로드 중 오류 발생: {e}")
        # 오류 발생 시에도 안전하게 메인 창으로 복귀 시도
        try:
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[0])
        except:
            pass
        return False

    except Exception as e:
        log(f"뷰어 제어 중 오류 발생: {e}")
        # 오류 발생 시 안전하게 메인 창으로 복귀 시도
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[0])
        return False

def main():
    log("main 함수 시작")

    try:

        # 1. 크롬 브라우저 실행 설정
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        wait = WebDriverWait(driver, 10)
        
        # 2. GMES 3.0 접속 주소 (실제 접속 URL로 수정하세요)
        url = "https://gmesdev-mcaal.hlmando.com/" 
        driver.get(url)

        # 브라우저가 로딩될 때까지 잠시 대기
        time.sleep(5)

        # 3. 로그인 정보 입력
        id_input = driver.find_element(By.ID, "txtUSER_ID") 
        pw_input = driver.find_element(By.ID, "txtPWD") 

        id_input.clear()
        id_input.send_keys(my_id)
        pw_input.clear()
        pw_input.send_keys(my_pw)

        # 4. 로그인 버튼 클릭
        login_btn = driver.find_element(By.ID, "btnLogin")
        login_btn.click()

        # 로그인 후 화면이 넘어갈 때까지 대기
        log("로그인 시도 완료!")
        
        # 5. 메뉴이동 : POP -> Effort Mngt. -> Daily Production Report
        click_menu(driver, wait, "POP", "POP")                                 # 네비 버튼 : POP 카테고리 클릭
        click_menu(driver, wait, "Effort Mngt.", "공수관리")                    # 중메뉴 : Effort Mngt. 클릭
        click_menu(driver, wait, "Daily Production Report", "일일공정보고")      # 실제 타겟 메뉴 : Daily Production Report 클릭

        log("메뉴이동 성공!!")
        time.sleep(5)

        close_alert_if_exists(driver, timeout=10)               # 알림창 처리 단계 추가
        time.sleep(10)

        # 날짜 설정 단계 추가
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
        set_calendar_date(driver, "20250924")
        # set_calendar_date(driver, yesterday)

        # 프린트 실행을 위해 print 버튼도 미리 찾아놓기
        btn_print = find_in_any_frame(driver, By.ID, "btnPrint1")

        # 콤보박스 클릭해서 드롭다운 열기
        log("공장코드 선택시작!!")

        combo = find_in_any_frame(driver, By.ID, "cboSearchFactory")
        driver.execute_script("arguments[0].click();", combo)

        
        # 현재 열린 드롭다운에서 모든 옵션(td 태그)을 가져와 개수 파악
        options = driver.find_elements(By.XPATH, "//td[starts-with(@id, 'cboSearchFactory_itemTable_')]")
        option_count = len(options)
        log(f"총 {option_count}개의 공장 코드를 찾았습니다.")

        for i in range(option_count):
            options = driver.find_elements(By.XPATH, "//td[starts-with(@id, 'cboSearchFactory_itemTable_')]")
            opt = options[i]
            code_name = opt.text.strip()

            driver.execute_script("arguments[0].click();", opt)
            time.sleep(1)       # 선택 후 화면이 갱신될 시간 부여

            log(f"[{i+1}/{option_count}] '{code_name}' 보고서 생성 버튼 클릭")
            driver.execute_script("arguments[0].click();", btn_print)
            
            log("pdf 출력 시작!!")
            time.sleep(30)

            click_pdf_print_button(driver)

            log("pdf 출력 종료!!")
            time.sleep(5)

        log("테스트 성공!!!")

        time.sleep(5)

    except Exception as e:
        log(f"오류 발생: {e}")

    finally:
        time.sleep(10)          # 확인을 위해 10초 대기 후 종료
        # driver.quit()

if __name__ == "__main__":
    main()