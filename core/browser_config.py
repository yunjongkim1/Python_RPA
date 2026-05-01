# browser_config.py

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from core.common_fn import log, getID, getPW


def click_menu(wait, menu_en, menu_ko):
    """ 영문/한글 메뉴명을 찾아 클릭 """
    xpath = f"//*[(self::a or self::div) and (contains(text(), '{menu_en}') or contains(text(), '{menu_ko}'))]"
    try:
        menu_element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        found_text = menu_element.text
        menu_element.click()
    #    log(f"메뉴 클릭 성공: {found_text}")
    except Exception as e:
        log(f"메뉴 클릭 실패: '{menu_en}' 또는 '{menu_ko}'를 찾을 수 없습니다.")
        raise e


def move_to_specific_menu(driver, wait, menu_list):
    """ 대중소 메뉴를 순차적으로 클릭하여 이동 """
    try:
        log(f"mnu 1. 메뉴 이동 시작: {' > '.join([m['kor'] for m in menu_list])}")
        driver.switch_to.default_content() 
        
        for i, menu in enumerate(menu_list):
            click_menu(wait, menu["eng"], menu["kor"])
        #    log(f"   Step {i+1}: '{menu['kor']}' 클릭 완료")
            time.sleep(1.5) 

        log("mnu 2. 메뉴 이동 성공!!")
        return True
    except Exception as e:
        log(f"mnu 99. 메뉴 이동 중 오류 발생: {e}")
        return False


def win_open(url, target_menus):
    
    options = Options()
    options.add_argument('--headless')              # 화면 없이 실행
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    log("win 1. 새로운 크롬 브라우저를 실행합니다.")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)

    try:
        driver.get(url)
        log("win 2. 로그인 시도 중...")
        
        # ID/PW 입력 전 요소 대기 강화
        wait.until(EC.presence_of_element_located((By.ID, "txtUSER_ID"))).send_keys(getID())
        wait.until(EC.presence_of_element_located((By.ID, "txtPWD"))).send_keys(getPW())
        wait.until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()
        
        wait.until(EC.visibility_of_element_located((By.ID, "navbox")))     # navbox 요소가 DOM에 존재하고 화면에 보일 때까지 대기

        log("win 2-1. 로그인 완료.")

        # 3. 메뉴 이동 (이 부분부터 실행됨)
        if target_menus:
            move_to_specific_menu(driver, wait, target_menus)
            time.sleep(3)
            log("win 3. 메뉴 이동 프로세스 완료.")
        else :
            log("win 3-1. 메뉴 이동 필요 없음. 로그인 완료.")
            return driver
        
    except Exception as e:
        log(f"win 99. Windows Open작업 중 오류 발생: {e}")
    
    return driver