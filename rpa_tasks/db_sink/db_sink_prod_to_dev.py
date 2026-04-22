import sys
import time
import pyodbc
import os
from pathlib import Path
from datetime import datetime, timedelta
from dotenv import load_dotenv

# .env 로드
load_dotenv()

# exec USP_COPY_INV_FROM_PROD 'G1RI02%1'

# 프로젝트 루트 경로를 sys.path에 추가하여 모듈 임포트 가능하게 설정
project_root = os.getenv("PROJECT_ROOT")
if project_root not in sys.path:
        sys.path.insert(0, project_root)

# core 모듈에서 공통 함수와 로깅 기능 가져오기
from core.common_fn import log, send_mail_with_attachments, get_log_for_mail

has_error = False
target_date = (datetime.now() - timedelta(days=90)).strftime('%Y%m%d')

# ---------------------------------------------------------
# 1. 개별 쿼리 정의 (SQL 원형 유지)
# ---------------------------------------------------------

def get_query_pop_discrete_jobs(ls):
    global target_date

    return f"""
    INSERT INTO POP_DISCRETE_JOBS 
    SELECT * FROM {ls}.GMES30.dbo.POP_DISCRETE_JOBS a WITH (NOLOCK) 
    WHERE a.WRK_DD >= '{target_date}'
    AND NOT EXISTS (
        SELECT 1 FROM POP_DISCRETE_JOBS x WITH (NOLOCK) 
        WHERE x.WRK_DD >= '{target_date}'
        AND x.WO_ID = a.WO_ID
    )
    """

def get_query_qms_mif_cnt(ls):
    global target_date

    return f"""
    INSERT INTO QMS_MIF_CNT_MINUTE (WRK_YMD, SYSUNIT, PLANT, SUBPLANT, WC, PLAN_RSC, TIME_BASE, PROGRESS_COUNT, WO_COUNT, REG_DT, REG_BY, UPD_DT, UPD_BY, LH_CNT, RH_CNT, COMMON_CNT)
    SELECT WRK_YMD, SYSUNIT, PLANT, SUBPLANT, WC, PLAN_RSC, TIME_BASE, PROGRESS_COUNT, WO_COUNT, REG_DT, REG_BY, UPD_DT, UPD_BY, LH_CNT, RH_CNT, COMMON_CNT 
    FROM {ls}.GMES30.dbo.QMS_MIF_CNT_MINUTE a WITH (NOLOCK) 
    WHERE a.WRK_YMD >= '{target_date}'
    AND NOT EXISTS (
        SELECT 1 FROM QMS_MIF_CNT_MINUTE x WITH (NOLOCK) 
        WHERE x.WRK_YMD >= '{target_date}'
        AND x.WRK_YMD = a.WRK_YMD 
        AND x.PLANT = a.PLANT 
        AND x.WC = a.WC 
        AND x.PLAN_RSC = a.PLAN_RSC 
        AND x.TIME_BASE = a.TIME_BASE
    )
    """

def get_query_qms_mif_cnt_shift(ls):
    global target_date

    return f"""
	INSERT	into QMS_MIF_CNT_SHIFT
	SELECT	*
	FROM	{ls}.GMES30.dbo.QMS_MIF_CNT_SHIFT a with (NOLOCK)
	WHERE	1=1
	AND	a.WRK_YMD >= '{target_date}'
	AND	not exists(
		SELECT	*
		FROM	QMS_MIF_CNT_SHIFT x with (NOLOCK)
		WHERE	x.WRK_YMD >= '{target_date}'
		AND	x.WRK_YMD = a.WRK_YMD
		AND	x.WC = a.WC
		AND	x.SHIFTID = a.SHIFTID
    );

    UPDATE a 
    SET a.END_COUNTER_TIME = b.END_COUNTER_TIME
      , a.RELAX1_END_COUNT_TIME = b.RELAX1_END_COUNT_TIME
      , a.RELAX2_END_COUNT_TIME = b.RELAX2_END_COUNT_TIME
      , a.RELAX3_END_COUNT_TIME = b.RELAX3_END_COUNT_TIME
      , a.PROGRESS_COUNT = b.PROGRESS_COUNT
      , a.RELAX1_MID_START_COUNT_TIME = b.RELAX1_MID_START_COUNT_TIME
      , a.RELAX1_MID_END_COUNT_TIME = b.RELAX1_MID_END_COUNT_TIME
      , a.RELAX2_MID_START_COUNT_TIME = b.RELAX2_MID_START_COUNT_TIME
      , a.RELAX2_MID_END_COUNT_TIME = b.RELAX2_MID_END_COUNT_TIME
      , a.RELAX3_MID_START_COUNT_TIME = b.RELAX3_MID_START_COUNT_TIME
      , a.RELAX3_MID_END_COUNT_TIME = b.RELAX3_MID_END_COUNT_TIME
    FROM dbo.QMS_MIF_CNT_SHIFT a
    INNER JOIN {ls}.GMES30.dbo.QMS_MIF_CNT_SHIFT b ON
          a.WRK_YMD = b.WRK_YMD AND a.SHIFTID = b.SHIFTID AND a.WC = b.WC
    WHERE a.WRK_YMD >= '{target_date}';
    """

def get_query_report_header(ls):
    return f"""
    SET IDENTITY_INSERT POP_DAILY_PROD_REPORT_HEADER ON;
    INSERT INTO POP_DAILY_PROD_REPORT_HEADER (
        PR_SEQ, SUBPLANT, WC, SHIFT_ID, WRK_DD, ATTRIBUTE01, ATTRIBUTE02, 
        ATTRIBUTE03, ATTRIBUTE04, ATTRIBUTE05, REG_BY, REG_DT, UPD_DT, UPD_BY
    ) 
    SELECT PR_SEQ, SUBPLANT, WC, SHIFT_ID, WRK_DD, ATTRIBUTE01, ATTRIBUTE02, 
           ATTRIBUTE03, ATTRIBUTE04, ATTRIBUTE05, REG_BY, REG_DT, UPD_DT, UPD_BY 
    FROM {ls}.GMES30.dbo.POP_DAILY_PROD_REPORT_HEADER a 
    WHERE NOT EXISTS (
        SELECT 1 FROM POP_DAILY_PROD_REPORT_HEADER x WITH (NOLOCK) 
        WHERE x.PR_SEQ = a.PR_SEQ
    );
    SET IDENTITY_INSERT POP_DAILY_PROD_REPORT_HEADER OFF;
    """

def get_query_report_line(ls):
    return f"""
    INSERT INTO POP_DAILY_PROD_REPORT_LINE 
    SELECT * FROM {ls}.GMES30.dbo.POP_DAILY_PROD_REPORT_LINE a 
    WHERE NOT EXISTS (
        SELECT 1 FROM POP_DAILY_PROD_REPORT_LINE x WITH (NOLOCK) 
        WHERE x.PR_SEQ = a.PR_SEQ AND x.CLASS = a.CLASS AND x.IDX = a.IDX
    )
    """


# ---------------------------------------------------------
# 2. 실행 엔진: 단일 서버 마이그레이션 수행 (10분 타임아웃)
# ---------------------------------------------------------

def run_migration(srv_config, query_funcs, timeout_limit=300):
    """ 주입받은 단일 서버 설정과 쿼리 리스트를 바탕으로 마이그레이션 수행 """
    driver = "{ODBC Driver 17 for SQL Server}"
    start_time = time.time()
    global has_error
    
    conn_str = (
        f"DRIVER={driver};SERVER={srv_config['ip']};DATABASE={srv_config['db']};"
        f"UID={srv_config['user']};PWD={srv_config['pw']};Connection Timeout=30;"
    )

    try:
        with pyodbc.connect(conn_str, autocommit=False) as conn:
            cursor = conn.cursor()
            
            for func in query_funcs:
                # 타임아웃 체크
                if (time.time() - start_time) > timeout_limit:
                    log(f"{srv_config['name']}: !!! TIMEOUT !!! 10분 초과로 해당 서버 작업 중단")
                    has_error = True
                    return False            # 타임아웃으로 인한 중단 알림

                q_name = func.__name__
                sql = func(srv_config['ls_name'])
                
                try:
                    log(f"{srv_config['name']}: {q_name} 실행 중")
                    cursor.execute(sql)
                    conn.commit()
                    log(f"{srv_config['name']}: {q_name} 완료 ({cursor.rowcount}건)")
                    time.sleep(1)           # 운영 서버 보호를 위한 미세 지연
                except Exception as e:
                    conn.rollback()
                    log(f"{srv_config['name']}: {q_name} 실패, 다음 단계 중단: {e}")
                    has_error = True
                    return False            # 단계 실패 시 해당 서버 후속 작업 중단

            return True                     # 모든 쿼리 성공 완료

    except Exception as e:
        log(f"{srv_config['name']}: 접속 오류: {e}")
        has_error = True
        return False

# ---------------------------------------------------------
# 3. 메인 컨트롤러: 전체 흐름 제어 및 설정 주입
# ---------------------------------------------------------

def main():
    # 서버 설정 정의 (직렬화된 구조 유지)
    configs = [
        {
            "name": "GA",
            "ip": "172.28.18.81",
            "db": "GMES30",
            "user": "mcait",
            "pw": "Mando0601",
            "ls_name": "GMES30PROD"
        },
        {
            "name": "AL",
            "ip": "172.28.10.90",
            "db": "GMES30",
            "user": "mcait",
            "pw": "Mando0601",
            "ls_name": "GMES30PROD"
        }
    ]

    # 타임아웃 설정 (초 단위, 예: 600초 = 10분)
    timeout_limit = 300

    # 실행할 쿼리 함수 목록 (순서대로)
    target_queries = [
        get_query_pop_discrete_jobs,
        get_query_qms_mif_cnt,
        get_query_qms_mif_cnt_shift,
        get_query_report_header,
        get_query_report_line
    ]

    # 개발자 이메일
    developer_email = ["yunjong.kim@hlcompany.com"]
    global has_error
    global target_date

    log(f"target_date : {target_date}")
    
    try:
        for idx, srv in enumerate(configs):
            # 1. 서버 간 대기 시간 제어 (main의 역할)
            if idx > 0:
                log(f"SYSTEM: {srv['name']} 작업 전 운영 서버 보호를 위해 {timeout_limit}초 대기 중...")
                time.sleep(timeout_limit)
            
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log(f"[{now}] SYSTEM: [{srv['name']}] 서버 마이그레이션 프로세스 시작")
            
            # 2. 실행 엔진 호출 (run_migration의 역할)
            success = run_migration(srv, target_queries, timeout_limit)
            
            if success:
                log(f"SYSTEM: [{srv['name']}] 모든 마이그레이션 단계 성공 완료")
            else:
                log(f"SYSTEM: [{srv['name']}] 마이그레이션 도중 중단됨 (오류 또는 타임아웃)")
                has_error = True

        log("SYSTEM: 모든 대상 서버에 대한 작업이 종료되었습니다.")
        
        # 3. 최종 결과 요약 및 메일 발송 (main의 역할)
        subject_base = f"운영DB 데이터 이관 작업 ({time.strftime('%Y-%m-%d')})"
        log_content = get_log_for_mail() # 메모리에 쌓인 전체 로그 가져오기
    except Exception as e:
        log(f"SYSTEM: 예기치 못한 오류 발생: {e}")
        has_error = True

    finally:
        if has_error:
            mail_subject = f"🚨 [오류 발생] {subject_base}"
            mail_body = f"자동화 프로세스 중 오류 또는 타임아웃이 발생했습니다.\n확인이 필요합니다.\n\n[실행 로그 요약]\n{log_content}"
        else:
            mail_subject = f"✅ [완료] {subject_base}"
            mail_body = f"모든 서버의 데이터 이관 작업이 성공적으로 완료되었습니다.\n\n[실행 로그 요약]\n{log_content}"
        
        send_mail_with_attachments([], developer_email, [], mail_subject, mail_body)
if __name__ == "__main__":
    main()