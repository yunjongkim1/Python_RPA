# HL Mando REST API Server
# rest_api_server.py

import os
import subprocess
import uvicorn
import psutil
import asyncio
from datetime import datetime

from dotenv import load_dotenv
from fastapi import FastAPI, Query, Security, HTTPException, status
from fastapi.responses import RedirectResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.security import APIKeyHeader
from urllib.parse import quote

# 스케줄러 라이브러리 추가
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

# .env 로드
load_dotenv()

# --- [Auth] API 키 인증 ---
_api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)

def verify_api_key(api_key: str = Security(_api_key_header)):
    """X-API-Key 헤더 값을 .env의 API_KEY와 비교하여 인증합니다."""
    expected = os.getenv("API_KEY", "")
    if not expected:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="API_KEY가 서버에 설정되지 않았습니다.")
    if not api_key or api_key != expected:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="유효하지 않은 API 키입니다.")

app = FastAPI(
    title="HL Mando REST API Server",
    description="사내 업무 자동화 통합 API 서버"
)

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- [Scheduler] 스케줄러 설정 ---
scheduler = BackgroundScheduler(daemon=True)

def setup_automation_schedule(env_key, script_path, job_id_prefix):
    """
    .env의 스케줄 설정을 읽어와서 APScheduler에 등록합니다.
    형식: NAME:DAYS:TIME (예: C:tue-sat:07:30)
    """
    raw_schedules = os.getenv(env_key, "")
    if not raw_schedules:
        return
        
    for item in raw_schedules.split(","):
        try:
            # 최대 2번만 나눠서 [이름, 요일, 시간] 3개로 만듭니다.
            parts = item.strip().split(":", 2) 
            if len(parts) != 3:
                print(f"   ⚠️ 형식 오류 (시프트:요일:시간): {item}")
                continue

            name, days, time_val = parts
            h, m = time_val.split(":")
            
            # 중복 실행 방지 로직이 포함된 subprocess 실행
            # n=name은 람다 캡처 문제를 피하기 위해 사용
            scheduler.add_job(
                lambda n=name: subprocess.Popen(["python", script_path, n]),
                trigger=CronTrigger(
                    day_of_week=days, 
                    hour=int(h), 
                    minute=int(m), 
                    timezone=os.getenv("APP_TIMEZONE", "America/Chicago")
                ),
                id=f"{job_id_prefix}_{name}",
                replace_existing=True,
                misfire_grace_time=1800     #1800초=30분 동안 지연되어도 실행함
            )
            print(f"   [Registered] {job_id_prefix} - Shift {name}: {days} {h}:{m}")
        except Exception as e:
            print(f"   ⚠️ {env_key} 스케줄 형식 오류 ({item}): {e}")

# 서버 시작 이벤트 수정
@app.on_event("startup")
def start_all_schedules():
    # 기존 자동화 스케줄 등록
    setup_automation_schedule("DAY_RPT_AUTO_SCHEDULES", "rpa_tasks/daily_mail/daily_printout_automail.py", "RPT")
    setup_automation_schedule("DB_SINK_AUTO_SCHEDULES", "rpa_tasks/db_sink/db_sink_prod_to_dev.py", "SINK")

    if not scheduler.running:
        scheduler.start()

    # [추가] 매일 새벽 00:00(또는 원하는 시간)에 작업 현황을 출력하도록 예약
    scheduler.add_job(
        print_job_status,
        trigger=CronTrigger(hour=5, minute=0), # 매일 자정에 실행
        id="JOB_STATUS",
        replace_existing=True
    )

    # 서버 켤 때도 한 번 보고 싶으므로 즉시 호출
    print_job_status()

# 출력 로직을 별도의 함수로 분리
def print_job_status():
    """현재 예약된 모든 작업의 리스트와 남은 시간을 출력하는 함수"""
    print("\n" + "="*70)
    print(f" 🔍 [정기 스케줄 점검] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("-"*70)

    jobs = scheduler.get_jobs()
    if not jobs:
        print(" 📌 등록된 예약 작업이 없습니다.")
    
    for job in jobs:
        next_run = job.next_run_time
        if next_run:
            now = datetime.now(next_run.tzinfo) 
            remaining = next_run - now
            
            rem_minutes = int(remaining.total_seconds() / 60)
            rem_hours = rem_minutes // 60
            display_min = rem_minutes % 60

            print(f" 📌 작업: {job.id:10} | 예정: {next_run.strftime('%m/%d %H:%M')} | 남은 시간: {rem_hours}시간 {display_min}분 후")
        else:
            print(f" 📌 작업: {job.id:10} | 예정: 없음 (스케줄 확인 필요)")
    print("="*70 + "\n")

# 서버 종료 시 스케줄러도 함께 종료
@app.on_event("shutdown")
def stop_scheduler():
    scheduler.shutdown()
    print("🛑 스케줄러가 정지되었습니다.")

# --- [UI] 엔드포인트 ---
@app.get("/", tags=["UI"])
async def main_page():
    if os.path.exists('app/web/index.html'):
        return FileResponse('app/web/index.html')
    return {"message": "REST API Server is running."}

@app.get("/view/daily_report", tags=["UI"])
async def view_daily_report_direct_page():
    path = 'app/web/daily_report.html'
    if os.path.exists(path):
        return FileResponse(path)
    return {"error": "File not found"}

# --- [Report] 리포트 호출 전용 API ---
@app.get("/report/daily_prod_report_direct", tags=["Production Report"])
async def get_daily_prod_report_direct(
        work_date: str = Query(..., description="작업일자 (YYYYMMDD)"),
        shift: str = Query(..., description="근무조 (A/B/C)"),
        sub_plant: str = Query(..., description="공장코드"),
        s_day: str = Query(..., description="표시날짜 (MM/DD/YYYY)"),
        wc: str = Query("", description="작업장 코드"),
        direct: bool = Query(True, description="즉시 리다이렉트 여부"),
        _: None = Security(verify_api_key)
    ):
    base_url = os.getenv("REPORT_ALDEV_URL")
    mrd_path = "/reportservice/pop/en/DailyProdReport.mrd"
    
    def build_crownix_url(base, path, w_date, sft, plant, center, day):
        rv_param = f"/rv WORKDATE[{w_date}] SHIFT[{sft}] SUBPLANT[{plant}] WC[{center}] sDay[{day}]"
        return f"{base}?mrdUrl={quote(path)}&rvParam={quote(rv_param)}"

    final_url = build_crownix_url(base_url, mrd_path, work_date, shift, sub_plant, wc, s_day)

    if direct:
        return RedirectResponse(url=final_url)
    return {"status": "success", "target": "DailyProdReport", "url": final_url}

# --- [Action] 수동 실행 및 스트리밍 ---
@app.get("/run/daily_report_automail", tags=["Action"])
async def run_daily_report_automail(extra_email: str = Query(None), _: None = Security(verify_api_key)):
    try:
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['cmdline'] and 'daily_printout_automail.py' in proc.info['cmdline']:
                return {"status": "error", "message": "Another process is already running."}

        cmd = ["python", "rpa_tasks/daily_mail/daily_printout_automail.py"]
        if extra_email:
            cmd.append(extra_email)

        subprocess.Popen(cmd)
        return {"status": "success", "message": "Process started in background."}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/run/daily_report_automail/stream", tags=["Action"])
async def run_daily_report_automail_stream(extra_email: str = Query(None), _: None = Security(verify_api_key)):
    async def log_generator():
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['cmdline'] and 'daily_printout_automail.py' in proc.info['cmdline']:
                yield f"data: ❌ Error: Another process is already running.\n\n"
                return
        
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        cmd = ["python", "-u", "rpa_tasks/daily_mail/daily_printout_automail.py"]
        if extra_email:
            cmd.append(extra_email)

        process = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.STDOUT,
            env=env
        )

        yield f"data: Process Started...\n\n"
        while True:
            line = await process.stdout.readline()
            if not line: break
            text = line.decode('utf-8', errors='replace').strip()
            if text: yield f"data: {text}\n\n"
        
        await process.wait()
        yield f"data: Done.\n\n"

    return StreamingResponse(log_generator(), media_type="text/event-stream")

if __name__ == "__main__":
    import socket
    
    host = os.getenv("SERVER_HOST", "0.0.0.0")
    port = int(os.getenv("SERVER_PORT", 8000))
    
    # .env에 설정값이 있으면 그걸 쓰고, 없으면 현재 장비의 IP를 자동으로 찾아옵니다.
    env_ip = os.getenv("MY_COMPUTER_IP")
    if not env_ip or env_ip == "127.0.0.1":
        try:
            # 현재 실행 중인 장비(서버 또는 PC)의 실제 IP 주소를 가져오는 코드
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            display_ip = s.getsockname()[0]
            s.close()
        except Exception:
            display_ip = "localhost"
    else:
        display_ip = env_ip

    print(f"\n" + "="*60)
    print(f"   🚀 HL Mando RPA API Portal Server 가동 중")
    print(f"   🌐 접속 주소: http://{display_ip}:{port}")
    print(f"   🏠 환경: {'Container/Server' if os.path.exists('/.dockerenv') else 'Local Windows'}")
    print("="*60)
    
    uvicorn.run(app, host=host, port=port)
