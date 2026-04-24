# HL Mando REST API Server
# rest_api_server.py

import os
import subprocess
import uvicorn
import psutil
import asyncio
import threading
from contextlib import asynccontextmanager
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
from pathlib import Path

# rest_api.env 로드 (app/ 폴더 기준)
load_dotenv(Path(__file__).parent / 'rest_api.env')

# --- [Auth] API 키 인증 ---
_api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)

def verify_api_key(api_key: str = Security(_api_key_header)):
    """X-API-Key 헤더 값을 .env의 API_KEY와 비교하여 인증합니다."""
    expected = os.getenv("API_KEY", "")
    if not expected:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="API_KEY가 서버에 설정되지 않았습니다.")
    if not api_key or api_key != expected:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="유효하지 않은 API 키입니다.")

# --- [Lifespan] 서버 시작/종료 이벤트 ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    # startup
    setup_automation_schedule()
    if not scheduler.running:
        scheduler.start()
    scheduler.add_job(
        print_job_status,
        trigger=CronTrigger(hour=5, minute=0),
        id="JOB_STATUS",
        replace_existing=True
    )
    print_job_status()
    yield
    # shutdown
    scheduler.shutdown()
    print("🛑 스케줄러가 정지되었습니다.")

app = FastAPI(
    title="HL Mando REST API Server",
    description="사내 업무 자동화 통합 API 서버",
    lifespan=lifespan
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

def _run_job_with_log(sp, name):
    """EXE/PY 실행 후 stdout을 서버 콘솔에 실시간 출력 (JOB_SILENT에 포함된 잡은 시작/종료만 출력)"""
    silent_jobs = [j.strip().upper() for j in os.getenv("JOB_SILENT", "").split(",") if j.strip()]
    is_silent = name.upper() in silent_jobs

    def _stream():
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        cmd = make_cmd(sp)
        print(f"\n▶ [{name}] 시작: {' '.join(cmd)}")
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE if not is_silent else subprocess.DEVNULL,
                stderr=subprocess.STDOUT if not is_silent else subprocess.DEVNULL,
                env=env
            )
            if not is_silent:
                for raw in proc.stdout:
                    line = raw.decode('utf-8', errors='replace').rstrip()
                    if line:
                        print(f"  [{name}] {line}")
            proc.wait()
            status = "✅ 성공" if proc.returncode == 0 else f"❌ 실패 (exit code: {proc.returncode})"
            print(f"■ [{name}] 종료: {status}")
        except Exception as e:
            print(f"  [{name}] ❌ 실행 오류: {e}")
    threading.Thread(target=_stream, daemon=True).start()

def make_cmd(sp):
    """ .exe는 직접 실행, .py는 python으로 실행 """
    if sp.endswith(".exe"):
        return [sp]
    return ["python", sp]

def setup_automation_schedule():
    """
    .env의 JOB_SCHEDULES를 읽어 APScheduler에 등록합니다.
    형식: 이름:요일:시:분:스크립트경로 (예: A:mon-fri:15:31:rpa_tasks/daily_mail/daily_printout_automail.py)
    새 Job 추가 시 .env의 JOB_SCHEDULES에 ,구분자로 항목 추가. EXE 재빌드 불필요.
    """
    raw_schedules = os.getenv("JOB_SCHEDULES", "")
    if not raw_schedules:
        print("   ⚠️ JOB_SCHEDULES 환경변수가 설정되지 않았습니다.")
        return
        
    for item in raw_schedules.splitlines():
        if not item.strip():   # 빈 줄 스킵
            continue
        try:
            # 최대 4번 split → [이름, 요일, 시, 분, 스크립트경로] 5개
            parts = item.strip().split(":", 4)
            if len(parts) != 5:
                print(f"   ⚠️ 형식 오류 (이름:요일:시:분:스크립트경로): {item}")
                continue

            name, days, h, m, script_path = parts

            # JOB_SCRIPT_DIR이 설정된 경우 경로 조합
            # 단, script_path에 이미 폴더 구분자가 있으면 그대로 사용 (TRANS 등 별도 폴더 잡)
            script_dir = os.getenv("JOB_SCRIPT_DIR", "").strip()
            if script_dir and '/' not in script_path and '\\' not in script_path:
                full_path = f"{script_dir}/{script_path}"
            else:
                full_path = script_path

            # days 형식: dom-N → 매달 N일, 그 외 → 요일(mon-fri 등)
            if days.startswith("dom-"):
                day_of_month = int(days.split("-")[1])
                trigger = CronTrigger(
                    day=day_of_month,
                    hour=int(h),
                    minute=int(m),
                    timezone=os.getenv("APP_TIMEZONE", "America/Chicago")
                )
            else:
                trigger = CronTrigger(
                    day_of_week=days,
                    hour=int(h),
                    minute=int(m),
                    timezone=os.getenv("APP_TIMEZONE", "America/Chicago")
                )

            scheduler.add_job(
                lambda sp=full_path, n=name: _run_job_with_log(sp, n),
                trigger=trigger,
                id=f"JOB_{name}",
                replace_existing=True,
                misfire_grace_time=1800     #1800초=30분 동안 지연되어도 실행함
            )
            print(f"   [Registered] {name}: {days} {h}:{m} -> {script_path}")
        except Exception as e:
            print(f"   ⚠️ JOB_SCHEDULES 형식 오류 ({item}): {e}")

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

            print(f" 📌 {job.id.replace('JOB_', ''):12} | 예정: {next_run.strftime('%m/%d %H:%M')} | 남은 시간: {rem_hours}시간 {display_min}분 후")
        else:
            print(f" 📌 {job.id.replace('JOB_', ''):12} | 예정: 없음 (스케줄 확인 필요)")
    print("="*70 + "\n")

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

        cmd = ["python", "rpa_tasks/dailyprintout/daily_printout_automail.py"]
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
        cmd = ["python", "-u", "rpa_tasks/dailyprintout/daily_printout_automail.py"]
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

@app.get("/run/db_sink/stream", tags=["Action"])
async def run_db_sink_stream(_: None = Security(verify_api_key)):
    async def log_generator():
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['cmdline'] and 'db_sink_prod_to_dev.py' in proc.info['cmdline']:
                yield f"data: ❌ Error: Another process is already running.\n\n"
                return

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        cmd = ["python", "-u", "rpa_tasks/dailyprintout/db_sink_prod_to_dev.py"]

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
