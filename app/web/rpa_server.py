# HL Mando REST API Server
# rpa_server.py

import os
import subprocess
import sys
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

# rpa_server.env 파일 경로 (env 수정 시 필요)
_ENV_FILE = Path(__file__).parent / 'rpa_server.env'
_HISTORY_FILE = Path(__file__).parent / 'job_history.json'

# rpa_server.env 로드 (app/ 폴더 기준)
load_dotenv(_ENV_FILE)

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
    _load_history()
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
    _save_history()
    print("🛑 스케줄러가 정지되었습니다.")

import logging
import time
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

# --- [Logging] 접속 로그 설정 ---
_log_formatter = logging.Formatter("%(asctime)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
_log_file_handler = logging.FileHandler(
    Path(__file__).parent.parent.parent / "Logs" / "access.log",
    encoding="utf-8"
)
_log_file_handler.setFormatter(_log_formatter)
_log_stream_handler = logging.StreamHandler()
_log_stream_handler.setFormatter(_log_formatter)

# uvicorn.error만 포맷 적용, uvicorn.access는 비활성화 (AccessLogMiddleware가 대신 처리)
for _logger_name in ("uvicorn", "uvicorn.error"):
    _uv_logger = logging.getLogger(_logger_name)
    _uv_logger.handlers.clear()
    _uv_logger.addHandler(_log_stream_handler)
    _uv_logger.addHandler(_log_file_handler)
    _uv_logger.propagate = False

logging.getLogger("uvicorn.access").disabled = True

access_logger = logging.getLogger("access")
access_logger.setLevel(logging.DEBUG)
access_logger.addHandler(_log_stream_handler)
access_logger.addHandler(_log_file_handler)
access_logger.propagate = False

app = FastAPI(
    title="HL Mando REST API Server",
    description="사내 업무 자동화 통합 API 서버",
    lifespan=lifespan
)

class AccessLogMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        start = time.time()
        response = await call_next(request)
        elapsed = int((time.time() - start) * 1000)
        client_ip = request.headers.get("x-forwarded-for", request.client.host if request.client else "-")
        # /api/status 는 30초마다 폴링되므로 별도 레벨로 축약
        if request.url.path == "/api/status":
            access_logger.debug(f"{client_ip} {request.method} {request.url.path} {response.status_code} {elapsed}ms")
        else:
            access_logger.info(f"{client_ip} {request.method} {request.url.path} {response.status_code} {elapsed}ms")
        return response

app.add_middleware(AccessLogMiddleware)

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

# job별 실행 이력 저장 {name: [{"time": str, "status": str, "exit_code": int}, ...]}
JOB_HISTORY_MAX = 20  # job별 최대 보관 건수
job_history: dict = {}
running_jobs: set = set()  # 현재 실행 중인 job 이름 집합
job_scripts: dict = {}    # job 이름 → 스크립트 경로 (psutil 조회용)

def _load_history():
    """job_history.json 가 있으면 로드"""
    global job_history
    if _HISTORY_FILE.exists():
        try:
            import json
            job_history = json.loads(_HISTORY_FILE.read_text(encoding='utf-8'))
            print(f"   💾 실행 이력 로드: {_HISTORY_FILE}")
        except Exception as e:
            print(f"   ⚠️ 이력 파일 로드 실패: {e}")

def _save_history():
    """job_history를 JSON 파일로 저장"""
    try:
        import json
        _HISTORY_FILE.write_text(json.dumps(job_history, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception as e:
        print(f"   ⚠️ 이력 파일 저장 실패: {e}")

def _get_server_ip() -> str:
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "unknown"

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
        run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        running_jobs.add(name)
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE if not is_silent else subprocess.DEVNULL,
                stderr=subprocess.STDOUT if not is_silent else subprocess.DEVNULL,
                env=env
            )
            if not is_silent:
                output_lines = []
                for raw in proc.stdout:
                    line = raw.decode('utf-8', errors='replace').rstrip()
                    if line:
                        print(f"  [{name}] {line}")
                        output_lines.append(line)
            proc.wait()
            # exit code 우선, exit code 0이어도 stdout에 실패 키워드 있으면 실패로 판정
            _fail_keywords = ["❌", "실패", "오류", "error", "exception", "timeout", "타임아웃"]
            _output_text = "\n".join(output_lines).lower() if not is_silent else ""
            _has_fail_keyword = any(k in _output_text for k in _fail_keywords)
            if proc.returncode != 0:
                status_str = f"❌ 실패 (exit code: {proc.returncode})"
            elif _has_fail_keyword:
                # 실패 키워드가 포함된 마지막 줄 표시
                _fail_line = next((l for l in reversed(output_lines) if any(k in l.lower() for k in _fail_keywords)), "")
                status_str = f"❌ 실패: {_fail_line.strip()[:60]}" if _fail_line else "❌ 실패 (출력에 오류 감지)"
            else:
                status_str = "✅ 성공"
            record = {"time": run_time, "status": status_str, "exit_code": proc.returncode}
            if name not in job_history:
                job_history[name] = []
            job_history[name].insert(0, record)
            job_history[name] = job_history[name][:JOB_HISTORY_MAX]
            _save_history()
            print(f"■ [{name}] 종료: {status_str}")
        except Exception as e:
            record = {"time": run_time, "status": f"❌ 오류: {e}", "exit_code": -1}
            if name not in job_history:
                job_history[name] = []
            job_history[name].insert(0, record)
            job_history[name] = job_history[name][:JOB_HISTORY_MAX]
            _save_history()
            print(f"  [{name}] ❌ 실행 오류: {e}")
        finally:
            running_jobs.discard(name)
    threading.Thread(target=_stream, daemon=True).start()

def make_cmd(sp):
    """ .exe는 직접 실행, .py는 python으로 실행 """
    if sp.endswith(".exe"):
        return [sp]
    return [sys.executable, sp]

def setup_automation_schedule():
    """
    SERVER_MODE=test  → JOB_AL 순서대로 '지금+2분' 간격으로 순차 실행
    SERVER_MODE=prod  → JOB_AL의 절대 시간 그대로 사용

    JOB_AL 형식: 이름:요일:시:분:스크립트경로
    """
    mode = os.getenv("SERVER_MODE", "prod").strip().lower()

    raw_jobs = os.getenv("JOB_AL", "")
    if not raw_jobs:
        print("   ⚠️ JOB_AL 환경변수가 설정되지 않았습니다.")
        return

    job_lines = []
    for item in raw_jobs.splitlines():
        if not item.strip():
            continue
        parts = item.strip().split(":", 4)
        if len(parts) != 5:
            print(f"   ⚠️ JOB_AL 형식 오류 (이름:요일:시:분:스크립트경로): {item}")
            continue
        job_lines.append(parts)  # [name, days, h, m, script_path]

    if not job_lines:
        print("   ⚠️ JOB_AL에 유효한 job이 없습니다.")
        return

    # 테스트 모드: 지금 시각 + 2분 간격으로 시간 덮어쓰기
    if mode == "test":
        now = datetime.now()
        base_total = now.hour * 60 + now.minute + 2
        time_override = {}
        for i, (name, _, h, m, _) in enumerate(job_lines):
            t = (base_total + i * 2) % (24 * 60)
            time_override[name] = (t // 60, t % 60)
        print(f"   🧪 TEST 모드: {now.strftime('%H:%M')} 기준 2분 간격 자동 배분")
    else:
        time_override = {}

    for name, days, h_str, m_str, script_path in job_lines:
        try:
            if name in time_override:
                h, m = time_override[name]
            else:
                h, m = int(h_str), int(m_str)

            # days 형식: dom-N → 매달 N일, 그 외 → 요일(mon-fri 등)
            if days.startswith("dom-"):
                day_of_month = int(days.split("-")[1])
                trigger = CronTrigger(
                    day=day_of_month,
                    hour=h,
                    minute=m,
                    timezone=os.getenv("APP_TIMEZONE", "America/Chicago")
                )
            else:
                trigger = CronTrigger(
                    day_of_week=days,
                    hour=h,
                    minute=m,
                    timezone=os.getenv("APP_TIMEZONE", "America/Chicago")
                )

            scheduler.add_job(
                lambda sp=script_path, n=name: _run_job_with_log(sp, n),
                trigger=trigger,
                id=f"JOB_{name}",
                replace_existing=True,
                misfire_grace_time=1800
            )
            job_scripts[name] = script_path
            print(f"   [Registered] {name}: {days} {h:02d}:{m:02d} -> {script_path}")
        except Exception as e:
            print(f"   ⚠️ job 등록 오류 ({name}): {e}")

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

def _is_job_running(name: str) -> bool:
    """running_jobs 추적 OR psutil로 실제 프로세스 실행 여부 확인"""
    if name in running_jobs:
        return True
    script = job_scripts.get(name, "")
    if not script:
        return False
    basename = os.path.basename(script).lower()
    try:
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                pname = (proc.info['name'] or "").lower()
                cmdline = " ".join(proc.info['cmdline'] or []).lower()
                if basename in pname or basename in cmdline:
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception:
        pass
    return False

@app.get("/api/status", tags=["Monitor"])
async def api_status():
    """등록된 잡 스케줄 + 마지막 실행 결과 반환"""
    jobs_info = []
    for job in scheduler.get_jobs():
        next_run = job.next_run_time
        if next_run:
            now = datetime.now(next_run.tzinfo)
            remaining_sec = int((next_run - now).total_seconds())
            next_run_str = next_run.strftime("%m/%d %H:%M")
        else:
            remaining_sec = None
            next_run_str = None
        name = job.id.replace("JOB_", "")
        hist_list = job_history.get(name, [])
        last = hist_list[0] if hist_list else None
        jobs_info.append({
            "name": name,
            "next_run": next_run_str,
            "remaining_sec": remaining_sec,
            "running": _is_job_running(name),
            "last_time": last["time"] if last else None,
            "last_status": last["status"] if last else None,
        })
    # remaining_sec 기준 오름차순 정렬 (None은 뒤로)
    jobs_info.sort(key=lambda x: (x["remaining_sec"] is None, x["remaining_sec"] or 0))
    return {
        "jobs": jobs_info,
        "server_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "server_name": os.getenv("SERVER_NAME", "RPA Server"),
        "server_ip": _get_server_ip(),
        "server_mode": os.getenv("SERVER_MODE", "prod").strip().lower(),
    }

@app.get("/api/history/{name}", tags=["Monitor"])
async def api_job_history(name: str):
    """job별 실행 이력 반환"""
    records = job_history.get(name.upper(), [])
    return {"name": name.upper(), "history": records}

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

        cmd = [sys.executable, "rpa_tasks/dailyprintout/daily_printout_automail.py"]
        if extra_email:
            cmd.append(extra_email)

        subprocess.Popen(cmd)
        return {"status": "success", "message": "Process started in background."}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/run/daily_report_automail/stream", tags=["Action"])
async def run_daily_report_automail_stream(extra_email: str = Query(None)):
    async def log_generator():
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['cmdline'] and 'daily_printout_automail.py' in proc.info['cmdline']:
                yield f"data: ❌ Error: Another process is already running.\n\n"
                return
        
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        cmd = [sys.executable, "-u", "rpa_tasks/dailyprintout/daily_printout_automail.py"]
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
async def run_db_sink_stream():
    async def log_generator():
        for proc in psutil.process_iter(['name', 'cmdline']):
            if proc.info['cmdline'] and 'db_sink_prod_to_dev.py' in proc.info['cmdline']:
                yield f"data: ❌ Error: Another process is already running.\n\n"
                return

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        cmd = [sys.executable, "-u", "rpa_tasks/dailyprintout/db_sink_prod_to_dev.py"]

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
    
    uvicorn.run(app, host=host, port=port, access_log=False)
