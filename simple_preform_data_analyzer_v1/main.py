# main.py
from __future__ import annotations
import sys, os
import subprocess
from pathlib import Path
from datetime import datetime
import time

print("안녕하십니까? 통신연구소 소속 김희서 연구원입니다.")
print("광섬유 특성 분석을 효율적으로 진행하기 위해서 초간단 프리폼 데이터값정리 툴을 만들었습니다.")
print("문의사항 있을 경우 hkim17@lscns.com으로 연락주시길 바랍니다.")
print(" ")
print(" ")

SCRIPTS = [
    {"filename": "resin.py",          "desc": "레진 집계 및 접두어별 폴더 구성 (ab.xlsx)"},
    {"filename": "zero.py",           "desc": "0 → 빈칸 정리 (alls.xlsx → alls_cleaned.xlsx)"},
    {"filename": "filter.py",         "desc": "조건 필터 후 4번째 열로 그룹 저장 + 평균행 추가 (grouped_by_col4 생성)"},
    {"filename": "final_average.py",  "desc": "접두어별 평균행 통합 파일 생성 (각 접두어 폴더 내 <접두어>.xlsx)"},
    {"filename": "type.py",           "desc": "타입/제조사 보유 요약 출력"},
    {"filename": "data_analyzer.py",  "desc": "데이터 통합/분석 실행 (total_final_result.xlsx 등 최종 산출)"},
]

STOP_ON_ERROR = True
WORKDIR = Path(__file__).resolve().parent

# ── (A) Windows 콘솔/파이썬 인코딩을 UTF-8로 맞추기 ─────────────
def _setup_utf8_console_and_env():
    # 1) 콘솔 코드페이지를 65001로
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleCP(65001)
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
    # 2) 파이썬 표준출력/에러 스트림 UTF-8로 재설정(3.7+)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    # 3) 자식 파이썬에도 UTF-8 강제
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env

# ── (B) 콘솔 + 파일 동시 로그 ───────────────────────────────────
LOG_DIR = WORKDIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_PATH = LOG_DIR / f"run_{datetime.now():%Y%m%d_%H%M%S}.txt"

class _Tee:
    def __init__(self, *streams):
        self.streams = streams
    def write(self, data):
        for s in self.streams:
            try:
                s.write(data)
                s.flush()
            except Exception:
                pass
    def flush(self):
        for s in self.streams:
            try: s.flush()
            except Exception: pass
    def isatty(self):
        return any(getattr(s, "isatty", lambda: False)() for s in self.streams)

# UTF-8 환경 먼저 구성
CHILD_ENV = _setup_utf8_console_and_env()
_log_file = open(LOG_PATH, "w", encoding="utf-8", newline="")
sys.stdout = _Tee(sys.__stdout__, _log_file)
sys.stderr = _Tee(sys.__stderr__, _log_file)

def run_step(idx: int, total: int, script_path: Path, desc: str) -> int:
    tag = f"[{idx}/{total}]"
    start_ts = time.perf_counter()
    print(f"{tag} {script_path.name} 실행 시작  | {desc}  | {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("-" * 100)

    rc = 1
    try:
        # 자식 출력도 UTF-8로 읽어서 콘솔+파일로 tee
        proc = subprocess.Popen(
            [sys.executable, str(script_path)],
            cwd=WORKDIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",   # 자식도 UTF-8로 강제
            errors="replace",
            bufsize=1,
            env=CHILD_ENV,      # ★ 자식 환경에 UTF-8 강제 변수 전달
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line, end="")
        rc = proc.wait()
    except FileNotFoundError:
        print(f"[오류] 파일을 찾을 수 없습니다: {script_path}")
        rc = 127
    except Exception as e:
        print(f"[오류] 실행 중 예외 발생: {e}")
        rc = 1
    finally:
        elapsed = time.perf_counter() - start_ts
        print("-" * 100)
        status = "성공" if rc == 0 else f"실패(rc={rc})"
        print(f"{tag} {script_path.name} 종료  | {status}  | 소요 {elapsed:.2f}s\n")
    return rc

def main() -> int:
    print(f"[log] 화면+파일 동시 기록: {LOG_PATH}")
    print(f"=== 파이썬 실행 파일: {sys.executable}")
    print(f"=== 작업 디렉터리: {WORKDIR}")
    print("=== 일괄 실행을 시작합니다.\n")

    # 파일 존재 확인
    missing = [ (WORKDIR / s["filename"]).name for s in SCRIPTS if not (WORKDIR / s["filename"]).exists() ]
    if missing:
        print("[오류] 다음 스크립트 파일이 없습니다:")
        for name in missing:
            print(f"  - {name}")
        _log_file.close()
        return 1

    total = len(SCRIPTS)
    failed, executed = [], 0
    for i, s in enumerate(SCRIPTS, start=1):
        executed += 1
        rc = run_step(i, total, WORKDIR / s["filename"], s["desc"])
        if rc != 0:
            failed.append((s["filename"], rc))
            if STOP_ON_ERROR:
                print(f"[중단] {s['filename']} 실패로 이후 작업을 중단합니다.")
                break

    print("=== 실행 요약 ===")
    print(f"- 총 스텝: {total}")
    print(f"- 성공: {executed - len(failed)}")
    print(f"- 실패: {len(failed)}")
    if failed:
        for name, rc in failed:
            print(f"  · {name}: rc={rc}")
        _log_file.close()
        return 1

    print("모든 스텝이 성공적으로 완료되었습니다.")
    _log_file.close()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
