#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PC 전용 실행앱 (Tkinter) — app.py (FULL)
- ab.xlsx / alls.xlsx 파일을 선택 → [실행] 버튼 → new_main.py(run-all) 실행
- 콘솔 로그를 실시간으로 텍스트 영역에 표시(버퍼링 방지/UTF-8 고정)
- EXE로 빌드해도 동일 동작(PyInstaller)

필수: 동일 폴더에 new_main.py 배치
권장: Python 3.9+, pandas, openpyxl, numpy 설치
"""

import os
import sys
import threading
import queue
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Integrated Fiber Analyzer - Runner"
DEFAULT_WIDTH = 1080
DEFAULT_HEIGHT = 700

# new_main.py 위치(동일 폴더 권장). EXE로 묶을 때는 sys._MEIPASS(임시폴더)도 확인
def get_script_path(name: str) -> Path:
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / name)  # PyInstaller onefile unpack dir
    candidates.append(Path(__file__).resolve().parent / name)
    for c in candidates:
        if c.exists():
            return c
    # 마지막 수단: 현재 작업 디렉터리
    return Path(name)

NEW_MAIN = get_script_path("new_main.py")

class RunnerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{DEFAULT_WIDTH}x{DEFAULT_HEIGHT}")
        self.minsize(900, 560)

        self._build_ui()
        self.proc: subprocess.Popen | None = None
        self.queue: "queue.Queue[str]" = queue.Queue()
        self.after(50, self._poll_queue)

    # ────────────────────────────────────────────── UI
    def _build_ui(self):
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill=tk.BOTH, expand=True)

        # ab.xlsx
        ab_row = ttk.Frame(outer)
        ab_row.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(ab_row, text="ab.xlsx 파일").pack(side=tk.LEFT)
        self.var_ab = tk.StringVar()
        self.ent_ab = ttk.Entry(ab_row, textvariable=self.var_ab)
        self.ent_ab.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(ab_row, text="찾기...", command=self._pick_ab).pack(side=tk.LEFT)

        # alls.xlsx
        alls_row = ttk.Frame(outer)
        alls_row.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(alls_row, text="alls.xlsx 파일").pack(side=tk.LEFT)
        self.var_alls = tk.StringVar()
        self.ent_alls = ttk.Entry(alls_row, textvariable=self.var_alls)
        self.ent_alls.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(alls_row, text="찾기...", command=self._pick_alls).pack(side=tk.LEFT)

        # 로그 + 실행 버튼
        log_row = ttk.Frame(outer)
        log_row.pack(fill=tk.BOTH, expand=True)

        self.txt = tk.Text(log_row, height=20, wrap=tk.NONE)
        yscroll = ttk.Scrollbar(log_row, orient=tk.VERTICAL, command=self.txt.yview)
        self.txt.configure(yscrollcommand=yscroll.set)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.LEFT, fill=tk.Y)

        side = ttk.Frame(log_row)
        side.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))
        self.btn_run = ttk.Button(side, text="▶ 실행", command=self._on_run)
        self.btn_run.pack(pady=(0, 6), fill=tk.X)
        self.btn_stop = ttk.Button(side, text="■ 중지", command=self._on_stop, state=tk.DISABLED)
        self.btn_stop.pack(fill=tk.X)

        # 상태바
        self.status = tk.StringVar(value="대기 중")
        sbar = ttk.Label(self, textvariable=self.status, anchor=tk.W)
        sbar.pack(side=tk.BOTTOM, fill=tk.X)

    # ────────────────────────────────────────────── Helpers
    def _pick_ab(self):
        p = filedialog.askopenfilename(title="ab.xlsx 선택", filetypes=[["Excel", "*.xlsx"], ["All", "*.*"]])
        if p:
            self.var_ab.set(p)

    def _pick_alls(self):
        p = filedialog.askopenfilename(title="alls.xlsx 선택", filetypes=[["Excel", "*.xlsx"], ["All", "*.*"]])
        if p:
            self.var_alls.set(p)

    def _append_log(self, text: str):
        self.txt.insert(tk.END, text)
        self.txt.see(tk.END)

    def _enable_controls(self, running: bool):
        self.btn_run.config(state=(tk.DISABLED if running else tk.NORMAL))
        self.btn_stop.config(state=(tk.NORMAL if running else tk.DISABLED))

    # ────────────────────────────────────────────── Run/Stop
    def _on_run(self):
        ab = self.var_ab.get().strip()
        alls = self.var_alls.get().strip()
        if not ab or not Path(ab).exists():
            messagebox.showwarning("입력 확인", "ab.xlsx 경로를 확인하세요.")
            return
        if not alls or not Path(alls).exists():
            messagebox.showwarning("입력 확인", "alls.xlsx 경로를 확인하세요.")
            return
        if not NEW_MAIN.exists():
            messagebox.showerror("실행 불가", f"new_main.py를 찾을 수 없습니다.\n경로: {NEW_MAIN}")
            return

        # 실행 커맨드 구성 (버퍼링 방지 -u)
        py = sys.executable  # exe로 묶여도 내장 파이썬 사용
        cmd = [py, "-u", str(NEW_MAIN), "--ab", ab, "--alls", alls, "run-all"]

        # 로그 초기화
        self.txt.delete("1.0", tk.END)
        self._append_log(f"실행 파일: {py}\n")
        self._append_log(f"스크립트:  {NEW_MAIN}\n")
        self._append_log(f"입력 ab:   {ab}\n입력 alls: {alls}\n\n")

        try:
            # 실시간 출력 및 인코딩, UTF-8 고정
            env = os.environ.copy()
            env.update({
                "PYTHONUNBUFFERED": "1",
                "PYTHONIOENCODING": "utf-8",
                "PYTHONUTF8": "1",
            })
            # Anaconda와 동일 체감: new_main.py 위치를 작업 디렉터리로 고정
            run_cwd = NEW_MAIN.parent

            self.proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,               # 텍스트 모드에서 라인 버퍼링
                universal_newlines=True, # 텍스트 모드
                encoding="utf-8",         # ★ 부모가 파이프를 UTF-8로 해석
                errors="replace",         # ★ 혹시 모르는 깨진 문자도 예외 없이 치환
                env=env,
                cwd=str(run_cwd),
            )
        except Exception as e:
            messagebox.showerror("실행 오류", str(e))
            return

        self._enable_controls(True)
        self.status.set("실행 중…")

        t = threading.Thread(target=self._reader_thread, daemon=True)
        t.start()

        w = threading.Thread(target=self._waiter_thread, daemon=True)
        w.start()

    def _on_stop(self):
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
            except Exception:
                pass
        self._enable_controls(False)
        self.status.set("중지 요청")

    # ────────────────────────────────────────────── Threads
    def _reader_thread(self):
        try:
            assert self.proc is not None
            for line in self.proc.stdout:
                self.queue.put(line)
        except Exception as e:
            self.queue.put(f"[LOG 읽기 오류] {e}\n")

    def _waiter_thread(self):
        if self.proc is None:
            return
        self.proc.wait()
        code = self.proc.returncode
        self.queue.put(f"\n=== 프로세스 종료 (rc={code}) ===\n")
        self.queue.put("__DONE__")

    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if item == "__DONE__":
                    self._enable_controls(False)
                    self.status.set("완료")
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.after(50, self._poll_queue)


def main():
    app = RunnerApp()
    app.mainloop()


if __name__ == "__main__":
    main()

"""
# ───────────────────────────────────────────────────────────── PyInstaller 빌드 메모 ─────────────────────────────────────────────────────────────
# 1) 가상환경(선택)에서 필요한 라이브러리 설치: pandas openpyxl numpy (new_main.py에 맞게)
# 2) app.py와 new_main.py를 같은 폴더에 둡니다.
# 3) 다음 커맨드로 단일 실행파일 생성 (콘솔창 포함, 로그 보기 용이):
#    pyinstaller --onefile --add-data "new_main.py;." --name FiberAnalyzerRunner app.py
#    (세미콜론은 Windows 구분자입니다. macOS/Linux는 콜론 ":" 사용)
# 4) dist/FiberAnalyzerRunner.exe 가 생성됩니다. ab.xlsx, alls.xlsx만 있으면 실행 가능.
#    첫 실행 시 SmartScreen 경고가 나올 수 있으니 사내 배포 시 서명 또는 내부 가이드 포함 권장.
"""