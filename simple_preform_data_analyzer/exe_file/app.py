#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PC 전용 실행앱 (Tkinter) — app.py
- ab.xlsx / alls.xlsx 선택 → [실행] → new_main.py(run-all) 실행
- 표준출력 실시간 표시(-u, PYTHONUNBUFFERED, UTF-8)
- EXE(onefile)일 때는 --worker 모드로 자기 자신을 실행하여 내부에서
  new_main을 모듈로 import → main(argv) 호출(자기 재실행 문제 해결)
"""

import os, sys, threading, queue, subprocess, importlib.util
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Integrated Fiber Analyzer - Runner"
DEFAULT_WIDTH, DEFAULT_HEIGHT = 1080, 700

# ─────────────────────────────────────────────────────────────
# new_main.py 위치 탐색 (개발/EXE 모두 지원) — 못 찾으면 None
# ─────────────────────────────────────────────────────────────
def get_script_path(name: str) -> Path | None:
    cands = []
    if hasattr(sys, "_MEIPASS"):  # PyInstaller onefile 임시폴더
        cands.append(Path(sys._MEIPASS) / name)
    cands.append(Path(__file__).resolve().parent / name)
    for p in cands:
        if p.exists():
            return p
    return None

NEW_MAIN = get_script_path("new_main.py")

# ─────────────────────────────────────────────────────────────
# EXE 전용 워커: new_main을 import하여 main(argv) 호출
# ─────────────────────────────────────────────────────────────
def _run_new_main_module(mod, argv: list[str]) -> int:
    # new_main.py에 main(argv)가 있으면 호출
    if hasattr(mod, "main") and callable(mod.main):
        try:
            return int(mod.main(argv))  # type: ignore[arg-type]
        except SystemExit as e:
            return int(getattr(e, "code", 0) or 0)
        except Exception as e:
            print(f"[worker] runtime error: {e}", flush=True)
            return 5
    return 0  # main 없음 → 이미 실행된 형태일 수 있음

def run_worker(argv: list[str]) -> int:
    """
    argv: new_main.py 에게 전달할 인자 리스트
          예: ["--ab", "...", "--alls", "...", "run-all"]
    """
    # argparse가 기대하는 argv 구성
    sys.argv = ["new_main.py"] + argv

    # 1) EXE(onefile)일 때: 모듈 import 우선 (PyInstaller가 모듈로 포함)
    if getattr(sys, "frozen", False):
        try:
            import new_main as mod  # ★ --hidden-import new_main 로 포함 필요
        except Exception as e:
            print(f"[worker] module import error: {e}", flush=True)
        else:
            return _run_new_main_module(mod, argv)

    # 2) 개발환경 또는 모듈 import 실패 시: 파일 경로 로드 폴백
    if NEW_MAIN is None:
        print("[worker] new_main.py not found", flush=True)
        return 2

    spec = importlib.util.spec_from_file_location("new_main", str(NEW_MAIN))
    if spec is None or spec.loader is None:
        print("[worker] spec/loader is None", flush=True)
        return 3
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    except SystemExit as e:
        return int(getattr(e, "code", 0) or 0)
    except Exception as e:
        print(f"[worker] import error (file path): {e}", flush=True)
        return 4

    return _run_new_main_module(mod, argv)

# ─────────────────────────────────────────────────────────────
# Tk 앱
# ─────────────────────────────────────────────────────────────
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

    # UI
    def _build_ui(self):
        outer = ttk.Frame(self, padding=16); outer.pack(fill=tk.BOTH, expand=True)

        row_ab = ttk.Frame(outer); row_ab.pack(fill=tk.X, pady=(0,10))
        ttk.Label(row_ab, text="ab.xlsx 파일").pack(side=tk.LEFT)
        self.var_ab = tk.StringVar()
        ttk.Entry(row_ab, textvariable=self.var_ab).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(row_ab, text="찾기...", command=self._pick_ab).pack(side=tk.LEFT)

        row_alls = ttk.Frame(outer); row_alls.pack(fill=tk.X, pady=(0,10))
        ttk.Label(row_alls, text="alls.xlsx 파일").pack(side=tk.LEFT)
        self.var_alls = tk.StringVar()
        ttk.Entry(row_alls, textvariable=self.var_alls).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(row_alls, text="찾기...", command=self._pick_alls).pack(side=tk.LEFT)

        log_row = ttk.Frame(outer); log_row.pack(fill=tk.BOTH, expand=True)
        self.txt = tk.Text(log_row, height=20, wrap=tk.NONE)
        yscroll = ttk.Scrollbar(log_row, orient=tk.VERTICAL, command=self.txt.yview)
        self.txt.configure(yscrollcommand=yscroll.set)
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); yscroll.pack(side=tk.LEFT, fill=tk.Y)

        side = ttk.Frame(log_row); side.pack(side=tk.LEFT, fill=tk.Y, padx=(8,0))
        self.btn_run  = ttk.Button(side, text="▶ 실행", command=self._on_run);  self.btn_run.pack(pady=(0,6), fill=tk.X)
        self.btn_stop = ttk.Button(side, text="■ 중지", command=self._on_stop, state=tk.DISABLED); self.btn_stop.pack(fill=tk.X)

        self.status = tk.StringVar(value="대기 중")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

    # 파일 선택
    def _pick_ab(self):
        p = filedialog.askopenfilename(title="ab.xlsx 선택", filetypes=[["Excel","*.xlsx"],["All","*.*"]])
        if p: self.var_ab.set(p)

    def _pick_alls(self):
        p = filedialog.askopenfilename(title="alls.xlsx 선택", filetypes=[["Excel","*.xlsx"],["All","*.*"]])
        if p: self.var_alls.set(p)

    # 로그
    def _append_log(self, s: str):
        self.txt.insert(tk.END, s); self.txt.see(tk.END)

    def _enable_controls(self, running: bool):
        self.btn_run.config(state=(tk.DISABLED if running else tk.NORMAL))
        self.btn_stop.config(state=(tk.NORMAL if running else tk.DISABLED))

    # 실행
    def _on_run(self):
        ab = self.var_ab.get().strip()
        alls = self.var_alls.get().strip()

        if not ab or not Path(ab).exists():
            messagebox.showwarning("입력 확인", "ab.xlsx 경로를 확인하세요."); return
        if not alls or not Path(alls).exists():
            messagebox.showwarning("입력 확인", "alls.xlsx 경로를 확인하세요."); return

        # new_main.py 확인
        if NEW_MAIN is None and not getattr(sys, "frozen", False):
            messagebox.showerror("실행 불가", "new_main.py를 찾을 수 없습니다.\napp.py와 같은 폴더에 두세요.")
            return

        # argparse 전역 옵션은 서브커맨드(run-all) 앞
        child_argv = ["--ab", ab, "--alls", alls, "run-all"]

        self.txt.delete("1.0", tk.END)
        self._append_log(f"스크립트: {NEW_MAIN}\n")
        self._append_log(f"인자: {' '.join(child_argv)}\n\n")

        try:
            env = os.environ.copy()
            env.update({
                "PYTHONUNBUFFERED": "1",
                "PYTHONIOENCODING": "utf-8",
                "PYTHONUTF8": "1",
            })
            run_cwd = (NEW_MAIN.parent if NEW_MAIN else Path.cwd())  # 결과/로그 위치 일관

            if getattr(sys, "frozen", False):
                # EXE: 같은 exe를 --worker 모드로 실행 → 내부에서 new_main.main(argv) 수행
                cmd = [sys.executable, "--worker", *child_argv]
            else:
                # Python: 해석기로 직접 실행
                cmd = [sys.executable, "-u", str(NEW_MAIN), *child_argv]

            self._append_log("실행 커맨드: " + " ".join(cmd) + "\n\n")

            self.proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True, universal_newlines=True,
                encoding="utf-8", errors="replace",  # cp949 이슈 방지
                bufsize=1,
                env=env, cwd=str(run_cwd),
            )
        except Exception as e:
            messagebox.showerror("실행 오류", str(e)); return

        self._enable_controls(True)
        self.status.set("실행 중…")
        threading.Thread(target=self._reader_thread, daemon=True).start()
        threading.Thread(target=self._waiter_thread, daemon=True).start()

    # 중지
    def _on_stop(self):
        if self.proc and self.proc.poll() is None:
            try: self.proc.terminate()
            except Exception: pass
        self._enable_controls(False)
        self.status.set("중지 요청")

    # 리더 스레드
    def _reader_thread(self):
        try:
            assert self.proc is not None
            for line in self.proc.stdout:  # type: ignore[arg-type]
                self.queue.put(line)
        except Exception as e:
            self.queue.put(f"[LOG 읽기 오류] {e}\n")

    # 종료 대기 스레드
    def _waiter_thread(self):
        if not self.proc: return
        self.proc.wait()
        self.queue.put(f"\n=== 프로세스 종료 (rc={self.proc.returncode}) ===\n")
        self.queue.put("__DONE__")

    # 메인 루프 큐 폴링
    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if item == "__DONE__":
                    self._enable_controls(False); self.status.set("완료")
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.after(50, self._poll_queue)

# ─────────────────────────────────────────────────────────────
# 엔트리포인트: --worker 모드 처리(동결 전용)
# ─────────────────────────────────────────────────────────────
def main():
    # EXE로 실행되었고 --worker가 붙었으면 워커로 동작
    if "--worker" in sys.argv:
        i = sys.argv.index("--worker")
        child_argv = sys.argv[i+1:]  # --worker 이후의 인자만 전달
        rc = run_worker(child_argv)
        raise SystemExit(rc)

    # 일반(UI) 모드
    app = RunnerApp()
    app.mainloop()

if __name__ == "__main__":
    main()
