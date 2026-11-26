#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 광섬유 데이터 분석 파이프라인 (One-File Edition)
작성자: ###
목적: 산재된 스크립트(resin/zero/group/final/type/analyzer)를 하나로 통합하여
     단일 파일에서 일괄 실행/부분 실행이 가능하도록 구성

Python 3.9+ 권장. 의존성: pandas, openpyxl

사용 예시:
  1) 전체 실행:          python integrated_fiber_analyzer.py run-all
  2) 일부 단계만:        python integrated_fiber_analyzer.py run zero group collect-avg reports collect-total
  3) 옵션 확인/도움말:   python integrated_fiber_analyzer.py -h

주요 입력/출력 경로(기본값):
  - ab.xlsx                 : 레진/접두어 분석용 원본
  - alls.xlsx               : 전체 데이터 원본
  - alls_cleaned.xlsx       : 0 → 빈칸 처리된 파일(중간 산출)
  - grouped_by_prefix/      : draw_no 앞 3글자 기준으로 폴더 구성
  - grouped_by_col4/        : 3/4열 기반 그루핑 결과 루트 및 후속 산출물 저장 위치
  - grouped_by_col4/<코드>/<코드>.xlsx                        : 접두어별 평균행 통합 파일
  - grouped_by_col4/<코드>/<코드>_final_result_report.xlsx    : 각 폴더 리포트
  - grouped_by_col4/total_final_result.xlsx                   : 전체 통합 리포트

주의:
  - 윈도우 콘솔 UTF-8, 화면+파일 동시 로깅 지원
  - 단계별 실패 시 STOP_ON_ERROR 설정에 따라 중단/계속
  - 열 인덱스는 0-based

추가(요청 반영):
  - group 단계에서 엑셀 저장 전, "3번째 열(0-based index 2)" 값이 같은 행은
    첫 번째만 남기고 제거한 뒤 평균행을 계산/추가합니다.

신규(요청 반영):
  - collect-total 이후 "post-analyze" 단계 추가
    · 23번째 열(0-based 22) delta(2m)-22m의 최솟/최댓값을 찾아 콘솔에 알리고, 해당 셀만 빨간색으로 표시
    · 25/26번째 열(0-based 24/25) Clad Dia. 값이 124.3 미만 또는 125.7 초과면 빨간색으로 표시,
      콘솔에 "이상값 발견" 및 해당 행의 2번째 열 값 출력
    · 결과 파일: grouped_by_col4/total_final_result_annotated.xlsx
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

# ──────────────────────────────────────────────────────────────────────
# 의존성 점검
# ──────────────────────────────────────────────────────────────────────
try:  # 지연 임포트 대비, 즉시 실패 시 친절 메시지
    import pandas as pd
    import numpy as np
except Exception as e:  # pragma: no cover
    print("[오류] pandas 또는 numpy 임포트 실패.")
    print("       pip로 설치해 주세요:  pip install pandas openpyxl numpy")
    raise

# openpyxl은 pandas가 내부에서 사용

# ──────────────────────────────────────────────────────────────────────
# 설정값
# ──────────────────────────────────────────────────────────────────────
@dataclass
class Config:
    # 주요 파일/폴더 경로
    excel_ab: Path = Path("ab.xlsx")
    excel_alls: Path = Path("alls.xlsx")
    excel_alls_cleaned: Path = Path("alls_cleaned.xlsx")
    out_grouped_by_prefix: Path = Path("grouped_by_prefix")
    out_grouped_by_col4: Path = Path("grouped_by_col4")

    # 열 인덱스(0-based)
    resin_col_idx: int = 4     # ab.xlsx의 5번째 열(E)
    drawno_col_idx: int = 0    # ab.xlsx의 1번째 열(A)
    col3_idx: int = 2          # alls_cleaned.xlsx 의 C열
    col4_idx: int = 3          # alls_cleaned.xlsx 의 D열

    # 규칙/동작 토글
    use_w_pattern_first: bool = False  # 접두 추출 시 W-패턴 우선 여부
    filter_second_last_zero: bool = True  # C열의 뒤에서 2번째가 '0'인 행만 사용
    stop_on_error: bool = True

    # 로깅
    log_dir: Path = Path("logs")

    # type.py 매핑
    type_map: Dict[str, Dict[str, List[str]]] = field(default_factory=lambda: {
        "LWPF(90)":  {"SEC": ["W00", "W0J"], "Sumitomo": ["20M"]},
        "LWPF(150)": {"SEC": ["L0E"],           "Sumitomo": ["L0M"]},
        "LWPF(180)": {"SEC": ["S0E"],           "Sumitomo": ["S0M"]},
        "A1(90)":    {"SEC": [],                "Sumitomo": ["Z0M"]},
        "A1(150)":   {"SEC": [],                "Sumitomo": ["Z0L"]},
        "A2(90)":    {"SEC": ["AJW", "AJF", "AJB"], "Sumitomo": []},
        "A2(150)":   {"SEC": ["AL"],            "Sumitomo": []},
    })


CFG = Config()

# ──────────────────────────────────────────────────────────────────────
# 콘솔/로깅 유틸
# ──────────────────────────────────────────────────────────────────────
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
            try:
                s.flush()
            except Exception:
                pass

    def isatty(self):
        return any(getattr(s, "isatty", lambda: False)() for s in self.streams)


def setup_utf8_console_and_env() -> Dict[str, str]:
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleCP(65001)
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


class Logger:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.cfg.log_dir.mkdir(exist_ok=True)
        self.log_path = self.cfg.log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.txt"
        self._log_f = open(self.log_path, "w", encoding="utf-8", newline="")
        sys.stdout = _Tee(sys.__stdout__, self._log_f)
        sys.stderr = _Tee(sys.__stderr__, self._log_f)

    def close(self):
        try:
            self._log_f.close()
        except Exception:
            pass


# ──────────────────────────────────────────────────────────────────────
# 공통 유틸 함수
# ──────────────────────────────────────────────────────────────────────
SAFE_NAME = re.compile(r"^[A-Za-z0-9_\-.]+$")
W_PREFIX_REGEX = re.compile(r"^([A-Z0-9]{3}\d{5}[A-Z]\d{2}W\d{2}[^0-9])")
GENERIC_RIGHTMOST_CHAR_BEFORE_DIGIT = re.compile(r"^(.+[A-Z])(?=\d)")
FILENAME_TO_PREFORM = re.compile(r"^([A-Z0-9]{3}\d{5}).*?([A-Z])$")
ZERO_LIKE = re.compile(r'^[\+\-]?\s*0+(?:[.,]0+)?\s*$')


def normalize_str(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None


def second_last_is_zero(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip()
    return len(s) >= 2 and s[-2] == "0"


def safe_filename(name: str) -> str:
    s = str(name).strip() or "EMPTY"
    return re.sub(r"[^A-Za-z0-9._-]+", "_", s)


def make_avg_row(df: pd.DataFrame) -> Dict[str, object]:
    avg = {}
    for col in df.columns:
        s_num = pd.to_numeric(df[col], errors="coerce")
        avg[col] = s_num.mean() if s_num.notna().any() else ""
    return avg


def extract_prefix_generic(s: str) -> str:
    if s is None:
        return ""
    t = str(s).strip().upper()
    if not t:
        return ""
    m = GENERIC_RIGHTMOST_CHAR_BEFORE_DIGIT.search(t)
    return m.group(1) if m else t


def extract_prefix_wpattern(s: str) -> str:
    if s is None:
        return ""
    t = str(s).strip().upper()
    m = W_PREFIX_REGEX.match(t)
    return m.group(1) if m else ""


def extract_group_prefix(s: str, use_w_first: bool) -> str:
    if use_w_first:
        p = extract_prefix_wpattern(s)
        return p if p else extract_prefix_generic(s)
    else:
        p = extract_prefix_generic(s)
        return p if p else extract_prefix_wpattern(s)


def is_empty(val) -> bool:
    if val is None:
        return True
    try:
        if pd.isna(val):
            return True
    except Exception:
        pass
    return str(val).strip() == ""


def candidate_files(prefix_dir: Path) -> List[Path]:
    out_file = prefix_dir / f"{prefix_dir.name}.xlsx"
    return sorted(
        p for p in prefix_dir.glob("*.xlsx")
        if p.name.lower() != out_file.name.lower() and not p.name.startswith("~$")
    )


def preform_from_filename(path: Path, fallback: Optional[str] = None) -> Optional[str]:
    stem = path.stem.upper()
    m = FILENAME_TO_PREFORM.match(stem)
    if m:
        base8, last_letter = m.group(1), m.group(2)
        return f"{base8}{last_letter}"
    return fallback


def _normalize_as_text(s: pd.Series) -> pd.Series:
    out = s.astype("string").fillna("")
    out = out.str.replace(r"\.0$", "", regex=True)
    return out.astype("object")


def _is_temp_or_hidden(p: Path) -> bool:
    name = p.name
    return name.startswith("~$") or name.startswith(".") or name.endswith(".tmp")


def pick_input_file(subfolder: Path) -> Optional[Path]:
    c1 = subfolder / f"{subfolder.name}.xlsx"
    c2 = subfolder / "final.xlsx"
    if c1.exists():
        return c1
    if c2.exists():
        return c2
    return None

# ──────────────────────────────────────────────────────────────────────
# 단계 구현
# ──────────────────────────────────────────────────────────────────────

def step_resin_analyze_and_group_ab(cfg: Config) -> int:
    print("[resin] ab.xlsx 레진/접두어 분석 및 폴더 구성")
    if not cfg.excel_ab.exists():
        print(f"[resin][오류] 엑셀 파일 없음: {cfg.excel_ab.resolve()}")
        return 1

    df = pd.read_excel(cfg.excel_ab, engine="openpyxl")

    # (1) 레진 집계
    if cfg.resin_col_idx >= df.shape[1]:
        print(f"[resin][오류] 5번째 열(인덱스 {cfg.resin_col_idx}) 없음. 실제 열 수: {df.shape[1]}")
        return 1

    resin_series = (
        df.iloc[:, cfg.resin_col_idx].map(normalize_str).dropna().map(lambda s: s.upper())
    )

    if len(resin_series) == 0:
        print("[resin] 유효한 레진 타입이 없습니다.")
    else:
        resin_counts = resin_series.value_counts().sort_index()
        types_str = ",".join(resin_counts.index)
        print(f"[resin] 레진 타입: {types_str}")
        for t, c in resin_counts.items():
            print(f"[resin] {t}: {c}개")

    # (2) draw_no 앞 3글자 기준 폴더 구성
    if cfg.drawno_col_idx >= df.shape[1]:
        print(f"[resin][오류] 1번째 열(인덱스 {cfg.drawno_col_idx}) 없음. 실제 열 수: {df.shape[1]}")
        return 1

    draw_series = df.iloc[:, cfg.drawno_col_idx].map(normalize_str).dropna()

    prefix_map: Dict[str, set] = {}
    for val in draw_series:
        if len(val) < 3:
            continue
        if not SAFE_NAME.match(val):
            continue
        prefix = val[:3]
        if not SAFE_NAME.match(prefix):
            continue
        prefix_map.setdefault(prefix, set()).add(val)

    cfg.out_grouped_by_prefix.mkdir(parents=True, exist_ok=True)
    for prefix, fullset in prefix_map.items():
        prefix_dir = cfg.out_grouped_by_prefix / prefix
        prefix_dir.mkdir(exist_ok=True)
        for full in sorted(fullset):
            (prefix_dir / full).mkdir(exist_ok=True)

    if prefix_map:
        prefix_list = ",".join(sorted(prefix_map.keys()))
        print(f"[resin] 조회된 접두어: {prefix_list}")
        for prefix in sorted(prefix_map.keys()):
            cnt = len(prefix_map[prefix])
            print(f"[resin] {prefix}: draw_no {cnt}개")
    else:
        print("[resin] 접두어 기반 폴더 생성 대상 없음")

    # (선택) 요약 CSV 저장
    try:
        if len(resin_series) > 0:
            resin_counts.to_frame("count").to_csv(cfg.excel_ab.with_name("resin_type_counts.csv"))
        if prefix_map:
            import csv
            out_csv = cfg.excel_ab.with_name("prefix_drawno_counts.csv")
            with out_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["prefix", "draw_no_count"])
                for p in sorted(prefix_map.keys()):
                    w.writerow([p, len(prefix_map[p])])
    except Exception as e:
        print(f"[resin](경고) 요약 CSV 저장 실패: {e}")

    return 0


def step_zero_to_blank_all(cfg: Config) -> int:
    print("[zero] 0 → 빈칸(None) 변환")
    if not cfg.excel_alls.exists():
        print(f"[zero][오류] 엑셀 파일 없음: {cfg.excel_alls.resolve()}")
        return 1

    df = pd.read_excel(cfg.excel_alls, engine="openpyxl")

    # 숫자형 0 → None (bool 제외)
    num_cols = df.select_dtypes(include=[np.number]).columns
    bool_cols = df.select_dtypes(include=["bool"]).columns
    num_cols = [c for c in num_cols if c not in bool_cols]

    for c in num_cols:
        df[c] = df[c].mask(df[c] == 0, other=None)

    # 비숫자형에서 '0' 변형들 → None
    obj_cols = df.columns.difference(num_cols).tolist()
    for c in obj_cols:
        s = df[c]
        s_str = s.astype(str)
        mask = s_str.str.match(ZERO_LIKE, na=False)
        num_eq_zero = pd.to_numeric(s_str.str.replace(",", ".", regex=False), errors="coerce").eq(0)
        final_mask = (mask | num_eq_zero) & s.notna()
        df.loc[final_mask, c] = None

    df.to_excel(cfg.excel_alls_cleaned, index=False, engine="openpyxl")
    print(f"[zero] 완료 → {cfg.excel_alls_cleaned.resolve()}")
    return 0


def step_group_by_col4_with_prefix_and_avg(cfg: Config) -> int:
    print("[group] 3/4열 기반 그룹 저장 + 평균행 추가 (중복 제거 후)")
    if not cfg.excel_alls_cleaned.exists():
        print(f"[group][오류] 파일 없음: {cfg.excel_alls_cleaned.resolve()}")
        return 1

    df = pd.read_excel(cfg.excel_alls_cleaned, engine="openpyxl")

    need_max = max(cfg.col3_idx, cfg.col4_idx)
    if df.shape[1] <= need_max:
        print(f"[group][오류] 열이 부족합니다. 현재 {df.shape[1]}열, 필요 최소 {need_max+1}")
        return 1

    col3 = df.columns[cfg.col3_idx]
    col4 = df.columns[cfg.col4_idx]

    # (A) C열 필터 (옵션)
    filtered = df.copy()
    if cfg.filter_second_last_zero:
        mask = filtered[col3].map(second_last_is_zero)
        filtered = filtered[mask].copy()

    # (B) D열 공백 제거
    filtered = filtered[filtered[col4].notna() & (filtered[col4].astype(str).str.strip() != "")]
    if filtered.empty:
        print("[group] 필터 후 데이터가 없습니다.")
        return 0

    # (C) 그룹 키 추출
    filtered["_group_key_"] = filtered[col3].apply(lambda s: extract_group_prefix(s, cfg.use_w_pattern_first))
    filtered = filtered[filtered["_group_key_"].astype(str).str.strip() != ""]
    if filtered.empty:
        print("[group] 유효한 그룹 키가 없습니다.")
        return 0

    cfg.out_grouped_by_col4.mkdir(parents=True, exist_ok=True)

    from collections import defaultdict
    prefix_counts: Dict[str, int] = defaultdict(int)

    for key, g in filtered.groupby("_group_key_", dropna=False):
        key_str = str(key).strip()
        if not key_str:
            continue

        prefix3 = key_str[:3] if len(key_str) >= 3 else "UNK"
        dest_dir = cfg.out_grouped_by_col4 / prefix3
        dest_dir.mkdir(parents=True, exist_ok=True)

        # === 저장 대상 테이블 구성 ===
        g_out = g.drop(columns=["_group_key_"]).copy()

        # 🔹 (추가) 평균 계산 전에 "3번째 열(인덱스 2)" 기준 중복 제거
        if g_out.shape[1] >= 3:
            dedup_col = g_out.columns[2]  # 0-based: 2 -> 3번째 열
            before_n = len(g_out)
            g_out["_dedup_key_"] = g_out[dedup_col].map(normalize_str)
            g_out = (
                g_out
                .drop_duplicates(subset=["_dedup_key_"], keep="first")
                .drop(columns=["_dedup_key_"])
                .reset_index(drop=True)
            )
            removed = before_n - len(g_out)
            if removed > 0:
                print(f"[group][중복제거] {key_str}: 3번째 열 '{dedup_col}' 기준 {removed}행 제거")
        else:
            print(f"[group][정보] {key_str}: 열 수가 3 미만이라 중복 제거 스킵")

        # 🔹 평균행 계산 및 부착
        avg_row = make_avg_row(g_out)
        g_out = pd.concat([g_out, pd.DataFrame([avg_row])], ignore_index=True)

        out_path = dest_dir / f"{safe_filename(key_str)}.xlsx"
        try:
            g_out.to_excel(out_path, index=False, engine="openpyxl")
        except Exception as e:
            print(f"[group][오류] 저장 실패: {out_path.name} → {e}")
            continue

        prefix_counts[prefix3] += 1

    for pfx in sorted(prefix_counts.keys()):
        print(f"[group] {pfx}: {prefix_counts[pfx]}개 파일 저장")

    return 0


def step_collect_all_prefix_averages(cfg: Config) -> int:
    print("[collect-avg] 접두어별 평균행 취합")
    base = cfg.out_grouped_by_col4
    if not base.exists():
        print(f"[collect-avg][오류] 폴더가 없습니다: {base.resolve()}")
        return 1

    prefix_dirs = sorted(p for p in base.iterdir() if p.is_dir() and not p.name.startswith("~$"))
    if not prefix_dirs:
        print("[collect-avg] 처리할 접두 폴더가 없습니다.")
        return 0

    for pdir in prefix_dirs:
        out_file = pdir / f"{pdir.name}.xlsx"
        excel_files = candidate_files(pdir)
        if not excel_files:
            print(f"[collect-avg][INFO] {pdir.name}: 수집할 파일 없음")
            continue

        last_rows: List[pd.DataFrame] = []
        for path in excel_files:
            try:
                df = pd.read_excel(path, engine="openpyxl")
            except Exception as e:
                print(f"[collect-avg][경고] 읽기 실패: {path.name} → {e}")
                continue
            if df.empty:
                print(f"[collect-avg][건너뜀] 빈 파일: {path.name}")
                continue

            last_idx = len(df) - 1
            if df.shape[1] > cfg.col4_idx and last_idx >= 1:
                if is_empty(df.iat[last_idx, cfg.col4_idx]):
                    df.iat[last_idx, cfg.col4_idx] = df.iat[last_idx - 1, cfg.col4_idx]

            try:
                current_val = df.iat[last_idx, cfg.col4_idx] if df.shape[1] > cfg.col4_idx else None
                new_preform = preform_from_filename(path, fallback=str(current_val) if current_val is not None else None)
                if new_preform is not None and df.shape[1] > cfg.col4_idx:
                    df.iloc[:, cfg.col4_idx] = df.iloc[:, cfg.col4_idx].astype("object")
                    df.iat[last_idx, cfg.col4_idx] = new_preform
            except Exception as e:
                print(f"[collect-avg][경고] {path.name}: preform 덮어쓰기 오류 → {e}")

            last_rows.append(df.iloc[[last_idx]].copy())

        if not last_rows:
            print(f"[collect-avg][INFO] {pdir.name}: 평균 행 없음")
            continue

        result = pd.concat(last_rows, ignore_index=True, sort=False)
        try:
            result.to_excel(out_file, index=False, engine="openpyxl")
            print(f"[collect-avg][저장] {out_file.resolve()} (총 {len(result)}행)")
        except Exception as e:
            print(f"[collect-avg][오류] {pdir.name} 저장 실패 → {e}")

    return 0


def step_copy_col4_to_col2_in_prefix_books(cfg: Config) -> int:
    print("[copy-42] 접두어 통합파일에서 4번째 열 → 2번째 열(문자열) 복사")
    root = cfg.out_grouped_by_col4
    if not root.exists():
        print(f"[copy-42][오류] 폴더 없음: {root.resolve()}")
        return 1

    SECOND_COL_IDX = 1
    FOURTH_COL_IDX = 3

    prefix_dirs = sorted(p for p in root.iterdir() if p.is_dir() and not _is_temp_or_hidden(p))
    if not prefix_dirs:
        print("[copy-42][정보] 처리할 접두어 폴더가 없습니다.")
        return 0

    for pdir in prefix_dirs:
        target_xlsx = pdir / f"{pdir.name}.xlsx"
        if not target_xlsx.exists():
            print(f"[copy-42][건너뜀] 대상 파일 없음: {target_xlsx}")
            continue

        try:
            df = pd.read_excel(target_xlsx, engine="openpyxl")
        except Exception as e:
            print(f"[copy-42][오류] 읽기 실패: {target_xlsx.name} → {e}")
            continue

        if df.empty:
            print(f"[copy-42][건너뜀] 빈 파일: {target_xlsx.name}")
            continue

        needed = max(SECOND_COL_IDX, FOURTH_COL_IDX) + 1
        if df.shape[1] < needed:
            print(f"[copy-42][경고] {target_xlsx.name}: 열 수 부족({df.shape[1]}열) → 복사 스킵")
            continue

        dst_col = df.columns[SECOND_COL_IDX]
        src_col = df.columns[FOURTH_COL_IDX]

        df[dst_col] = _normalize_as_text(df[dst_col])
        src_as_text = _normalize_as_text(df[src_col])
        df[dst_col] = src_as_text

        try:
            df.to_excel(target_xlsx, index=False, engine="openpyxl")
            print(f"[copy-42][완료] {target_xlsx}")
        except Exception as e:
            print(f"[copy-42][오류] 저장 실패: {target_xlsx.name} → {e}")

    return 0


def step_summarize_types(cfg: Config) -> int:
    print("[types] 타입/제조사 보유 요약")
    base = cfg.out_grouped_by_col4
    if not base.exists():
        print(f"[types][오류] 폴더가 없습니다: {base.resolve()}")
        return 1

    folder_names = set()
    for p in base.iterdir():
        if p.is_dir():
            name = p.name.strip()
            if name and not name.startswith("~$") and not name.startswith("."):
                folder_names.add(name.upper())

    print("[types] 현재 보유 폴더 코드:", ", ".join(sorted(folder_names)) if folder_names else "(없음)")

    defined_upper = set()
    for vendors in cfg.type_map.values():
        for codes in vendors.values():
            defined_upper.update(c.upper() for c in codes)

    matched_present_upper = set()
    any_printed = False

    for type_name, vendors in cfg.type_map.items():
        vendor_parts = []
        for vendor, codes in vendors.items():
            if not codes:
                continue
            present_codes = [c for c in codes if c.upper() in folder_names]
            if present_codes:
                vendor_parts.append(f"{vendor}=" + ", ".join(present_codes))
                matched_present_upper.update(c.upper() for c in present_codes)
        if vendor_parts:
            any_printed = True
            print(f"[types] 타입 {type_name}: " + " / ".join(vendor_parts) + " 보유")

    others = sorted(folder_names - defined_upper)
    if others:
        print("[types] 기타:", ", ".join(others))
    if not any_printed and not others:
        print("[types] (일치하는 타입 코드가 아직 보유되어 있지 않습니다.)")

    return 0


# data_analyzer: 열 맵 정의
COLUMN_INFO: List[Tuple[str, Optional[int], Optional[str], Optional[float]]] = [
    ("spoolno2", 1, None, None),
    ("OTDR length", 9, None, None),
    ("Attenuation 1310 I/E", 5, None, None),
    ("Attenuation 1310 O/E", 6, None, None),
    ("Attenuation 1383 I/E", 73, None, None),
    ("Attenuation 1383 O/E", 74, None, None),
    ("Attenuation 1550 I/E", 7, None, None),
    ("Attenuation 1550 O/E", 8, None, None),
    ("Attenuation 1625 I/E", 75, None, None),
    ("Attenuation 1625 O/E", 76, None, None),
    ("MFD 1310nm I/E", 12, None, None),
    ("MFD 1310nm O/E", 13, None, None),
    ("", None, None, None),
    ("", None, None, None),
    ("", None, None, None),
    ("", None, None, None),
    ("", None, None, None),
    ("", None, None, None),
    ("Cutoff 2m I/E", 14, None, None),
    ("Cutoff 2m O/E", 15, None, None),
    ("Cutoff 22m", 24, None, None),
    ("delta 2m-22m", None, "delta", None),
    ("Mac value", None, "mac", None),
    ("Clad Dia. I/E", 16, None, None),
    ("Clad Dia. O/E", 17, None, None),
    ("Clad Ovality I/E", 18, None, None),
    ("Clad Ovality O/E", 19, None, None),
    ("Core Ovality I/E", 20, None, None),
    ("Core Ovality O/E", 21, None, None),
    ("ECC I/E", 22, None, None),
    ("ECC O/E", 23, None, None),
    ("Zero Dispersion Wave.", 30, None, None),
    ("dispslope at ZDW", 31, None, None),
    ("Dispersion 1285", 32, None, None),
    ("Dispersion 1290", 33, None, None),
    ("Dispersion 1330", 34, None, None),
    ("Dispersion 1550", 35, None, None),
    ("", None, None, None),
    ("PMD", 37, None, None),
    ("R7.5mm 1t 1550", 26, "scale", 0.1),
    ("R7.5mm 1t 1625", 69, "scale", 0.1),
    ("R10mm 1t 1550", 70, "scale", 0.1),
    ("R10mm 1t 1625", 71, "scale", 0.1),
    ("R15mm 10t 1550", 81, "scale", 0.5),
    ("R15mm 10t 1625", 82, "scale", 0.5),
]


def _safe_series(df: pd.DataFrame, col: Optional[int]) -> pd.Series:
    if col is None or col >= df.shape[1]:
        return pd.Series([], dtype="float64")
    return df.iloc[1:, col]


def build_folder_report(subfolder: Path) -> Optional[Path]:
    src = pick_input_file(subfolder)
    if src is None:
        print(f"[report]❌ 입력 없음: {subfolder.name} (<폴더명>.xlsx / final.xlsx)")
        return None

    try:
        df = pd.read_excel(src, header=None, engine="openpyxl")
    except Exception as e:
        print(f"[report]⚠️ 읽기 오류({subfolder.name}): {e}")
        return None

    out = pd.DataFrame()
    for out_idx, (title, src_col, calc, factor) in enumerate(COLUMN_INFO):
        out.loc[0, out_idx] = title
        if calc is None:
            if src_col is None:
                continue
            series = _safe_series(df, src_col)
            for r, v in enumerate(series.tolist()):
                out.loc[r + 1, out_idx] = v
        elif calc == "delta":
            col_20 = pd.to_numeric(out.iloc[1:, 19], errors="coerce")
            col_21 = pd.to_numeric(out.iloc[1:, 20], errors="coerce")
            delta = (col_20 - col_21).round(4)
            for r, v in enumerate(delta.tolist(), start=1):
                out.loc[r, out_idx] = "" if pd.isna(v) else v
        elif calc == "mac":
            mfd_oe = pd.to_numeric(out.iloc[1:, 11], errors="coerce")
            cut_ie = pd.to_numeric(out.iloc[1:, 18], errors="coerce")
            mac = (mfd_oe / cut_ie * 1000).round(2)
            for r, v in enumerate(mac.tolist(), start=1):
                out.loc[r, out_idx] = "" if pd.isna(v) else v
        elif calc == "scale":
            series = _safe_series(df, src_col)
            scaled = (pd.to_numeric(series, errors="coerce") * (factor or 1.0)).round(4)
            for r, v in enumerate(scaled.tolist()):
                out.loc[r + 1, out_idx] = "" if pd.isna(v) else v

    dst = subfolder / f"{subfolder.name}_final_result_report.xlsx"
    try:
        out.to_excel(dst, index=False, header=False, engine="openpyxl")
        print(f"[report]✅ 저장: {dst}")
        return dst
    except Exception as e:
        print(f"[report]⚠️ 저장 오류({subfolder.name}): {e}")
        return None


def step_build_reports(cfg: Config) -> int:
    print("[reports] 하위 폴더별 *_final_result_report.xlsx 생성")
    root = cfg.out_grouped_by_col4
    if not root.exists():
        print(f"[reports][오류] 폴더 없음: {root.resolve()}")
        return 1

    subfolders = [p for p in root.iterdir() if p.is_dir() and not p.name.startswith(("~$", "."))]
    if not subfolders:
        print("[reports] 처리할 하위 폴더가 없습니다.")
        return 0

    for sub in sorted(subfolders, key=lambda x: x.name):
        build_folder_report(sub)

    return 0


def collect_to_root(root: Path, total_filename: str = "total_final_result.xlsx") -> Optional[Path]:
    report_paths = sorted(root.glob("*/*_final_result_report.xlsx"))
    if not report_paths:
        print("[collect-total]⚠️ 통합할 리포트가 없습니다.")
        return None

    merged_list: List[pd.DataFrame] = []
    for p in report_paths:
        try:
            raw = pd.read_excel(p, header=None, engine="openpyxl")
            if raw.empty:
                continue
            headers = raw.iloc[0].tolist()
            data = raw.iloc[1:].reset_index(drop=True)
            data.columns = headers
            data.insert(0, "GROUP", p.parent.name)
            merged_list.append(data)
        except Exception as e:
            print(f"[collect-total]⚠️ 통합 중 읽기 오류: {p} → {e}")

    if not merged_list:
        print("[collect-total]⚠️ 유효 데이터가 없습니다.")
        return None

    total_df = pd.concat(merged_list, ignore_index=True, sort=False)
    total_path = root / total_filename
    try:
        total_df.to_excel(total_path, index=False, engine="openpyxl")
        print(f"[collect-total]📦 저장: {total_path}")
        print("[collect-total] 통합모드 엑셀파일 작성완료")
        return total_path
    except Exception as e:
        print(f"[collect-total]⚠️ 저장 실패: {e}")
        return None


def step_collect_total(cfg: Config) -> int:
    print("[collect-total] 전체 통합 파일 생성")
    root = cfg.out_grouped_by_col4
    if not root.exists():
        print(f"[collect-total][오류] 폴더 없음: {root.resolve()}")
        return 1
    collect_to_root(root, "total_final_result.xlsx")
    return 0


# ──────────────────────────────────────────────────────────────────────
# ★ 신규 단계: total_final_result 후 추가 분석/강조 표시
# ──────────────────────────────────────────────────────────────────────
def step_post_analyze_and_highlight(cfg: Config) -> int:
    """
    total_final_result.xlsx 생성 후 추가 분석/강조 표시:
      1) delta(2m)-22m 검사 수행
         - 23번째 열(0-based index=22)을 숫자로 변환
         - 최댓값/최솟값을 찾고, 해당 셀만 빨간색 표시
         - 각 값에 대응하는 2번째 열(0-based index=1)의 값을 함께 콘솔에 출력
      2) cladding dia 검사 수행
         - 25, 26번째 열(0-based index=24, 25)을 숫자로 변환
         - 124.3 미만, 125.7 초과인 값만 빨간색으로 표시
         - 이상값 발견 시 "이상값 발견" 및 해당 행의 2번째 열 값 출력
      3) 결과는 grouped_by_col4/total_final_result_annotated.xlsx 로 저장
    """
    root = cfg.out_grouped_by_col4
    total_xlsx = root / "total_final_result.xlsx"
    if not total_xlsx.exists():
        print("[post-analyze][오류] 통합 파일이 없습니다:", total_xlsx.resolve())
        return 1

    try:
        df = pd.read_excel(total_xlsx, engine="openpyxl")
    except Exception as e:
        print(f"[post-analyze][오류] 통합 파일 읽기 실패: {e}")
        return 1

    work = df.copy()

    # 열 인덱스(0-based)
    COL_SECOND = 1          # (n,1) = 2번째 열
    COL_DELTA = 22          # 23번째 열: delta(2m)-22m
    COL_CLAD_IE = 24        # 25번째 열: Clad Dia. I/E
    COL_CLAD_OE = 25        # 26번째 열: Clad Dia. O/E

    max_needed = max(COL_SECOND, COL_DELTA, COL_CLAD_IE, COL_CLAD_OE)
    if work.shape[1] <= max_needed:
        print(f"[post-analyze][오류] 열 수가 부족합니다. (현재 {work.shape[1]}열, 필요 {max_needed+1}열)")
        return 1

    # 숫자 변환
    s_delta = pd.to_numeric(work.iloc[:, COL_DELTA], errors="coerce")
    s_clad_ie = pd.to_numeric(work.iloc[:, COL_CLAD_IE], errors="coerce")
    s_clad_oe = pd.to_numeric(work.iloc[:, COL_CLAD_OE], errors="coerce")

    # ── 콘솔 출력 ────────────────────────────────────────────
    print("결과를 분석합니다.")
    print("1. delta(2m)-22m 검사 수행")

    valid_delta = s_delta.dropna()
    min_idx_list: List[int] = []
    max_idx_list: List[int] = []
    if valid_delta.empty:
        print("[post-analyze] delta(2m)-22m 유효 데이터가 없습니다.")
    else:
        min_val = valid_delta.min()
        max_val = valid_delta.max()
        min_idx_list = s_delta.index[s_delta == min_val].tolist()
        max_idx_list = s_delta.index[s_delta == max_val].tolist()

        print("delta(2m)-22m의 최댓값, 최솟값은 다음과 같습니다.")
        for ridx in min_idx_list:
            sec_val = work.iat[ridx, COL_SECOND]
            print(f"  · 최솟값: {min_val}  |  2번째 열 값: {sec_val}")
        for ridx in max_idx_list:
            sec_val = work.iat[ridx, COL_SECOND]
            print(f"  · 최댓값: {max_val}  |  2번째 열 값: {sec_val}")

    print()
    print("2. cladding dia 검사 수행")
    LOW, HIGH = 124.3, 125.7

    ie_out_mask = (s_clad_ie < LOW) | (s_clad_ie > HIGH)
    oe_out_mask = (s_clad_oe < LOW) | (s_clad_oe > HIGH)

    any_abnormal = False
    for ridx in ie_out_mask[ie_out_mask].index.tolist():
        any_abnormal = True
        sec_val = work.iat[ridx, COL_SECOND]
        val = s_clad_ie.iat[ridx]
        print(f"이상값 발견: Clad Dia. I/E = {val} (행 {ridx})  |  2번째 열 값: {sec_val}")
    for ridx in oe_out_mask[oe_out_mask].index.tolist():
        any_abnormal = True
        sec_val = work.iat[ridx, COL_SECOND]
        val = s_clad_oe.iat[ridx]
        print(f"이상값 발견: Clad Dia. O/E = {val} (행 {ridx})  |  2번째 열 값: {sec_val}")

    if not any_abnormal:
        print("이상값 없음")

    # ── 스타일 적용 준비 (빨간 글자색) ──────────────────────
    style_df = pd.DataFrame("", index=work.index, columns=work.columns)

    for ridx in min_idx_list:
        style_df.iat[ridx, COL_DELTA] = "color: red;"
    for ridx in max_idx_list:
        style_df.iat[ridx, COL_DELTA] = "color: red;"

    for ridx in ie_out_mask[ie_out_mask].index.tolist():
        style_df.iat[ridx, COL_CLAD_IE] = "color: red;"
    for ridx in oe_out_mask[oe_out_mask].index.tolist():
        style_df.iat[ridx, COL_CLAD_OE] = "color: red;"

    annotated_path = root / "total_final_result_annotated.xlsx"
    try:
        styler = work.style.apply(lambda _: style_df, axis=None)
        styler.to_excel(annotated_path, index=False, engine="openpyxl")
        print(f"[post-analyze] 스타일 적용 파일 저장: {annotated_path.name}")
    except Exception as e:
        print(f"[post-analyze][경고] 스타일 적용 저장 실패: {e}")
        try:
            work.to_excel(annotated_path, index=False, engine="openpyxl")
            print(f"[post-analyze] 데이터만 저장 완료(스타일 미포함): {annotated_path.name}")
        except Exception as e2:
            print(f"[post-analyze][오류] 데이터 저장도 실패: {e2}")
            return 1

    return 0


# ──────────────────────────────────────────────────────────────────────
# 실행 엔진
# ──────────────────────────────────────────────────────────────────────

STEPS: Dict[str, Tuple[str, callable]] = {
    "resin":        ("레진/접두어 분석 및 폴더 구성 (ab.xlsx)", step_resin_analyze_and_group_ab),
    "zero":         ("0 → 빈칸 정리 (alls.xlsx → alls_cleaned.xlsx)", step_zero_to_blank_all),
    "group":        ("3/4열 기반 그룹 저장 + 평균행 추가", step_group_by_col4_with_prefix_and_avg),
    "collect-avg":  ("접두어별 평균행 통합 파일 생성 (<코드>.xlsx)", step_collect_all_prefix_averages),
    "copy-42":      ("추가 보정: 4번째 열 → 2번째 열 복사(문자열)", step_copy_col4_to_col2_in_prefix_books),
    "types":        ("타입/제조사 보유 요약 출력", step_summarize_types),
    "reports":      ("하위 폴더별 *_final_result_report.xlsx 생성", step_build_reports),
    "collect-total":("모든 리포트 통합 (total_final_result.xlsx)", step_collect_total),

    # ★ 신규 단계 등록
    "post-analyze": ("최종 결과 추가 분석/강조 표시", step_post_analyze_and_highlight),
}

DEFAULT_ORDER = [
    "resin",
    "zero",
    "group",
    "collect-avg",
    "copy-42",
    "types",
    "reports",
    "collect-total",
    # ★ collect-total 이후 자동 실행
    "post-analyze",
]


def run_steps(step_keys: Iterable[str], cfg: Config) -> int:
    env = setup_utf8_console_and_env()
    logger = Logger(cfg)

    print("안녕하십니까? 통신연구소 소속 김희서 연구원입니다.")
    print("광섬유 특성 분석을 효율적으로 진행하기 위해 통합 파이프라인을 실행합니다.")
    print("ab.xlsx파일 - Draw 공정실적 조회 값, alls.xlsx파일 - 측정 실적 조회 값")
    print(f"[log] 화면+파일 동시 기록: {logger.log_path}")
    print(f"=== 파이썬 실행 파일: {sys.executable}")
    print(f"=== 작업 디렉터리: {Path(__file__).resolve().parent}")
    print("=== 실행을 시작합니다.\n")

    total = len(list(step_keys))
    executed = 0
    failed: List[Tuple[str, int]] = []

    for i, key in enumerate(step_keys, start=1):
        executed += 1
        title, fn = STEPS[key]
        tag = f"[{i}/{total}]"
        start_ts = time.perf_counter()
        print(f"{tag} {key} 시작 | {title} | {datetime.now():%Y-%m-%d %H:%M:%S}")
        print("-" * 100)
        try:
            rc = fn(cfg)
        except Exception as e:  # pragma: no cover
            rc = 1
            print(f"[오류] 단계 실행 중 예외: {e}")
        elapsed = time.perf_counter() - start_ts
        print("-" * 100)
        status = "성공" if rc == 0 else f"실패(rc={rc})"
        print(f"{tag} {key} 종료 | {status} | 소요 {elapsed:.2f}s\n")

        if rc != 0:
            failed.append((key, rc))
            if cfg.stop_on_error:
                print(f"[중단] {key} 실패로 이후 작업을 중단합니다.")
                break

    print("=== 실행 요약 ===")
    print(f"- 총 스텝: {total}")
    print(f"- 성공: {executed - len(failed)}")
    print(f"- 실패: {len(failed)}")
    for name, rc in failed:
        print(f"  · {name}: rc={rc}")

    logger.close()
    return 0 if not failed else 1


# ──────────────────────────────────────────────────────────────────────
# CLI  (★ 공통 옵션을 메인 파서에 추가한 버전)
# ──────────────────────────────────────────────────────────────────────

def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="통합 광섬유 데이터 분석 파이프라인",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    # ✅ 공통 옵션: 어떤 서브커맨드(run-all/run)로 실행하든 ns에 항상 존재
    p.add_argument("--ab", dest="excel_ab", type=Path, default=CFG.excel_ab, help="ab.xlsx 경로")
    p.add_argument("--alls", dest="excel_alls", type=Path, default=CFG.excel_alls, help="alls.xlsx 경로")
    p.add_argument("--alls-cleaned", dest="excel_alls_cleaned", type=Path, default=CFG.excel_alls_cleaned, help="alls_cleaned.xlsx 경로")
    p.add_argument("--out-prefix", dest="out_grouped_by_prefix", type=Path, default=CFG.out_grouped_by_prefix, help="grouped_by_prefix 출력 폴더")
    p.add_argument("--out-col4", dest="out_grouped_by_col4", type=Path, default=CFG.out_grouped_by_col4, help="grouped_by_col4 출력 폴더")
    p.add_argument("--resin-col", dest="resin_col_idx", type=int, default=CFG.resin_col_idx, help="ab.xlsx 레진 열 인덱스(0-based)")
    p.add_argument("--drawno-col", dest="drawno_col_idx", type=int, default=CFG.drawno_col_idx, help="ab.xlsx draw_no 열 인덱스(0-based)")
    p.add_argument("--col3", dest="col3_idx", type=int, default=CFG.col3_idx, help="alls_cleaned.xlsx 3번째 열 인덱스(0-based)")
    p.add_argument("--col4", dest="col4_idx", type=int, default=CFG.col4_idx, help="alls_cleaned.xlsx 4번째 열 인덱스(0-based)")
    p.add_argument("--use-wpattern-first", action="store_true", help="접두 추출 시 W-패턴 우선")
    p.add_argument("--no-second-last-zero-filter", action="store_true", help="C열의 뒤에서 2번째=0 필터 비활성화")
    p.add_argument("--no-stop-on-error", action="store_true", help="오류 발생해도 계속 진행")

    sub = p.add_subparsers(dest="cmd")

    # run-all (기본 순서 전체 실행)
    sub.add_parser("run-all", help="전체 파이프라인 실행")

    # 개별/복수 단계 실행: positional 로 단계 키 나열
    sp_some = sub.add_parser("run", help="지정한 단계만 실행")
    sp_some.add_argument("steps", nargs="+", choices=list(STEPS.keys()), help="실행할 단계 키(여러 개 지정 가능)")

    # 기본은 run-all로 동작
    p.set_defaults(cmd="run-all")
    return p.parse_args(argv)


def args_to_config(ns: argparse.Namespace) -> Config:
    cfg = Config(
        excel_ab=ns.excel_ab,
        excel_alls=ns.excel_alls,
        excel_alls_cleaned=ns.excel_alls_cleaned,
        out_grouped_by_prefix=ns.out_grouped_by_prefix,
        out_grouped_by_col4=ns.out_grouped_by_col4,
        resin_col_idx=ns.resin_col_idx,
        drawno_col_idx=ns.drawno_col_idx,
        col3_idx=ns.col3_idx,
        col4_idx=ns.col4_idx,
        use_w_pattern_first=bool(ns.use_wpattern_first),
        filter_second_last_zero=not bool(ns.no_second_last_zero_filter),
        stop_on_error=not bool(ns.no_stop_on_error),
    )
    return cfg


def main(argv: Optional[List[str]] = None) -> int:
    ns = parse_args(argv)
    cfg = args_to_config(ns)

    if ns.cmd == "run-all":
        steps = DEFAULT_ORDER
    elif ns.cmd == "run":
        steps = ns.steps
    else:
        steps = DEFAULT_ORDER

    # 존재하지 않는 단계 키 방지(방어)
    steps = [s for s in steps if s in STEPS]
    if not steps:
        print("[오류] 실행할 단계가 없습니다.")
        return 1

    return run_steps(steps, cfg)


if __name__ == "__main__":
    raise SystemExit(main())
