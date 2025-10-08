# group_by_col4_with_prefix_and_avg.py
import pandas as pd
import numpy as np
from pathlib import Path
import re
from collections import defaultdict

INPUT = Path("alls_cleaned.xlsx")
OUTPUT_DIR = Path("grouped_by_col4")   # 결과 루트 폴더
COL3_IDX = 2  # 3번째 열(C열): 예) "20M25224A00W01A0500100E"
COL4_IDX = 3  # 4번째 열(D열): 예) "20M25224"

def second_last_is_zero(val) -> bool:
    """문자열의 뒤에서 두 번째 문자가 '0'이면 True"""
    if pd.isna(val):
        return False
    s = str(val).strip()
    if len(s) < 2:
        return False
    return s[-2] == "0"

def safe_filename(name: str) -> str:
    """파일명 안전 치환"""
    s = str(name).strip()
    if not s:
        s = "EMPTY"
    return re.sub(r'[^A-Za-z0-9._-]+', "_", s)

def make_avg_row(df: pd.DataFrame) -> dict:
    """
    df의 '모든 데이터 행'을 대상으로 열별 평균을 계산하여 마지막 줄로 추가할 값을 반환.
    - 숫자형/숫자로 변환 가능한 값만 평균 계산
    - 비숫자형은 공란("") 유지
    """
    avg = {}
    for col in df.columns:
        # 숫자로 변환 시도 (문자형 숫자 포함), 변환 불가는 NaN
        s_num = pd.to_numeric(df[col], errors="coerce")
        if s_num.notna().any():
            avg[col] = s_num.mean()
        else:
            avg[col] = ""  # 비숫자형은 빈칸
    return avg

def main():
    print("spoolno2의 단선 등을 의미하는 값 제거합니다.")

    if not INPUT.exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {INPUT.resolve()}")

    df = pd.read_excel(INPUT, engine="openpyxl")

    # 열 체크
    need_max = max(COL3_IDX, COL4_IDX)
    if df.shape[1] <= need_max:
        raise IndexError(
            f"열 개수가 부족합니다. 현재 열 수: {df.shape[1]} (필요: 최소 {need_max+1})"
        )

    col3 = df.columns[COL3_IDX]
    col4 = df.columns[COL4_IDX]

    # (A) 3번째 열에서 '뒤에서 2번째가 0'인 행만 유지
    mask = df[col3].map(second_last_is_zero)
    filtered = df[mask].copy()

    # (B) 4번째 열이 비어있지 않은 행만 대상
    filtered = filtered[filtered[col4].notna() & (filtered[col4].astype(str).str.strip() != "")]
    if filtered.empty:
        print("필터 통과 후 저장할 데이터가 없습니다.")
        return

    # 결과 폴더 준비
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # 접두어별(앞 3글자) 결과 개수 집계
    prefix_counts = defaultdict(int)

    # (C) 4번째 열 값으로 그룹 → 각 그룹별 파일 생성
    for key, g in filtered.groupby(col4, dropna=False):
        key_str = str(key).strip()
        if not key_str:
            continue

        # 파일명 앞 3글자 기준 하위 폴더
        prefix3 = key_str[:3] if len(key_str) >= 3 else "UNK"
        dest_dir = OUTPUT_DIR / prefix3
        dest_dir.mkdir(parents=True, exist_ok=True)

        # 마지막 줄에 평균 행 추가 (엑셀의 2번째 줄부터 마지막 줄까지 = 모든 데이터 행)
        g_out = g.copy()
        avg_row = make_avg_row(g_out)
        g_out = pd.concat([g_out, pd.DataFrame([avg_row])], ignore_index=True)

        # 저장
        out_path = dest_dir / f"{safe_filename(key_str)}.xlsx"
        g_out.to_excel(out_path, index=False, engine="openpyxl")
      #  print(f"[저장] {out_path} (원본 행 {len(g)}개, 평균행 1개 추가)")

        prefix_counts[prefix3] += 1

    # (D) 접두어별 결과 개수 출력 (예: "20M에는 #개의 결과값이 조회됩니다.")
    for pfx in sorted(prefix_counts.keys()):
        print(f"{pfx}에는 {prefix_counts[pfx]}개의 결과값이 조회됩니다.")

if __name__ == "__main__":
    main()
