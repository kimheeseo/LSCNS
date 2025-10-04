# analyze_and_group_ab.py
import pandas as pd
from pathlib import Path
import re

# ===== 사용자 설정 =====
EXCEL_PATH = Path("ab.xlsx")   # 엑셀 파일 경로
RESIN_COL_IDX = 4              # 0-based: 5번째 열(E열)
DRAWNO_COL_IDX = 0             # 0-based: 1번째 열(A열)
BASE_OUTPUT_DIR = Path("grouped_by_prefix")  # 접두어 기준 폴더를 만들 부모 폴더

def normalize_str(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None

def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH.resolve()}")

    # 엑셀 로드
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")

    # ===== 1) 레진 타입(5번째 열) 집계 =====
    if RESIN_COL_IDX >= df.shape[1]:
        raise IndexError(f"5번째 열(인덱스 {RESIN_COL_IDX})이 없습니다. 실제 열 수: {df.shape[1]}")

    resin_series = (
        df.iloc[:, RESIN_COL_IDX]
          .map(normalize_str)
          .dropna()
          .map(lambda s: s.upper())
    )

    if len(resin_series) == 0:
        print("유효한 레진 타입 값을 찾지 못했습니다.")
    else:
        resin_counts = resin_series.value_counts().sort_index()
        types_str = ",".join(resin_counts.index)
        print(f"레진 타입은 {types_str}으로 구성되어 있습니다.")
        for t, c in resin_counts.items():
            print(f"레진 타입 {t}는 총 {c}개입니다.")

    # ===== 2) 1번째 열의 draw_no에서 앞 3글자 접두어로 폴더 구성 =====
    if DRAWNO_COL_IDX >= df.shape[1]:
        raise IndexError(f"1번째 열(인덱스 {DRAWNO_COL_IDX})이 없습니다. 실제 열 수: {df.shape[1]}")

    draw_series = df.iloc[:, DRAWNO_COL_IDX].map(normalize_str).dropna()

    # 유효한 형식만 사용(영문/숫자/대시/언더스코어만 허용; 필요시 규칙 수정 가능)
    safe_name = re.compile(r"^[A-Za-z0-9_\-\.]+$")

    # 접두어 -> 해당 접두로 시작하는 전체 draw_no 집합
    prefix_map = {}
    for val in draw_series:
        if len(val) < 3:
            continue
        if not safe_name.match(val):
            # 파일/폴더명으로 부적합하면 스킵(필요시 치환 로직으로 변경 가능)
            continue
        prefix = val[:3]
        if not safe_name.match(prefix):
            continue
        prefix_map.setdefault(prefix, set()).add(val)

    # 폴더 생성
    BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for prefix, fullset in prefix_map.items():
        prefix_dir = BASE_OUTPUT_DIR / prefix
        prefix_dir.mkdir(exist_ok=True)
        for full in sorted(fullset):
            (prefix_dir / full).mkdir(exist_ok=True)

    # ===== 3) 접두어(앞 3글자) 집계 출력 =====
    if prefix_map:
        prefix_list = ",".join(sorted(prefix_map.keys()))
        print(" ")
        print(f"해당 측정 실적 조회 결과는 {prefix_list} 등으로 구성되어 있습니다.")
        for prefix in sorted(prefix_map.keys()):
            cnt = len(prefix_map[prefix])
            print(f"{prefix}는 총 {cnt}개의 draw_no가 조회됩니다.")
    else:
        print("접두어(앞 3글자) 기반으로 생성할 폴더 대상이 없습니다.")

    # (선택) 요약 CSV 저장
    #  - resin_type_counts.csv: 레진 타입 집계
    #  - prefix_drawno_counts.csv: 접두어별 draw_no 개수
    try:
        if len(resin_series) > 0:
            resin_counts.to_frame("count").to_csv(EXCEL_PATH.with_name("resin_type_counts.csv"))
        if prefix_map:
            import csv
            out_csv = EXCEL_PATH.with_name("prefix_drawno_counts.csv")
            with out_csv.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["prefix", "draw_no_count"])
                for p in sorted(prefix_map.keys()):
                    w.writerow([p, len(prefix_map[p])])
    except Exception as e:
        # 파일 저장 실패하더라도 콘솔 출력은 이미 완료됨
        print(f"(경고) 요약 CSV 저장 중 문제가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
