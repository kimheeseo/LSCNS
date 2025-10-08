# collect_all_prefix_averages_progress_v2.py
import pandas as pd
from pathlib import Path

BASE_DIR = Path("grouped_by_col4")   # 접두어별 하위 폴더들이 있는 루트
FOURTH_COL_IDX = 3                   # 0-based: 4번째 열

def is_empty(val):
    """NA 또는 공백 문자열을 빈값으로 간주"""
    if val is None:
        return True
    try:
        import pandas as pd
        if pd.isna(val):
            return True
    except Exception:
        pass
    return str(val).strip() == ""

def candidate_files(prefix_dir: Path):
    """해당 접두 폴더에서 처리 대상 파일 목록 반환 (자기 결과파일/엑셀 임시파일 제외)"""
    out_file = prefix_dir / f"{prefix_dir.name}.xlsx"
    return sorted(
        p for p in prefix_dir.glob("*.xlsx")
        if p.name.lower() != out_file.name.lower() and not p.name.startswith("~$")
    )

def collect_for_prefix(prefix_dir: Path):
    """한 접두 폴더의 각 xlsx 마지막 행(=평균행)만 모아 <prefix>.xlsx 저장"""
    prefix = prefix_dir.name
    out_file = prefix_dir / f"{prefix}.xlsx"

    excel_files = candidate_files(prefix_dir)
    if not excel_files:
        print(f"[INFO] {prefix} 폴더에 수집할 엑셀 파일이 없습니다.")
        return

    last_rows = []
    for path in excel_files:
        try:
            df = pd.read_excel(path, engine="openpyxl")
        except Exception as e:
            print(f"[경고] 읽기 실패: {path.name} -> {e}")
            continue

        if df.empty:
            print(f"[건너뜀] 빈 파일: {path.name}")
            continue

        last_idx = len(df) - 1

        # 마지막 행의 4번째 열이 비어 있으면 직전 행의 4번째 열로 보정
        if df.shape[1] > FOURTH_COL_IDX and last_idx >= 1:
            if is_empty(df.iat[last_idx, FOURTH_COL_IDX]):
                df.iat[last_idx, FOURTH_COL_IDX] = df.iat[last_idx - 1, FOURTH_COL_IDX]

        # 마지막 행만 수집
        last_rows.append(df.iloc[[last_idx]].copy())

    if not last_rows:
        print(f"[INFO] {prefix}: 수집할 평균 행이 없습니다.")
        return

    # 서로 다른 컬럼 구성도 합쳐서 저장
    result = pd.concat(last_rows, ignore_index=True, sort=False)

    try:
        result.to_excel(out_file, index=False, engine="openpyxl")
        print(f"[저장 완료] {out_file.resolve()} (총 {len(result)}행)")
    except Exception as e:
        print(f"[오류] {prefix} 결과 저장 실패: {e}")

def main():
    print("이제 마지막 결과를 출력합니다.")

    if not BASE_DIR.exists():
        raise FileNotFoundError(f"폴더가 없습니다: {BASE_DIR.resolve()}")

    # 보이는 폴더(디렉터리) 개수 기준으로 진행률 표시
    prefix_dirs = sorted(
        p for p in BASE_DIR.iterdir()
        if p.is_dir() and not p.name.startswith("~$")
    )
    if not prefix_dirs:
        print("[INFO] 처리할 접두 폴더가 없습니다.")
        return

    total_prefixes = len(prefix_dirs)   # ← 분모를 '폴더 개수'로!
    for idx, pdir in enumerate(prefix_dirs, start=1):
        print(f"[{idx}/{total_prefixes}] {pdir.name}의 값을 통합합니다.")
        collect_for_prefix(pdir)

if __name__ == "__main__":
    main()
