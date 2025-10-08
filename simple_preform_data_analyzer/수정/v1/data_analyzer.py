# auto_analyzer_and_total.py
import os
from pathlib import Path
import pandas as pd

# ====== 설정 ======
ROOT = Path("grouped_by_col4")          # 하위에 20M, L0E, Z0M ... 폴더가 존재
TOTAL_FILENAME = "total_final_result.xlsx"

# (출력 열 제목, 입력엑셀 열 인덱스(0-based), 계산식(None/'delta'/'mac'/'scale'), 스케일 계수)
COLUMN_INFO = [
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
    ("delta 2m-22m", None, "delta", None),   # (Cutoff 2m O/E) - (Cutoff 22m)
    ("Mac value", None, "mac", None),        # (MFD 1310 O/E / Cutoff 2m I/E) * 1000
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

def _pick_input_file(subfolder: Path) -> Path | None:
    """우선순위: <폴더명>.xlsx > final.xlsx"""
    c1 = subfolder / f"{subfolder.name}.xlsx"
    c2 = subfolder / "final.xlsx"
    if c1.exists(): return c1
    if c2.exists(): return c2
    return None

def _safe_series(df: pd.DataFrame, col: int) -> pd.Series:
    """3번째 줄부터(slice 2:) 안전 가져오기"""
    if col is None or col >= df.shape[1]:
        return pd.Series([], dtype="float64")
    return df.iloc[2:, col]

def build_folder_report(subfolder: Path) -> Path | None:
    """
    subfolder 안에서 입력 파일을 찾아 가공하여
    subfolder/<폴더명>_final_result_report.xlsx 생성 후 경로 반환
    """
    src = _pick_input_file(subfolder)
    if src is None:
        print(f"❌ 입력 없음: {subfolder.name} ( <폴더명>.xlsx / final.xlsx )")
        return None

    try:
        df = pd.read_excel(src, header=None, engine="openpyxl")
    except Exception as e:
        print(f"⚠️ 읽기 오류({subfolder.name}): {e}")
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
            # (Cutoff 2m O/E) - (Cutoff 22m)
            col_20 = pd.to_numeric(out.iloc[1:, 19], errors="coerce")  # 19: Cutoff 2m O/E
            col_21 = pd.to_numeric(out.iloc[1:, 20], errors="coerce")  # 20: Cutoff 22m
            delta = (col_20 - col_21).round(4)
            for r, v in enumerate(delta.tolist(), start=1):
                out.loc[r, out_idx] = "" if pd.isna(v) else v

        elif calc == "mac":
            # (MFD 1310 O/E / Cutoff 2m I/E) * 1000
            mfd_oe = pd.to_numeric(out.iloc[1:, 11], errors="coerce")  # 11: MFD 1310 O/E
            cut_ie = pd.to_numeric(out.iloc[1:, 18], errors="coerce")  # 18: Cutoff 2m I/E
            mac = (mfd_oe / cut_ie * 1000).round(2)
            for r, v in enumerate(mac.tolist(), start=1):
                out.loc[r, out_idx] = "" if pd.isna(v) else v

        elif calc == "scale":
            series = _safe_series(df, src_col)
            scaled = (pd.to_numeric(series, errors="coerce") * (factor or 1.0)).round(4)
            for r, v in enumerate(scaled.tolist()):
                out.loc[r + 1, out_idx] = "" if pd.isna(v) else v

    dst = subfolder / f"{subfolder.name}_final_result_report.xlsx"  # 하위 폴더 안에 저장
    try:
        out.to_excel(dst, index=False, header=False, engine="openpyxl")
        print(f"✅ 저장 완료: {dst}")
        return dst
    except Exception as e:
        print(f"⚠️ 저장 오류({subfolder.name}): {e}")
        return None

def collect_to_root(root: Path, total_filename: str = TOTAL_FILENAME) -> Path | None:
    """모든 *_final_result_report.xlsx를 모아 ROOT에 total_final_result.xlsx로 저장"""
    report_paths = sorted(root.glob("*/*_final_result_report.xlsx"))
    if not report_paths:
        print("⚠️ 통합할 리포트가 없습니다.")
        return None

    merged_list = []
    for p in report_paths:
        try:
            raw = pd.read_excel(p, header=None, engine="openpyxl")
            if raw.empty:
                continue
            headers = raw.iloc[0].tolist()
            data = raw.iloc[1:].reset_index(drop=True)
            data.columns = headers  # 첫 행을 헤더로
            # 출처 폴더 표시(선택)
            data.insert(0, "GROUP", p.parent.name)
            merged_list.append(data)
        except Exception as e:
            print(f"⚠️ 통합 중 읽기 오류: {p} -> {e}")

    if not merged_list:
        print("⚠️ 통합할 유효 데이터가 없습니다.")
        return None

    total_df = pd.concat(merged_list, ignore_index=True, sort=False)

    total_path = root / total_filename  # ROOT 에 저장
    try:
        total_df.to_excel(total_path, index=False, engine="openpyxl")
        print(f"📦 통합 저장: {total_path}")
        print("통합모드 엑셀파일 작성완료")
        return total_path
    except Exception as e:
        print(f"⚠️ 통합 저장 실패: {e}")
        return None

def main():
    if not ROOT.exists():
        raise FileNotFoundError(f"폴더가 없습니다: {ROOT.resolve()}")

    # 1) 각 폴더 보고서 생성
    subfolders = [p for p in ROOT.iterdir() if p.is_dir() and not p.name.startswith(("~$", "."))]
    if not subfolders:
        print("처리할 하위 폴더가 없습니다.")
    else:
        for sub in sorted(subfolders, key=lambda x: x.name):
            build_folder_report(sub)

    # 2) 통합 파일을 ROOT에 생성
    collect_to_root(ROOT, TOTAL_FILENAME)

if __name__ == "__main__":
    main()
