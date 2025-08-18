import os
import pandas as pd

def run_auto_analyzer(root_folder):
    """
    root_folder: grouped_by_prefix_split
    """
    # (출력 열 이름, final 엑셀 열 번호, 계산식 lambda 또는 None, 곱셈 계수 또는 None)
    column_info = [
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
        ("delta 2m-22m", None, 'delta', None),
        ("Mac value", None, 'mac', None),
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
        ("R7.5mm 1t 1550", 26, 'scale', 0.1),
        ("R7.5mm 1t 1625", 69, 'scale', 0.1),
        ("R10mm 1t 1550", 70, 'scale', 0.1),
        ("R10mm 1t 1625", 71, 'scale', 0.1),
        ("R15mm 10t 1550", 81, 'scale', 0.5),
        ("R15mm 10t 1625", 82, 'scale', 0.5),
    ]

    for subfolder in os.listdir(root_folder):
        subfolder_path = os.path.join(root_folder, subfolder)
        if not os.path.isdir(subfolder_path):
            continue

        final_files = [f for f in os.listdir(subfolder_path) if f.endswith("final.xlsx")]
        if not final_files:
            print(f"❌ '{subfolder}' 폴더에는 final.xlsx 파일이 없습니다.")
            continue

        final_path = os.path.join(subfolder_path, final_files[0])
        try:
            df = pd.read_excel(final_path, header=None, engine='openpyxl')
            df_result = pd.DataFrame()
            n_rows = len(df) - 2  # 3번째 줄부터 시작

            for col_idx, (title, src_col, calc_type, factor) in enumerate(column_info):
                df_result.loc[0, col_idx] = title  # 첫 행 제목

                if calc_type is None:
                    if src_col is not None:
                        col_values = df.iloc[2:, src_col].tolist()
                        for row_idx, val in enumerate(col_values):
                            df_result.loc[row_idx + 1, col_idx] = val
                elif calc_type == 'delta':
                    col_20 = df_result.iloc[1:, 19].astype(float)
                    col_21 = df_result.iloc[1:, 20].astype(float)
                    df_result.iloc[1:, col_idx] = (col_20 - col_21).round(4)
                elif calc_type == 'mac':
                    col_12 = pd.to_numeric(df_result.iloc[1:, 11], errors='coerce')
                    col_19 = pd.to_numeric(df_result.iloc[1:, 18], errors='coerce')
                    df_result.iloc[1:, col_idx] = ((col_12 / col_19) * 1000).round(2)
                elif calc_type == 'scale':
                    raw_values = df.iloc[2:, src_col]
                    scaled_values = []
                    for val in raw_values:
                        try:
                            scaled_values.append(round(float(val) * factor, 4))
                        except:
                            scaled_values.append("")
                    for row_idx, val in enumerate(scaled_values):
                        df_result.loc[row_idx + 1, col_idx] = val

            save_path = os.path.join(subfolder_path, f"{subfolder}_final_result_report.xlsx")
            df_result.to_excel(save_path, index=False, header=False)
            print(f"✅ 저장 완료: {save_path}")

        except Exception as e:
            print(f"⚠️ 오류 발생 ({subfolder}): {e}")
