import pandas as pd
import os

def normalize_prefix(user_prefix):
    return user_prefix.upper().replace("O", "0")

def run_classification(prefix_list, input_excel="alls.xlsx", output_dir="grouped_by_prefix"):
    """
    prefix_list: ["20M", "L0E", ...] (이미 정규화된 상태로 전달 권장)
    input_excel: 엑셀 원본 파일명
    output_dir: 결과 폴더명
    """
    preform_prefix_list = [normalize_prefix(p) for p in prefix_list]
    if not preform_prefix_list:
        raise ValueError("입력된 프리폼 접두사가 없습니다.")

    if not os.path.exists(input_excel):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {input_excel}")

    df_all = pd.read_excel(input_excel, header=None)
    header_row = df_all.iloc[[0]]
    df_data = df_all.iloc[1:]

    filtered_rows = []
    for idx in range(len(df_data)):
        val_col3 = df_data.iloc[idx, 2]
        if isinstance(val_col3, str) and len(val_col3) >= 2:
            if val_col3[-2] == '0':
                filtered_rows.append(df_data.iloc[idx])

    if not filtered_rows:
        raise ValueError("조건(3번째 열의 뒤에서 두 번째 문자가 '0')에 맞는 행이 없습니다.")

    df_filtered = pd.DataFrame(filtered_rows)
    df_final = pd.concat([header_row, df_filtered], ignore_index=True)
    df_final.to_excel("remove_not_zero.xlsx", index=False, header=False)

    df = pd.read_excel("remove_not_zero.xlsx", header=None)
    header = df.iloc[[0]]
    df_body = df.iloc[1:]

    os.makedirs(output_dir, exist_ok=True)
    saved_files = []

    for prefix in preform_prefix_list:
        group = df_body[df_body[3].astype(str).str[:3].str.upper() == prefix]
        if not group.empty:
            df_prefixed = pd.concat([header, group], ignore_index=True)
            filename = os.path.join(output_dir, f"{prefix}.xlsx")
            df_prefixed.to_excel(filename, index=False, header=False)
            saved_files.append(filename)

    return saved_files
