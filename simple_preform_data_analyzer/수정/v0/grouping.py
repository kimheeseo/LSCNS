import pandas as pd
import os

def run_grouping(input_folder, output_base_folder):
    """
    input_folder: 'grouped_by_prefix'
    output_base_folder: 'grouped_by_prefix_split'
    """
    os.makedirs(output_base_folder, exist_ok=True)
    results = []
    for file in os.listdir(input_folder):
        if not file.endswith(".xlsx"):
            continue
        file_path = os.path.join(input_folder, file)
        prefix = os.path.splitext(file)[0]

        df = pd.read_excel(file_path, header=None)
        header_row = df.iloc[[0]]
        df_body = df.iloc[1:].reset_index(drop=True)

        if df_body.shape[1] <= 3:
            print(f"⚠ '{file}'는 4번째 열이 없어 건너뜁니다.")
            continue

        df_body['__preform__'] = df_body[3].astype(str)
        unique_preforms = df_body['__preform__'].unique()

        output_folder = os.path.join(output_base_folder, prefix)
        os.makedirs(output_folder, exist_ok=True)

        for val in unique_preforms:
            group = df_body[df_body['__preform__'] == val].drop(columns='__preform__')
            group = group.replace("0", "")
            result_df = pd.concat([header_row, group], ignore_index=True)
            save_path = os.path.join(output_folder, f"{val}.xlsx")
            result_df.to_excel(save_path, index=False, header=False)
            results.append(save_path)
        print(f"✅ '{file}' 처리 완료 → {len(unique_preforms)}개 파일 저장됨")
    return results
