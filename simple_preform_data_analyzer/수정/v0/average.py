import os
import pandas as pd

def run_average(root_folder):
    """
    root_folder: grouped_by_prefix_split
    """
    save_count = 0
    result_files = []
    for subfolder in os.listdir(root_folder):
        subfolder_path = os.path.join(root_folder, subfolder)
        if not os.path.isdir(subfolder_path):
            continue

        average_rows = []
        header_row = None

        for file in os.listdir(subfolder_path):
            if file.endswith(".xlsx") and not (file.endswith("_zerox.xlsx") or file.endswith("_average.xlsx")):
                file_path = os.path.join(subfolder_path, file)
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    df_zerox = df.where(df != 0, "")
                    filename_wo_ext = os.path.splitext(file)[0]
                    zerox_filename = f"{filename_wo_ext}_zerox.xlsx"
                    zerox_path = os.path.join(subfolder_path, zerox_filename)
                    df_zerox.to_excel(zerox_path, index=False)
                    save_count += 1
                    if save_count % 100 == 0:
                        print(f"✅ 총 {save_count}개 파일 저장 완료")

                    df_zerox = pd.read_excel(zerox_path, engine='openpyxl')
                    col_4_name = df_zerox.columns[3]
                    filtered_df = df_zerox[df_zerox[col_4_name].astype(str) == filename_wo_ext]

                    sum_series = []
                    count_series = []
                    for col in filtered_df.columns:
                        values = pd.to_numeric(filtered_df[col], errors='coerce')
                        values_for_sum = values.fillna(0)
                        values_for_count = values.notna().astype(int)
                        sum_series.append(values_for_sum.sum())
                        count_series.append(values_for_count.sum())

                    avg_values = []
                    for total, count in zip(sum_series, count_series):
                        avg = round(total / count, 4) if count > 0 else ""
                        avg_values.append(avg)

                    average_row = pd.DataFrame([avg_values], columns=df_zerox.columns)
                    average_row = average_row.astype("object")
                    average_row.iat[0, 1] = filename_wo_ext
                    first_row = df_zerox.iloc[[0]]
                    result_df = pd.concat([first_row, average_row], ignore_index=True)
                    average_filename = f"{filename_wo_ext}_average.xlsx"
                    average_path = os.path.join(subfolder_path, average_filename)
                    result_df.to_excel(average_path, index=False)
                except Exception as e:
                    print(f"⚠️ 오류 발생 ({file}): {e}")

        # 평균 행만 모아서 최종 파일 저장
        for file in os.listdir(subfolder_path):
            if file.endswith("_average.xlsx"):
                try:
                    df_avg = pd.read_excel(os.path.join(subfolder_path, file), engine='openpyxl')
                    if header_row is None:
                        header_row = df_avg.iloc[[0]]
                    last_row = df_avg.iloc[[-1]]
                    average_rows.append(last_row)
                except Exception as e:
                    print(f"⚠️ 평균 파일 처리 오류 ({file}): {e}")

        if average_rows and header_row is not None:
            final_df = pd.concat([header_row] + average_rows, ignore_index=True)
            final_filename = f"{subfolder}_final.xlsx"
            final_path = os.path.join(subfolder_path, final_filename)
            final_df.to_excel(final_path, index=False)
            result_files.append(final_path)
            print(f"📦 최종 요약 저장 완료: {final_filename}")
    return result_files
