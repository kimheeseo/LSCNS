import pandas as pd
import os
from collections import Counter

def run_preform(choice, input_excel="ab.xlsx", desktop_path=None):
    """
    choice: "1" or "2"
    input_excel: ab.xlsx (default)
    desktop_path: 사용자 바탕화면 경로 (default: 현재 사용자)
    """
    if desktop_path is None:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    if choice == "1":
        folder_name = "month_rit_number"
        column_title = "rit_no"
        prefix_length = 3
    elif choice == "2":
        folder_name = "month_resin_type"
        column_title = "resin_type"
        prefix_length = 1
    else:
        raise ValueError("잘못된 choice 값입니다. 1 또는 2만 허용.")

    folder_path = os.path.join(desktop_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    file_path = os.path.join(desktop_path, input_excel)

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {file_path}")

    df = pd.read_excel(file_path, header=None, engine='openpyxl')
    df = df.drop_duplicates(subset=1).reset_index(drop=True)  # rit_no 중복 제거
    rows, _ = df.shape

    # 월별 데이터 및 카운터 초기화
    month_data = {f'mon_{str(m).zfill(2)}': [] for m in range(1, 13)}
    month_count = {f'mon_{str(m).zfill(2)}': Counter() for m in range(1, 13)}

    for col in range(rows):
        cell_value = str(df.iloc[col, 2])
        if cell_value.startswith("2025") and len(cell_value) >= 6:
            month = cell_value[4:6]
            key = f'mon_{month}'
            if key in month_data:
                if column_title == 'rit_no':
                    value_main = str(df.iloc[col, 1])
                    value_work = str(df.iloc[col, 3])
                    prefix = value_main[:prefix_length]
                else:
                    value_main = str(df.iloc[col, 4])
                    value_work = str(df.iloc[col, 3])
                    prefix = value_main[:prefix_length]
                month_data[key].append([value_main, value_work])
                month_count[key][prefix] += 1

    # 요약 정보
    summary_lines = []
    summary_lines.append(f"📊 {column_title} 월별 분포 요약:\n")
    for month in sorted(month_count.keys()):
        counts = month_count[month]
        if counts:
            summary_lines.append(f"▶ {month}:")
            for prefix, cnt in counts.items():
                summary_lines.append(f"  - {prefix}: {cnt}개")
        else:
            summary_lines.append(f"▶ {month}: 없음")

    output_path = os.path.join(folder_path, f"mon_split_{column_title}.xlsx")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_summary = pd.DataFrame(summary_lines, columns=[f"{column_title}_summary"])
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        for month, data in month_data.items():
            if data:
                df_month = pd.DataFrame(data, columns=[column_title, "work_time"])
            else:
                df_month = pd.DataFrame(columns=[column_title, "work_time"])
            df_month.to_excel(writer, sheet_name=month, index=False)

    return output_path, summary_lines
