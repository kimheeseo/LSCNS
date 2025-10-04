# zero_to_blank_all.py
import pandas as pd
import numpy as np
from pathlib import Path
import re

INPUT = Path("alls.xlsx")
OUTPUT = Path("alls_cleaned.xlsx")

if not INPUT.exists():
    raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {INPUT.resolve()}")

df = pd.read_excel(INPUT, engine="openpyxl")

# 1) 숫자형(정수/실수) 0 → 빈칸(None)
num_cols = df.select_dtypes(include=[np.number]).columns
# bool 컬럼은 제외(원하면 제거)
bool_cols = df.select_dtypes(include=["bool"]).columns
num_cols = [c for c in num_cols if c not in bool_cols]

for c in num_cols:
    # 값이 0인 곳만 None으로
    df[c] = df[c].mask(df[c] == 0, other=None)

# 2) 문자열 등 비-숫자형에서 '0', '000', '0.0', ' 0 ', '0,000' 등 → 빈칸(None)
#    - 숫자 0 이외의 문자(예: 'A0B')는 그대로 둡니다.
obj_cols = df.columns.difference(num_cols).tolist()

# 정규식: 0 또는 0들, 선택적으로 소수점/콤마 뒤도 전부 0만 있는 경우
zero_like = re.compile(r'^[\+\-]?\s*0+(?:[.,]0+)?\s*$')

for c in obj_cols:
    s = df[c]
    # 문자열로 바꾸고 검사하되, 원래 NaN이면 건드리지 않음
    s_str = s.astype(str)
    mask = s_str.str.match(zero_like, na=False)
    # 단, 진짜 숫자로 읽힌 값인데 우연히 object가 된 경우도 0이면 비움
    # → 안전하게 to_numeric으로 0인지 재확인
    # (문자→숫자 변환 불가면 NaN)
    num_eq_zero = pd.to_numeric(s_str.str.replace(",", ".", regex=False), errors="coerce").eq(0)
    final_mask = mask | num_eq_zero
    # 이미 비어있는(NA-like) 값은 유지
    final_mask = final_mask & s.notna()
    df.loc[final_mask, c] = None

# 저장 (None/NaN은 엑셀에서 빈칸으로 보입니다)
df.to_excel(OUTPUT, index=False, engine="openpyxl")
print(f"완료! 변환된 파일: {OUTPUT.resolve()}")
