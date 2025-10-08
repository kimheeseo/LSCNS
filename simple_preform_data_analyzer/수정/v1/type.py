# summarize_types_from_folders_with_vendors.py
from pathlib import Path

BASE_DIR = Path("grouped_by_col4")

# 분류 매핑
TYPE_MAP = {
    "LWPF(90)": {
        "SEC": ["W00", "W0J"],
        "Sumitomo": ["20M"],  # 사용자 지정
    },
    "LWPF(150)": {
        "SEC": ["L0E"],
        "Sumitomo": ["L0M"],
    },
    "LWPF(180)": {
        "SEC": ["S0E"],
        "Sumitomo": ["S0M"],
    },
    "A1(90)": {
        "SEC": [],
        "Sumitomo": ["Z0M"],
    },
    "A1(150)": {
        "SEC": [],
        "Sumitomo": ["Z0L"],
    },
    "A2(90)": {
        "SEC": ["AJW", "AJF", "AJB"],
        "Sumitomo": [],
    },
    "A2(150)": {
        "SEC": ["AL"],
        "Sumitomo": [],
    },
}

def main():
    if not BASE_DIR.exists():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {BASE_DIR.resolve()}")

    # grouped_by_col4의 직속 하위 폴더명 수집(대소문자 무시, 임시/숨김 제외)
    folder_names = set()
    for p in BASE_DIR.iterdir():
        if p.is_dir():
            name = p.name.strip()
            if name and not name.startswith("~$") and not name.startswith("."):
                folder_names.add(name.upper())

    print("현재 보유한 사항은 다음과 같습니다.")

    # 매핑에 정의된 모든 코드(대문자)
    defined_codes_upper = set()
    for vendors in TYPE_MAP.values():
        for codes in vendors.values():
            defined_codes_upper.update(c.upper() for c in codes)

    matched_present_upper = set()
    any_printed = False

    # 타입별로 제조사 단위 출력
    for type_name, vendors in TYPE_MAP.items():
        vendor_parts = []
        for vendor, codes in vendors.items():
            if not codes:
                continue
            present_codes = [c for c in codes if c.upper() in folder_names]
            if present_codes:
                vendor_parts.append(f"{vendor}=" + ", ".join(present_codes))
                matched_present_upper.update(c.upper() for c in present_codes)

        if vendor_parts:
            any_printed = True
            print(f"타입: {type_name}: " + " / ".join(vendor_parts) + " 보유중입니다.")

    # 기타(매핑에 없는 폴더)
    others = sorted(folder_names - defined_codes_upper)
    if others:
        print("기타: " + ", ".join(others))

    if not any_printed and not others:
        print("(일치하는 타입 코드가 아직 보유되어 있지 않습니다.)")

if __name__ == "__main__":
    main()
