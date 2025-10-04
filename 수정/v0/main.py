import os
import time
from preform import run_preform
from classification import run_classification
from grouping import run_grouping
from average import run_average
from auto_analyzer import run_auto_analyzer

def main():
    # 1. 첫 번째 자동 입력
    print("월별 어떤 값을 추출할 것입니까?\n1. rit_no\n2. resin_type")
    time.sleep(2)
    print("1이 자동으로 입력됩니다.")
    result_path_1, summary_1 = run_preform("1")
    print("\n[1번 결과]")
    print("\n".join(summary_1))
    print(f"✔️ preform 완료: {result_path_1}\n")

    # 2. 두 번째 자동 입력
    time.sleep(2)
    print("2이 자동으로 입력됩니다.")
    result_path_2, summary_2 = run_preform("2")
    print("\n[2번 결과]")
    print("\n".join(summary_2))
    print(f"✔️ preform 완료: {result_path_2}\n")
    
    # 3. 관심 프리폼 접두사 입력
    print("\n관심 프리폼 접두사를 하나씩 입력하세요. (끝내려면 'a' 입력)")
    prefixes = []
    while True:
        pf = input("▶ 프리폼 입력: ").strip()
        if pf.lower() == 'a':
            break
        if pf:
            prefixes.append(pf)
    if not prefixes:
        print("❗입력된 프리폼이 없습니다. 프로그램을 종료합니다.")
        return

    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    # 4. classification.py 실행
    print("\n[2/5] classification.py 실행중...")
    classified_files = run_classification(prefixes)
    print(f"✔️ classification 완료: {len(classified_files)}개 파일 분류")

    # 5. grouping.py 실행
    print("\n[3/5] grouping.py 실행중...")
    grouped_files = run_grouping("grouped_by_prefix", "grouped_by_prefix_split")
    print(f"✔️ grouping 완료: {len(grouped_files)}개 파일 생성")

    # 6. average.py 실행
    print("\n[4/5] average.py 실행중...")
    average_files = run_average(os.path.join(desktop_path, "grouped_by_prefix_split"))
    print(f"✔️ average 완료: {len(average_files)}개 요약 파일 생성")

    # 7. auto_analyzer.py 실행
    print("\n[5/5] auto_analyzer.py 실행중...")
    run_auto_analyzer(os.path.join(desktop_path, "grouped_by_prefix_split"))
    print(f"✔️ auto_analyzer 완료")

    print("\n🎉 전체 자동화 파이프라인이 종료되었습니다.")

if __name__ == "__main__":
    main()
