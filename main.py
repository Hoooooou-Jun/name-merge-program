import pandas as pd
import os
import sys

if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.abspath(".")

folder_path = os.path.join(base_path, 'excel_data')

name_data = []

print("Excel file 로딩 중...")

for filename in os.listdir(folder_path):
    print(filename + "파일 조회 중..")
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        full_file_path = os.path.join(folder_path, filename)

        data_frame = pd.read_excel(full_file_path, engine='openpyxl')

        for col in data_frame.columns:
            for cell in data_frame[col]:
                if pd.notnull(cell):
                    name_data.append(str(cell))

        print(filename + "파일 조회 완료")

print("Excel file 로딩 완료!")

def save_extract_file():
    file_name = os.path.join(base_path, "extract_data.xlsx")
    extract_data = pd.DataFrame(name_data, columns=["Names"])
    extract_data.to_excel(file_name, index=False, header=False)
    print(f"'{file_name}' 파일로 저장되었습니다.")

save_extract_file()

def main():
    print("\n1. 모든 데이터 출력하기.")
    print("2. 이름 중복 조회 및 제거하기.")
    print("3. 이름 추가하기")
    print("4. 프로그램 종료\n")
    process_input = int(input("숫자를 입력하세요 ; "))

    if process_input == 1:
        print(name_data)
        main()
    elif process_input == 2:
        print("\n중복된 이름을 확인합니다.")
        duplicates = [name for name in set(name_data) if name_data.count(name) > 1]
        print("중복된 이름:", duplicates if duplicates else "중복 없음")

        process_merge_input = int(input("\n중복을 제거할까요?(1: 중복 제거, 2: 취소) ; "))
        if process_merge_input == 1:
            unique_names = list(set(name_data))
            name_data.clear()
            name_data.extend(unique_names)
            print("중복된 이름 병합이 완료되었습니다.")
            save_extract_file()
        elif process_merge_input == 2:
            main()
        main()
    elif process_input == 3:
        print("\n이름을 추가합니다.")
        process_name_input = input("추가할 이름을 입력하세요: ")
        name_data.append(process_name_input)
        print(f"{process_name_input} 이름이 추가되었습니다.")
        save_extract_file()
        main()
    elif process_input == 4:
        print("\n프로그램을 종료합니다.")
        return
    else:
        print("잘못된 입력입니다. 1, 2, 또는 3을 입력해 주세요.")
        main()

main()