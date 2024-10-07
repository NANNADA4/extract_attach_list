"""
main 함수
"""

import os


from module.process_folder import process_folder


def main():
    """main 함수. PDF의 4단계에서 별도제출자료 리스트를 추출합니다"""
    print("-"*24)
    print("\n>>>>>>별도제출자료 리스트 추출 (구축세부내역)<<<<<<\n")
    print("-"*24)
    input_path = input("폴더 경로를 입력하세요 (종료는 0을 입력) : ")

    if input_path == '0':
        return 0

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 한번 확인하세요")
        return main()

    input_path = os.path.join('\\\\?\\', input_path)
    process_folder(input_path)

    print("\n~~~폴더 내 PDF에서 별도제출자료 리스트 추출이 완료되었습니다~~~")

    return main()


if __name__ == "__main__":
    main()
