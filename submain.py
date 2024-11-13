import os

desktop_folder_path = "C:/Users/djdjd/OneDrive/바탕 화면/카카오톡다운로드" 

def rename_images_in_order(folder_path):
    # 폴더의 모든 파일 목록 가져오기
    files = os.listdir(folder_path)
    
    # 파일을 수정 시간 기준으로 정렬
    files.sort(key=lambda file: os.path.getmtime(os.path.join(folder_path, file)))
    
    # 파일명 순차적으로 변경
    for index, filename in enumerate(files, start=1):
        # 확장자 분리
        ext = os.path.splitext(filename)[1]
        # 새로운 파일명 생성
        new_name = f"{index}{ext}"
        # 파일 경로 변경
        os.rename(os.path.join(folder_path, filename), os.path.join(folder_path, new_name))

    print("파일 이름이 순서대로 변경되었습니다.")

rename_images_in_order(desktop_folder_path)
    