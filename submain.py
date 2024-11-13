import os
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import * # 셀 꾸미는 모듈
from googleapiclient.discovery import build # 폴더 이미지를 끌어오는 코드
from googleapiclient.http import MediaFileUpload

creds_path = "C:/Users/djdjd/Downloads/sanguine-sign-436506-k7-ccbb794af6c1.json"

# 구글 스프레드시트와 드라이브 접근을 위한 권한 범위 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# 자격 증명 로드
credentials = Credentials.from_service_account_file(creds_path, scopes=SCOPES)

# gspread 클라이언트 생성
client = gspread.authorize(credentials)

spreadsheet_id = '187W2mFAy5fq7bo7d8HQsjBTkP5ue4PvOFn4JxtRdYFs'
spreadsheet = client.open_by_key(spreadsheet_id)
drive_service = build('drive', 'v3', credentials=credentials)

desktop_folder_path = "C:/Users/djdjd/OneDrive/바탕 화면/카카오톡다운로드" 
folder_id = '1CSjt_1nRMZrguO_06RVs4zeBQRt4l8yp'

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
    

def upload_images_to_drive():
    # A 폴더에서 모든 파일을 탐색하며 업로드
    for filename in os.listdir(desktop_folder_path):
        file_path = os.path.join(desktop_folder_path, filename)
        
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, mimetype='image/jpeg')
        drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f'Uploaded {filename} to Google Drive folder.')