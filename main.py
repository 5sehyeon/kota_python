import re
import gspread
import time
from google.oauth2.service_account import Credentials
from gspread_formatting import * # 셀 꾸미는 모듈
from googleapiclient.discovery import build # 폴더 이미지를 끌어오는 코드
from gspread_formatting import CellFormat, Color, TextFormat, format_cell_range
from submain import upload_images_to_drive, rename_images_in_order
# '롯대 전주점' 이렇게 쓰게해서, '점'을 붙일 것인가.. '점'을 꼭 써줘야 하나,,
# 자격 증명 파일 경로
creds_path = "C:/Users/djdjd/Downloads/sanguine-sign-436506-k7-ccbb794af6c1.json"

# 구글 스프레드시트와 드라이브 접근을 위한 권한 범위 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# 자격 증명 로드
credentials = Credentials.from_service_account_file(creds_path, scopes=SCOPES)

# gspread 클라이언트 생성
client = gspread.authorize(credentials)

# 스프레드시트 ID로 연결 (스프레드시트 URL에서 ID를 가져옴)
spreadsheet_id = '187W2mFAy5fq7bo7d8HQsjBTkP5ue4PvOFn4JxtRdYFs'
spreadsheet = client.open_by_key(spreadsheet_id)
drive_service = build('drive', 'v3', credentials=credentials)

index_common = 0

desktop_folder_path = "C:/Users/djdjd/OneDrive/바탕 화면/카카오톡다운로드" 
folder_id = '1V2to8AtiGQdugF7WCKAtG-pwmoSY-Lyl'
move_drive_folder_id = '1ndsAt7MTbwxsLKUHu69vijVDqDv9YKpS'

rename_images_in_order(desktop_folder_path)
time.sleep(2)

upload_images_to_drive()
time.sleep(50)


query = f"'{folder_id}' in parents and (mimeType='image/jpeg' or mimeType='image/png')"
results = drive_service.files().list(
    q=query,
    fields="files(id, name)"
).execute()
time.sleep(10)
items = results.get('files', [])
items.sort(key=lambda x: int(re.findall(r'\d+', x['name'])[0]))
item = items[index_common]

time.sleep(10)
print(items)

fmt_전주점 = CellFormat(
    backgroundColor=Color(0.427, 0.620, 0.922),
    textFormat=TextFormat(bold=True)
)

fmt_김해점 = CellFormat(
    backgroundColor=Color(0.576, 0.769, 0.490),
    textFormat=TextFormat(bold=True)
)

fmt_창원점 = CellFormat(
    backgroundColor=Color(1.0, 0.851, 0.400),
    textFormat=TextFormat(bold=True)
)

fmt_잠실점 = CellFormat(
    backgroundColor=Color(0.918, 0.600, 0.600),
    textFormat=TextFormat(bold=True)
)

fmt_기타 = CellFormat(
    backgroundColor=Color(1.0, 0.427, 0.004),
    textFormat=TextFormat(bold=True)
)

fmt_가운데_맞춤 = CellFormat(
    wrapStrategy="WRAP",
    horizontalAlignment="CENTER",   # 가로 가운데 정렬 (가운데 맞춤)
    verticalAlignment="MIDDLE",
)

with open("python_kota/1.txt", "r", encoding="UTF-8") as f:
    text = f.read().replace("\n", " ") # 67, 68번 문장 지우기
    ### 쿄타 삼촌에게 요청할 부분
    pattern = r'--------------- 2024년 11월 4일 월요일 ---------------(.*?)--------------- 2024년 11월 6일 수요일 --------------- ' # 날짜 변경
    text = re.sub(pattern,"",text)
    pattern = r'(\.+날)'
    text = re.sub(pattern,". 날",text)
    pattern = r'(\.+성)'
    text = re.sub(pattern,". 성",text)
    pattern = r'(\.+연)'
    text = re.sub(pattern,". 연",text)
    pattern = r'(\.+수)'
    text = re.sub(pattern,". 수",text)
    pattern = r'(\.+상)'
    text = re.sub(pattern,". 상",text)
    pattern = r'(\.+구)'
    text = re.sub(pattern,". 구",text)
    pattern = r'(\.+옵)'
    text = re.sub(pattern,". 구",text)
    
    ### 쿄타 삼촌에게 요청할 부분
    pattern = r'(\. 2)'
    text = re.sub(pattern," 2",text)
    pattern = r'(\. 3)'
    text = re.sub(pattern," 3",text)
    pattern = r'(\. 4)'
    text = re.sub(pattern," 4",text)
    pattern = r'(\. 5)'
    text = re.sub(pattern," 5",text)
    pattern = r'(\. 6)'
    text = re.sub(pattern," 6",text)
    pattern = r'(\. 7)'
    text = re.sub(pattern," 7",text)
    pattern = r'(\d+),(\d+),(\d+)'
    text = re.sub(pattern,r'\1.\2.\3',text)
    ### 여기까지
    
    pattern = r'(\d+)년 (\d+)월 (\d+)일'
    text = re.sub(pattern,r'\1.\2.\3',text)
    pattern = r'(\d+)년(\d+)월(\d+)일'
    text = re.sub(pattern,r'\1.\2.\3',text)
    pattern = r'(\:+7)'
    text = re.sub(pattern,": 7",text)
    
    ### 쿄타 삼촌에게 부탁할 부분
    pattern = r'(다\.)'
    text = re.sub(pattern,"다",text)
    pattern = r'(요\.)'
    text = re.sub(pattern,"요",text)
    pattern = r'(음\.)'
    text = re.sub(pattern,"음",text)
    pattern = r'(짐\.)'
    text = re.sub(pattern,"짐",text)
    ### 여기까지
    
    result = re.findall(r'\{([^}]*)\}', text)
    #print(result)
    for i in result:
        i_lower = i.lower()
        # 주문건
        out_key = 0
        if "주문건" in i_lower:
            sheet = spreadsheet.get_worksheet(0)
            data = sheet.get_all_values('C:G')
            time.sleep(1)

            # 비어 있는 첫 번째 행 번호 찾기
            def find_first_empty_row(data):
                for i, row in enumerate(data):
                    # 비어 있는 행을 찾으면 해당 행 번호 반환 (1-based index)
                    if all(cell == '' for cell in row):
                        return i + 1  # 인덱스는 0부터 시작하므로 +1
                return len(data) + 1  # 모든 행이 차 있다면 다음 행 번호 반환

            
            first_empty_row = find_first_empty_row(data)
            time.sleep(4)
            print("--주문건--")
            i_lower = i_lower.split('. ')
            print(i_lower)
            count_jumun_1 = 0
            option = i_lower[len(i_lower)-1] #옵션
            lst = option.split(":")
            real_option = lst[len(lst)-1] # 진짜 옵션 내용
            print(real_option)
            sheet.update_cell(first_empty_row,13,real_option)
            f = str(first_empty_row)
            format_cell_range(sheet,('M'+f),fmt_가운데_맞춤)
            for j in i_lower:
                if j[len(j)-1] in ["1", "2", "3", "4", "5", "6", "7", "8"]:
                    result = j[:len(j)-2]
                    if ":" not in result:
                        jumun = result # 주문
                        print(jumun)
                        a = re.findall(r'\(([^)]*)\)', jumun)
                        jijom = a[0]
                        print(jijom)
                        match = re.findall(r'(전주|김해|창원|잠실)',jijom)
                        print(match)
                        try:
                            match[0] == 1
                        except IndexError:
                            b = jijom.split(" ")
                            majun = ""
                            try:
                                b[1]
                            except:
                                majun = b[0]
                                
                                ### 쿄타님에게 추가하는 거 요청하기
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 7,majun)
                                e = str(first_empty_row)
                                format_cell_range(sheet, ('G'+e),fmt_기타)
                                time.sleep(3)
                                ### 여기까지
                                
                            else:
                                for i in range(len(b)):
                                    majun += b[i]
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 7,majun)
                                e = str(first_empty_row)
                                format_cell_range(sheet, ('G'+e),fmt_기타)
                                time.sleep(3)
                        else :
                            if match[0] == "전주":
                                majun = "전주점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 6,majun)
                                a = str(first_empty_row)
                                format_cell_range(sheet, ('F'+a),fmt_전주점)
                                time.sleep(3)
                            elif match[0] == "김해":
                                majun = "김해점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 5,majun)
                                b = str(first_empty_row)
                                format_cell_range(sheet, ('E'+b),fmt_김해점)
                                time.sleep(3)
                            elif match[0] == "창원":
                                majun = "창원점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 4,majun)
                                c = str(first_empty_row)
                                format_cell_range(sheet, ('D'+c),fmt_창원점)
                                time.sleep(3)
                            elif match[0] == "잠실":
                                majun = "잠실점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 3,majun)
                                d = str(first_empty_row)
                                format_cell_range(sheet, ('C'+d),fmt_잠실점)
                                time.sleep(3)
                    elif ":" in result:
                    #print(result) # result[len(result)-1]에는 옵션이 들어가 있다.
                       bunli = result.split(":") # 날짜:5월 25일 ... 이런식으로 ": "를 ":"로 바꿔줘야한다 !
                       if "구매일" == bunli[0]: # 주문건에서는 구매일을 쓰지 않는다.
                           pass
                       else: 
                           print(bunli)
                           fom = bunli[0]
                           nayoung = bunli[1]
                           print(nayoung)
                           if fom == "날짜":
                               sheet.update_cell(first_empty_row,8,nayoung)
                               f = str(first_empty_row)
                               format_cell_range(sheet,('H'+f),fmt_가운데_맞춤)
                               time.sleep(3)
                           elif fom == "성함":
                               sheet.update_cell(first_empty_row,9,nayoung)
                               f = str(first_empty_row)
                               format_cell_range(sheet,("I"+f),fmt_가운데_맞춤)
                               time.sleep(3)
                           elif fom == "연락처":
                               sheet.update_cell(first_empty_row,10,nayoung)
                               f = str(first_empty_row)
                               format_cell_range(sheet,("J"+f),fmt_가운데_맞춤)
                               time.sleep(3)
                           elif fom == "수령방법":
                               sheet.update_cell(first_empty_row,11,nayoung)
                               f = str(first_empty_row)
                               format_cell_range(sheet,("K"+f),fmt_가운데_맞춤)
                               time.sleep(3)
                           elif fom == "상품명":
                               sheet.update_cell(first_empty_row,12,nayoung)
                               f = str(first_empty_row)
                               format_cell_range(sheet,("L"+f),fmt_가운데_맞춤)
                               time.sleep(3)
                           
                # 사진 넣는 코드(폴더에서 사진 하나씩)
            if index_common <= (len(items) - 1):
                item = items[index_common]
                image_url = f"https://drive.google.com/uc?id={item['id']}"
                image_formula = f'=IMAGE("{image_url}")'
                sheet.update_cell(first_empty_row, 14, image_formula)
                time.sleep(3)
                index_common += 1
            

                    
        # as건
        elif "as건" in i_lower:
            sheet = spreadsheet.get_worksheet(1)
            data = sheet.get_all_values('C:G')

            # 비어 있는 첫 번째 행 번호 찾기
            def find_first_empty_row(data):
                for i, row in enumerate(data):
                    # 비어 있는 행을 찾으면 해당 행 번호 반환 (1-based index)
                    if all(cell == '' for cell in row):
                        return i + 1  # 인덱스는 0부터 시작하므로 +1
                return len(data) + 1  # 모든 행이 차 있다면 다음 행 번호 반환           
            
            first_empty_row = find_first_empty_row(data)
            time.sleep(3)
            print("--as건--")
            i_lower = i_lower.split('. ')
            option = i_lower[len(i_lower)-1] #옵션
            print(i_lower)
            print(option)
            lst = option.split(":")
            print(lst)
            real_option = lst[len(lst)-1]# 진짜 옵션 내용
            print(real_option)
            sheet.update_cell(first_empty_row,14,real_option)
            f = str(first_empty_row)
            format_cell_range(sheet, ("N"+f), fmt_가운데_맞춤)
            time.sleep(3)
            for j in i_lower:
                if j[len(j)-1] in ["1", "2", "3", "4", "5", "6", "7", "8"]:
                    result = j[:len(j)-2]
                    if ":" not in result:
                        asjun = result # as
                        a = re.findall(r'\(([^)]*)\)', asjun)
                        asjun = a[0]
                        match = re.findall(r'(전주|김해|창원|잠실)',asjun)
                        try:
                            match[0] == 1
                        
                        except IndexError:
                            b = asjun.split(" ")
                            majun = ""
                            try:
                                b[1]
                            except:
                                majun = b[0]
                                
                                ### 쿄타님에게 추가하는거 요청하기
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 7,majun)
                                e = str(first_empty_row)
                                format_cell_range(sheet, ('G'+e),fmt_기타)
                                time.sleep(3)
                                ### 여기까지         
                                                  
                            else:
                                for i in range(len(b)):
                                    majun += b[i]
                            print("매점은", majun)
                            sheet.update_cell(first_empty_row, 7,majun)
                            e = str(first_empty_row)
                            format_cell_range(sheet, ('G'+e),fmt_기타)
                            time.sleep(3)
                        else:      
                            if match[0] == "전주":
                                majun = "전주점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 6,majun)
                                a = str(first_empty_row)
                                format_cell_range(sheet, ('F'+a),fmt_전주점)
                                time.sleep(3)
                            elif match[0] == "김해":
                                majun = "김해점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 5,majun) 
                                b = str(first_empty_row)
                                format_cell_range(sheet, ('E'+b),fmt_김해점)
                                time.sleep(3)
                            elif match[0] == "창원":
                                majun = "창원점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 4,majun)
                                c = str(first_empty_row)
                                format_cell_range(sheet, ('D'+c),fmt_창원점)
                                time.sleep(3)
                            elif match[0] == "잠실":
                                majun = "잠실점"
                                print("매점은", majun)
                                sheet.update_cell(first_empty_row, 3,majun)
                                d = str(first_empty_row)
                                format_cell_range(sheet, ('C'+d),fmt_잠실점)
                                time.sleep(3)
                    elif ":" in result:
                       bunli = result.split(":") 
                       print(bunli)
                       fom = bunli[0]
                       nayoung = bunli[1]
                       print(fom, nayoung)
                       if fom == "날짜":
                           sheet.update_cell(first_empty_row,8,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('H'+f),fmt_가운데_맞춤)
                           time.sleep(3)
                       elif fom == "성함":
                           sheet.update_cell(first_empty_row,10,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('J'+f),fmt_가운데_맞춤)
                           time.sleep(3)
                       elif fom == "연락처":
                           sheet.update_cell(first_empty_row,11,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('K'+f),fmt_가운데_맞춤)
                           time.sleep(3)
                       elif fom == "수령방법":
                           sheet.update_cell(first_empty_row,12,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('L'+f),fmt_가운데_맞춤)
                           time.sleep(3)
                       elif fom == "상품명":
                           sheet.update_cell(first_empty_row,13,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('M'+f),fmt_가운데_맞춤)
                           time.sleep(3)
                       elif fom == "구매일":
                           sheet.update_cell(first_empty_row,18,nayoung)
                           f = str(first_empty_row)
                           format_cell_range(sheet,('R'+f),fmt_가운데_맞춤)
                           time.sleep(3)

            if index_common <= (len(items) - 1):
                item = items[index_common]
                image_url = f"https://drive.google.com/uc?id={item['id']}"
                image_formula = f'=IMAGE("{image_url}")'
                sheet.update_cell(first_empty_row, 15, image_formula)
                time.sleep(3)
                index_common += 1                    



def move_files(source_folder_id, destination_folder_id):
    # '업로드 이미지' 폴더의 모든 파일 검색
    query = f"'{source_folder_id}' in parents and mimeType contains 'image/'"
    results = drive_service.files().list(q=query, fields="files(id, name, parents)").execute()
    files = results.get('files', [])

    for file in files:
        # 파일을 '유지용 이미지' 폴더로 이동
        file_id = file['id']
        previous_parents = ",".join(file['parents'])
        drive_service.files().update(
            fileId=file_id,
            addParents=destination_folder_id,
            removeParents=previous_parents,
            fields='id, parents'
        ).execute()
        print(f"Moved file: {file['name']}")

# 파일 이동 실행
time.sleep(3)
move_files(folder_id, move_drive_folder_id)