import re
from konlpy.tag import Okt
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import * # 셀 꾸미는 모듈
from googleapiclient.discovery import build # 폴더 이미지를 끌어오는 코드

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
spreadsheet_id = '1FJ-NYU-Hf9Ig4KgmKpPMy50TKYpwPSFtBoG6SIh8PH0'
spreadsheet = client.open_by_key(spreadsheet_id)
drive_service = build('drive', 'v3', credentials=credentials)

data_1 = "전주\n김해\n잠실\n창원\n롯대" # '점'을 붙이면 안된다.. 그냥 롯대잠실, 신세계김해 이런식으로
okt = Okt()

index_common = 0
index_jumun = 0
index_asjun = 0
count_jumun_1 = 0
count_jumun_2 = 0
count_asjun_1 = 0
count_asjun_2 = 0

folder_id = '1CSjt_1nRMZrguO_06RVs4zeBQRt4l8yp'
query = f"'{folder_id}' in parents and (mimeType='image/jpeg' or mimeType='image/png')"
results = drive_service.files().list(q=query, fields="files(id, name)").execute()
items = results.get('files', [])
item = items[index_common]


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

with open("C:/Users/djdjd/OneDrive/바탕 화면/Python/python_kota/study.txt", "r", encoding="UTF-8") as f:
    text = f.read().replace("\n", " ")
    result = re.findall(r'\{([^}]*)\}', text)
    print(result)
    for i in result:
        i_lower = i.lower()
        # 주문건
        if "주문건" in i_lower:
            sheet = spreadsheet.get_worksheet(0)
            count_jumun_2 += 1
            count_jumun_1 = 0
            print("--주문건--")
            i_lower = i_lower.split('. ')
            option = i_lower[len(i_lower)-1] #옵션
            lst = option.split(":")
            real_option = lst[len(lst)-1]# 진짜 옵션 내용
            sheet.update_cell(105+count_jumun_2,13,real_option)
            for j in i_lower:
                if j[len(j)-1] in ["1", "2", "3", "4", "5", "6", "7", "8"]:
                    result = j[:len(j)-2]
                    if ":" not in result:
                        jumun = result # 주문
                        print(jumun)
                        pos_tag = okt.pos(jumun)
                        filt_pos= [word for word, pos in pos_tag if pos in ["Noun"]]
                        if "전주" in filt_pos:
                            majun = "전주점"
                            print("매점은", majun)
                            sheet.update_cell(105+count_jumun_2, 6,majun) # 105는 내가 수동으로 바꿔야하는 행이다.
                            a = str(105+count_jumun_2)
                            format_cell_range(sheet, ('F'+a),fmt_전주점)
                        elif "김해" in filt_pos:
                            majun = "김해점"
                            print("매점은", majun)
                            sheet.update_cell(105+count_jumun_2, 5,majun)
                            b = str(105+count_jumun_2)
                            format_cell_range(sheet, ('E'+b),fmt_김해점)
                        elif "창원" in filt_pos:
                            majun = "창원점"
                            print("매점은", majun)
                            sheet.update_cell(105+count_jumun_2, 4,majun)
                            c = str(105+count_jumun_2)
                            format_cell_range(sheet, ('D'+c),fmt_창원점)
                        elif "잠실" in filt_pos:
                            majun = "잠실점"
                            print("매점은", majun)
                            sheet.update_cell(105+count_jumun_2, 3,majun)
                            d = str(105+count_jumun_2)
                            format_cell_range(sheet, ('C'+d),fmt_잠실점)
                        else :
                            a = re.findall(r'\(([^)]*)\)', jumun)
                            b = a[0].split(" ")
                            majun = ""
                            try:
                                b[1]
                            except:
                                majun = b[0]
                                
                            else:
                                for i in range(len(b)):
                                    majun += b[i]
                                print("매점은", majun)
                                sheet.update_cell(105+count_jumun_2, 7,majun)
                                e = str(105+count_jumun_2)
                                format_cell_range(sheet, ('G'+e),fmt_기타)
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
                           sheet.update_cell(105+count_jumun_2,8+count_jumun_1,nayoung)
                           count_jumun_1 += 1
                # 사진 넣는 코드(폴더에서 사진 하나씩)
            if index_common <= (len(items) - 1):
                item = items[index_common]
                image_url = f"https://drive.google.com/uc?id={item['id']}"
                image_formula = f'=IMAGE("{image_url}")'
                sheet.update_cell(106+ index_jumun, 14, image_formula) # 106번행 부터
                index_common += 1
                index_jumun += 1
            

                    
        # as건
        elif "as건" in i_lower:
            sheet = spreadsheet.get_worksheet(1)
            count_asjun_2 += 1
            count_asjun_1 = 0          
            print("--as건--")
            i_lower = i_lower.split('. ')
            option = i_lower[len(i_lower)-1] #옵션
            lst = option.split(":")
            real_option = lst[len(lst)-1]# 진짜 옵션 내용
            sheet.update_cell(95+count_jumun_2,13,real_option)
            for j in i_lower:
                if j[len(j)-1] in ["1", "2", "3", "4", "5", "6", "7", "8"]:
                    result = j[:len(j)-2]
                    if ":" not in result:
                        asjun = result # as
                        pos_tag = okt.pos(asjun)
                        filt_pos= [word for word, pos in pos_tag if pos in ["Noun"]]
                        if "전주" in filt_pos:
                            majun = "전주점"
                            print("매점은", majun)
                            sheet.update_cell(95+count_asjun_2, 6,majun) # 95는 내가 수동으로 바꿔야하는 행이다.
                            a = str(95+count_asjun_2)
                            format_cell_range(sheet, ('F'+a),fmt_전주점)
                        elif "김해" in filt_pos:
                            majun = "김해점"
                            print("매점은", majun)
                            sheet.update_cell(95+count_asjun_2, 5,majun) 
                            b = str(95+count_asjun_2)
                            format_cell_range(sheet, ('E'+b),fmt_김해점)
                        elif "창원" in filt_pos:
                            majun = "창원점"
                            print("매점은", majun)
                            sheet.update_cell(95+count_asjun_2, 4,majun)
                            c = str(95+count_asjun_2)
                            format_cell_range(sheet, ('D'+c),fmt_창원점)
                        elif "잠실" in filt_pos:
                            majun = "잠실점"
                            print("매점은", majun)
                            sheet.update_cell(95+count_asjun_2, 3,majun)
                            d = str(95+count_asjun_2)
                            format_cell_range(sheet, ('C'+d),fmt_잠실점)
                        else :
                            a = re.findall(r'\(([^)]*)\)', asjun)
                            print(asjun)
                            b = a[0].split(" ")
                            majun = ""
                            try:
                                b[1]
                            except:
                                majun = b[0]
                                print(majun)
                            else:
                                for i in range(len(b)):
                                    majun += b[i]
                                    print(majun)
                            print("매점은", majun)
                            sheet.update_cell(95+count_jumun_2, 7,majun)
                            e = str(95+count_jumun_2)
                            format_cell_range(sheet, ('G'+e),fmt_기타)
                    elif ":" in result:
                       bunli = result.split(":") 
                       print(bunli)
                       fom = bunli[0]
                       nayoung = bunli[1]
                       print(nayoung)
                       sheet.update_cell(95+count_asjun_2,8+count_asjun_1,nayoung)
                       count_asjun_1 += 1

            if index_common <= (len(items) - 1):
                item = items[index_common]
                image_url = f"https://drive.google.com/uc?id={item['id']}"
                image_formula = f'=IMAGE("{image_url}")'
                sheet.update_cell(96+ index_asjun, 14, image_formula) # 106번행 부터
                index_common += 1
                index_asjun += 1                    


'''
items = results.get('files', [])
print(items, index_asjun, index_jumun,index_common)
'''