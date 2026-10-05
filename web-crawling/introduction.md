2026.10.05

멋쟁이사자처럼 온라인 강의

일은 간편해지고 수익은 성장하는 자동화 스킬 : 파이썬부터 Chat GPT 까지

---

- 웹(web): 인터넷 공간
- 크롤링(Crawling): 기어가는 것
- 웹크롤링: 웹페이지에서 데이터를 추출하는 행위

  ex. 구글, 네이버 스토어, 멜론 차트 수집
  
  예시 코드: 15-웹크롤링-멜론차트수집.py 
  -> 엑셀파일 생성됨

```python
import urllib.request as req

from bs4 import BeautifulSoup
import os
import openpyxl
import datetime
from openpyxl.drawing.image import Image

# 기존 엑셀파일 삭제
if os.path.exists("./멜론_크롤링.xlsx"):
    os.remove("./멜론_크롤링.xlsx")

# 이미지 저장할 폴더 생성
if not os.path.exists("./멜론이미지"):
    os.mkdir("./멜론이미지")

header = req.Request("https://www.melon.com/chart/index.htm", headers={"User-Agent":"Mozilla/5.0"})
code = req.urlopen(header)
soup = BeautifulSoup(code, "html.parser")
title = soup.select("div.ellipsis.rank01 > span > a")
name = soup.select("div.ellipsis.rank02 > span")
album = soup.select("div.ellipsis.rank03 > a")
img = soup.select("a.image_typeAll > img")

# 엑셀 파일 생성
if not os.path.exists("./멜론_크롤링.xlsx"):
    openpyxl.Workbook().save("./멜론_크롤링.xlsx")

# 엑셀 파일 불러오기
book = openpyxl.load_workbook("./멜론_크롤링.xlsx")
# 쓸데 없는 시트 지우기
if "Sheet" in book.sheetnames:
    book.remove(book["Sheet"])
sheet = book.create_sheet()
now = datetime.datetime.now()
sheet.title = f"{now.year}년 {now.month}월 {now.day}일 {now.hour}시 {now.minute}분 {now.second}초"
# 열 너비 조절
sheet.column_dimensions["A"].width = 15
sheet.column_dimensions["B"].width = 50
sheet.column_dimensions["C"].width = 30
sheet.column_dimensions["D"].width = 50

for i in range(len(title)):
    img_file_name = f"./멜론이미지/{i+1}.png"
    req.urlretrieve(img[i].attrs["src"], img_file_name)
    print(f"{i+1}위. {title[i].text} - {name[i].text}")
    img_for_excel = Image(img_file_name)
    sheet.add_image(img_for_excel, f"A{i+1}")
    sheet.cell(row=i + 1, column=2).value = title[i].text
    sheet.cell(row=i + 1, column=3).value = name[i].text
    sheet.cell(row=i + 1, column=4).value = album[i].text
    sheet.row_dimensions[i+1].height = 90
    book.save("./멜론_크롤링.xlsx")

```

---

## 원리

주소창에 주소 치면, 

**클라이언트(내 컴퓨터) --> 서버**

우리 컴퓨터에 연결된 인터넷 통해서,

*"서버야 너네 웹페이지 좀 줘!(요청)"* (파이썬에서도 가능)

**클라이언트 <-- 서버**

"HTML 코드 보낼게~"

**브라우저(크롬)가 HTML 시각화 해줌**

내 컴퓨터에 우리가 보는 웹 사이트 보임

---

  ## 실습 CGV 무비차트
  1. cgv 무비차트 사이트 들어가기
  2. 마우스 우클릭->검사: 개발자도구 창 열림 (HTML)
  3. 코드에서 웹사이트 링크 가져오면 HTML 코드 가져올 수 있음
  4. HTML 코드에서 원하는 정보 뽑는 노가다 시킬건데, 효율을 위해서 *HTML 코드 예쁘게 정리하기* BeautifulSoup(code.text, features:"html.parser")
  5. 원하는 요소 알려주




