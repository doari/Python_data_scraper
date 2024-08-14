# Python_data_scraper
A Python script to scrape product information and images from a women's belt webpage, saving the data into an Excel file with images.

이 파이썬 스크립트는 웹사이트에서 제품의 정보와 이미지를 스크래핑하여 Excel 파일로 저장합니다.

## 주요 기능

1. **웹 페이지 접근**: 웹사이트 사용하여 제품의 정보 페이지에 접속합니다.
2. **스크롤 및 데이터 수집**: 페이지를 끝까지 스크롤하며, 모든 제품 정보를 BeautifulSoup으로 파싱합니다.
3. **이미지 다운로드**: 제품 이미지를 다운로드하여, 상품 번호를 파일명으로 저장합니다.
4. **데이터 저장**: 수집된 정보를 Pandas를 이용해 Excel 파일로 저장하고, Openpyxl을 통해 스타일을 적용합니다.

## 필요한 라이브러리

- Selenium
- BeautifulSoup
- Requests
- Pillow
- Pandas
- Openpyxl

## 실행 방법

1. 위에 명시된 라이브러리를 설치합니다.
2. ChromeDriver 경로를 설정한 후 스크립트를 실행합니다.
3. 결과는 `products.xlsx` 파일과 `images` 폴더에 저장됩니다.
