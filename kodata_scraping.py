import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas as pd

# 브라우저 자동 꺼짐 방지 옵션
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# 크롬 드라이버 생성
driver = webdriver.Chrome(options=chrome_options)
# 페이지 로딩 대기
driver.implicitly_wait(5)

# 조회대상명단 불러오기
df = pd.read_excel('C:/Users/heum/Desktop/탐나는인재/프로젝트뱅크/사업자등록번호 리스트.xlsx', engine='openpyxl')
df = df['사업자등록번호'].to_list()

# 중간 저장할 파일 이름
file_name = 'C:/Users/heum/Desktop/탐나는인재/프로젝트뱅크/중소기업현황정보시스템.csv'

# 중간 저장된 데이터 불러오기
try:
    result = pd.read_csv(file_name, encoding='cp949')
    saved_business_numbers = result['사업자등록번호'].tolist()
except FileNotFoundError:
    # 파일이 없을 경우 새로운 데이터프레임 생성
    result = pd.DataFrame(columns=('사업자등록번호', '업체명', '발급번호', '유효기간', '기업규모'))
    saved_business_numbers = []

# 조회 대상 중 이미 저장된 데이터는 건너뛰기
for business_number in df:
    if business_number in saved_business_numbers:
        print(f"Already scraped: {business_number}")
        continue
    
    # 조회 페이지로 이동
    driver.get(url='https://sminfo.mss.go.kr/sc/sy/SSY004R0.do')
    
    # 값 입력 및 조회
    driver.find_element(By.XPATH, '//*[@id="searchTxt"]').send_keys(business_number)
    driver.find_element(By.XPATH, '//*[@id="tab01"]/div/div[2]/button[1]').click()
    
    # "조회된 내용이 없습니다" 메시지 확인
    try:
        no_result_message = driver.find_element(By.XPATH, '//*[contains(text(), "조회된 내용이 없습니다")]')
        if no_result_message:
            print(f"{business_number}: 조회불가")
            result = result.append({
                '사업자등록번호': business_number,
                '업체명': '조회불가',
                '발급번호': '조회불가',
                '유효기간': '조회불가',
                '기업규모': '조회불가'
            }, ignore_index=True)
            # 바로 중간 저장 후 다음으로 넘어가기
            result.to_csv(file_name, encoding='cp949', index=False)
            continue
    except:
        pass  # 조회된 내용이 없는 메시지가 없을 경우 계속 진행
    
    # XPath로 요소 찾기
    try:
        name = driver.find_element(By.XPATH, '//*[@id="tab01"]/div/div[3]/table/tbody/tr/td[1]').text
        issue_num = driver.find_element(By.XPATH, '//*[@id="tab01"]/div/div[3]/table/tbody/tr/td[3]').text
        expiration_date = driver.find_element(By.XPATH, '//*[@id="tab01"]/div/div[3]/table/tbody/tr/td[4]').text
        size_company = driver.find_element(By.XPATH, '//*[@id="tab01"]/div/div[3]/table/tbody/tr/td[5]').text
        
        # 결과 데이터프레임에 추가
        result = result.append({
            '사업자등록번호': business_number,
            '업체명': name,
            '발급번호': issue_num,
            '유효기간': expiration_date,
            '기업규모': size_company
        }, ignore_index=True)

    # 조회 불가 시 예외 처리
    except Exception as e:
        print(f"Error scraping {business_number}: {e}")
        result = result.append({
            '사업자등록번호': business_number,
            '업체명': '조회불가',
            '발급번호': '조회불가',
            '유효기간': '조회불가',
            '기업규모': '조회불가'
        }, ignore_index=True)
    
    # 데이터가 스크래핑될 때마다 중간 저장
    result.to_csv(file_name, encoding='cp949', index=False)
    
    # 잠시 대기 (과부하 방지)
    time.sleep(1)

# 최종 완료 메시지
print("데이터 조회 및 저장 완료.")