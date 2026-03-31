import os
import requests
import tempfile
import time
import getpass
import pandas as pd
import subprocess
from bs4 import BeautifulSoup

def download_kisa_ip_xls(output_path: str):
    """
    KISA 대한민국 IP 주소 XLS 파일 다운로드 예시 함수.
    실제 동작하려면 '엑셀 다운로드' 버튼이 연결된 실제 다운로드 URL을 확인/교체해야 합니다.
    """
    # 세션 생성
    session = requests.Session()

    # 1) KISA 메인 페이지 접근
    kisa_main_url = "https://xn--3e0bx5euxnjje69i70af08bea817g.xn--3e0b707e/jsp/statboard/IPAS/inter/sec/currentV4Addr.jsp"
    resp = session.get(kisa_main_url)
    resp.raise_for_status()

    # 2) HTML 파싱하여 "엑셀 다운로드" 버튼 링크 추출 (예시)
    soup = BeautifulSoup(resp.text, "html.parser")
    
    # 실제로 엑셀 다운로드가 가능한 링크 찾아보기 (예시는 가정)
    # 만약 <a href="getFile.jsp?filename=XXXX.xls">엑셀 다운로드</a> 같은 형태라면:
    download_link = "https://한국인터넷정보센터.한국/jsp/statboard/IPAS/inter/sec/ipv4AddrListExcel.jsp"
    for a_tag in soup.find_all("a"):
        if "엑셀 다운로드" in a_tag.get_text():
            download_link = a_tag.get("href")
            break

    if not download_link:
        raise ValueError("KISA 사이트에서 엑셀 다운로드 링크를 찾을 수 없습니다.")
    
    # 필요한 경우 절대 URL로 변환
    # (링크가 상대 경로라면 아래와 같은 방식으로 변환)
    kisa_download_url = (
        download_link if "http" in download_link
        else os.path.join(os.path.dirname(kisa_main_url), download_link)
    )
    
    # 3) 실제 XLS 파일 다운로드
    download_resp = session.get(kisa_download_url)
    download_resp.raise_for_status()
    
    with open(output_path, "wb") as f:
        f.write(download_resp.content)
    
    print(f"KISA XLS 파일이 다운로드 되었습니다: {output_path}")


def download_maxmind_csv_zip(output_path: str):
    """
    MaxMind 사이트에 로그인하여 GeoLite2 Country CSV ZIP 파일을 다운로드하는 예시 함수.
    실제 동작하려면 MaxMind 로그인 로직과 2FA를 위한 사용자 입력 처리 등을 제대로 구현해야 합니다.
    """
    # MaxMind 로그인 및 다운로드에 사용할 세션
    session = requests.Session()

    # 1) 사용자 자격증명 입력 받기 (예시)
    username = input("MaxMind 계정 Username(Email)을 입력하세요: ")
    password = getpass.getpass("MaxMind 계정 Password를 입력하세요: ")

    # 2) 로그인 페이지 접근 (예시 URL)
    login_page_url = "https://www.maxmind.com/en/account/sign-in"
    resp = session.get(login_page_url)
    resp.raise_for_status()

    # 3) 로그인 폼 전송 (예시)
    #    실제로는 로그인 페이지 소스에서 form의 name, action, csrf_token 등을 확인해야 함
    soup = BeautifulSoup(resp.text, "html.parser")
    csrf_token_tag = soup.find("input", attrs={"name": "authenticity_token"})
    if not csrf_token_tag:
        raise ValueError("CSRF 토큰을 찾을 수 없습니다. 페이지 구조가 변경된 듯 합니다.")
    csrf_token = csrf_token_tag.get("value")

    login_data = {
        "user_email": username,
        "user_password": password,
        "authenticity_token": csrf_token
        # 필요 시 추가 파라미터
    }
    # 실제 로그인 action URL 확인 후 사용
    login_action_url = "https://www.maxmind.com/en/login"  
    login_resp = session.post(login_action_url, data=login_data)
    login_resp.raise_for_status()

    # 4) 2단계 인증(OTP, 이메일 코드 등) 처리 (예시)
    #    실제로는 OTP 코드 입력창이 있을 것이며,
    #    그 페이지의 form name, csrf_token 등을 다시 추출하여 전송
    #    여기서는 단순히 사용자 입력만 받고, 그대로 전송하는 예시를 가정
    #    2FA 페이지 URL도 실제 페이지 구조 따라 변경해야 함
    if "2-Step Verification" in login_resp.text or "Please enter your 2FA code" in login_resp.text:
        two_fa_code = input("2단계 인증 코드를 입력하세요(OTP/이메일 전송 코등): ")
        two_fa_soup = BeautifulSoup(login_resp.text, "html.parser")
        two_fa_token_tag = two_fa_soup.find("input", attrs={"name": "authenticity_token"})
        if not two_fa_token_tag:
            raise ValueError("2FA 페이지의 CSRF 토큰을 찾을 수 없습니다.")
        two_fa_token = two_fa_token_tag.get("value")
        
        two_fa_data = {
            "token": two_fa_code,
            "authenticity_token": two_fa_token
            # 기타 필요한 파라미터
        }
        two_fa_action_url = "https://www.maxmind.com/en/2fa/verify"  # 예시
        verify_resp = session.post(two_fa_action_url, data=two_fa_data)
        verify_resp.raise_for_status()
        if "Invalid two-factor code" in verify_resp.text:
            raise ValueError("2단계 인증 코드가 잘못되었습니다.")

    # 5) 다운로드 페이지 접근
    #    로그인이 정상 완료됐다면, 아래 URL에서 파일 다운로드 가능
    download_page_url = "https://www.maxmind.com/en/accounts/XXXXXX/geoip/downloads"  # 실제 계정 ID 부분 수정
    page_resp = session.get(download_page_url)
    page_resp.raise_for_status()
    page_soup = BeautifulSoup(page_resp.text, "html.parser")

    # 6) "GeoLite2 Country: CSV Format" ZIP 파일 링크 파싱
    #    실제 페이지에서 "GeoLite2-Country-CSV.zip" 등 원하는 요소 찾기
    csv_link = None
    for link_tag in page_soup.find_all("a"):
        href = link_tag.get("href")
        if href and "GeoLite2-Country-CSV.zip" in href:
            csv_link = href
            break

    if not csv_link:
        raise ValueError("GeoLite2 Country CSV ZIP 파일 링크를 찾을 수 없습니다.")

    # 링크가 절대경로인지, 상대경로인지에 따라 처리
    if not csv_link.startswith("http"):
        csv_link = os.path.join(os.path.dirname(download_page_url), csv_link)

    # 7) ZIP 파일 다운로드
    print("MaxMind GeoLite2 Country CSV ZIP 파일 다운로드 중...")
    zip_resp = session.get(csv_link)
    zip_resp.raise_for_status()

    with open(output_path, "wb") as f:
        f.write(zip_resp.content)
    
    print(f"MaxMind CSV ZIP 파일이 다운로드 되었습니다: {output_path}")


def remove_first_three_rows_from_xls(xls_path: str, output_path: str):
    """
    KISA에서 받은 XLS 파일의 상단 3행(title 영역 등)을 제거한 뒤,
    새로운 XLS 파일로 저장하는 예시 함수.
    """
    # pandas로 파일을 읽을 때 skiprows 파라미터 사용 가능 (행 개수만큼 건너뛰기)
    # 단, KISA XLS 구조에 따라 sheet_name, header 등 옵션을 맞춰야 함
    # 여기서는 간단히 첫 번째 시트만 읽는다고 가정
    df = pd.read_excel(xls_path, sheet_name=0, header=None, skiprows=3)

    # 필요한 경우 컬럼명을 다시 설정하거나, 가공 로직 추가
    # 예: df.columns = ["Col1","Col2", ...] 등

    # 처리된 내용을 새로운 XLS로 저장 (xlsx 또는 csv 등)
    # XLS 포맷 그대로 저장하려면 별도의 엔진(xlwt, openpyxl 등)이 필요
    # 여기서는 xlsx로 저장 예시
    df.to_excel(output_path, index=False, header=False)
    print(f"상위 3행이 제거된 XLS 파일 저장 완료: {output_path}")


def run_merge_jar_with_bat(processed_kisa_xls_path: str, maxmind_csv_path: str, result_csv_path: str):
    """
    .bat 스크립트를 통해 .jar 파일을 실행하여 두 파일을 합치는 프로세스.
    .bat 스크립트는 아래처럼 가정 (stdin으로 3개 인자 받음).
    """
    # Windows 환경에서 stdin으로 인자를 전달해야 한다면, subprocess.Popen으로 표준입력 연결
    # 여기서는 간단히 echo 등을 통해 입력을 전달하는 예시
    # 예) test_merge.bat 파일이 stdin에서 3줄을 읽어온다고 가정
    # (각 줄은 각각의 인자: kisa_xls_path / csv_path / output_csv_path)
    
    bat_file_path = r"C:\path\to\merge_run.bat"  # 실제 배치 파일 경로
    # .bat 파일에서 내부적으로 "java -jar merge.jar" 이런 식으로 처리한다고 가정
    
    input_lines = f"{processed_kisa_xls_path}\n{maxmind_csv_path}\n{result_csv_path}\n"
    
    print("배치 스크립트를 실행합니다...")

    process = subprocess.Popen(
        [bat_file_path],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    stdout, stderr = process.communicate(input=input_lines)

    if process.returncode != 0:
        print("배치 스크립트 실행 중 오류가 발생했습니다.")
        print("stderr:", stderr)
    else:
        print("배치 스크립트 실행 완료!")
        print("stdout:", stdout)


def main():
    # 0) 임시 폴더 등 지정(원하는 폴더를 사용)
    work_dir = os.getcwd()
    
    kisa_xls_raw_path = os.path.join(work_dir, "KISA_raw.xls")
    kisa_xls_processed_path = os.path.join(work_dir, "KISA_processed.xlsx")
    maxmind_zip_path = os.path.join(work_dir, "GeoLite2-Country-CSV.zip")

    # 1) KISA 대한민국 IP XLS 다운로드
    download_kisa_ip_xls(kisa_xls_raw_path)
    
    # 2) MaxMind CSV(zip) 다운로드
    download_maxmind_csv_zip(maxmind_zip_path)
    
    # 3) KISA XLS 상단 3행 제거 → 새로운 XLS (또는 XLSX)로 저장
    remove_first_three_rows_from_xls(kisa_xls_raw_path, kisa_xls_processed_path)

    # 4) JAR 실행(배치 스크립트)으로 두 파일 합치기
    #    실제 CSV 파일은 maxmind_zip_path 압축을 풀어서
    #    "GeoLite2-Country-Blocks-IPv4.csv"와 "GeoLite2-Country-Locations.csv" 등을
    #    합쳐 쓰거나, 작업 상황에 맞춰 사용해야 합니다.
    #    여기서는 간단히 "압축 해제 후 나온 csv" 라고 가정하여 경로 예시만 표기

    # 예: 압축 해제 후 "GeoLite2-Country-Blocks-IPv4.csv"가 있다고 가정
    maxmind_csv_extracted_path = os.path.join(work_dir, "GeoLite2-Country-Blocks-IPv4.csv")

    # 결과물 csv 저장 경로
    result_csv_path = os.path.join(work_dir, "merged_result.csv")

    run_merge_jar_with_bat(
        processed_kisa_xls_path=kisa_xls_processed_path,
        maxmind_csv_path=maxmind_csv_extracted_path,
        result_csv_path=result_csv_path
    )

if __name__ == "__main__":
    main()
