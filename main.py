import requests
import xml.etree.ElementTree as ET
import pandas as pd
import google.generativeai as genai
import time
import os
import json
import smtplib
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 환경 변수 (GitHub Secrets)
# ==========================================
LAW_API_KEY = os.environ.get("LAW_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")      
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")  
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")

# ==========================================
# 2. 날짜 및 AI 세팅 (제미나이 2.5 원상복구!)
# ==========================================
KST = timezone(timedelta(hours=9))
today = datetime.now(KST)
TARGET_DATE = today.strftime("%Y%m%d")
FILE_PREFIX = today.strftime("%Y년_%m월_%d일")

genai.configure(api_key=GEMINI_API_KEY)
# 선생님의 최애 버전으로 고정!
model = genai.GenerativeModel('gemini-2.5-flash') 

HEADERS = {'User-Agent': 'Mozilla/5.0'}
session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

# ==========================================
# 3. 공단 전용 491개 자격 종목 파일(CSV) 읽어오기
# ==========================================
try:
    # 한글 깨짐 방지를 위해 인코딩 2중 체크
    try:
        df_certs = pd.read_csv('certs.csv', encoding='utf-8')
    except UnicodeDecodeError:
        df_certs = pd.read_csv('certs.csv', encoding='cp949')
        
    cert_list = df_certs['종목명'].dropna().tolist()
    QNET_CERTS = ", ".join(cert_list)
    print(f"✅ 공단 전용 국가기술자격 {len(cert_list)}개 종목 목록 로드 완료!")
except Exception as e:
    print(f"❌ [에러] certs.csv 파일을 찾을 수 없거나 읽을 수 없습니다: {e}")
    QNET_CERTS = "국가기술자격증" # 파일이 없을 경우의 기본값 방어

# ==========================================
# 4. 핵심 실행 함수들
# ==========================================
def get_todays_laws(api_key, target_date):
    all_laws_dict = {}
    search_date_range = f"{target_date}~{target_date}"
    print(f"\n📅 [{target_date}] 법제처 데이터를 수집합니다...")
    
    for target_type in ['law', 'histlaw']:
        page = 1
        while True:
            search_url = f"https://www.law.go.kr/DRF/lawSearch.do?OC={api_key}&target={target_type}&type=XML&efYd={search_date_range}&display=100&page={page}"
            try:
                response = session.get(search_url, headers=HEADERS, timeout=15)
                root = ET.fromstring(response.text)
                law_nodes = root.findall('.//law')
                if not law_nodes: break
                
                for law in law_nodes:
                    law_id = law.find('법령일련번호').text if law.find('법령일련번호') is not None else ""
                    law_name = law.find('법령명한글').text if law.find('법령명한글') is not None else "이름없음"
                    enforce_date = law.find('시행일자').text if law.find('시행일자') is not None else ""
                    
                    if not law_id or law_name in all_laws_dict: continue
                    
                    detail_url = f"https://www.law.go.kr/DRF/lawService.do?OC={api_key}&target={target_type}&MST={law_id}&type=XML"
                    detail_response = session.get(detail_url, headers=HEADERS, timeout=15)
                    detail_root = ET.fromstring(detail_response.text)
                    
                    reason_text = ""
                    for tag in ['.//개정이유', './/제개정이유']:
                        r_node = detail_root.find(tag)
                        if r_node is not None and r_node.text: reason_text += r_node.text.strip() + "\n"
                    
                    body_text = "\n".join([j.text.strip() for j in detail_root.findall('.//조문내용') if j.text])
                    stars = "\n".join([s.text.strip() for s in detail_root.findall('.//별표내용') if s.text])
                    full_text = f"[개정이유]\n{reason_text}\n[조문내용]\n{body_text}\n[별표]\n{stars}"[:20000]
                    
                    all_laws_dict[law_name] = {"법령명": law_name, "시행일자": enforce_date, "원본": full_text}
                    print(f"  📥 수집 중: {law_name}")
                    time.sleep(0.2)
                
                if len(law_nodes) < 100: break
                page += 1
            except: break
    return list(all_laws_dict.values())

def apply_excel_formatting(filename, df_summary, df_detail):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='요약표', index=False)
        df_detail.to_excel(writer, sheet_name='상세분석', index=False)
    
    wb = load_workbook(filename)
    ws = wb['상세분석']
    # 엑셀 헤더: 활용도 분석으로 완벽 변경
    new_headers = ["시행일자", "법령명", "주요 제·개정내용", "법령 관련 국가기술자격 종목", "활용도 분석 구분", "활용도 분석 상세"]
    for i, h in enumerate(new_headers, 1):
        ws.cell(row=1, column=i).value = h
        ws.cell(row=1, column=i).font = Font(bold=True)
        ws.cell(row=1, column=i).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws.column_dimensions[col].width = 45
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')
    wb.save(filename)

def send_naver_email(filename, total, important):
    print("\n📧 네이버 메일 발송 준비...")
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = f"🤖 [일일 보고] {FILE_PREFIX} 국가기술자격 관계 법령 분석"
    
    body = f"오늘 시행법령 총 {total}건 중 자격 활용도 유의미 {important}건 분석 완료."
    msg.attach(MIMEText(body, 'plain'))
    
    with open(filename, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filename)}")
        msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.naver.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL.split('@')[0], EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("✅ 네이버 메일 발송 성공!")
    except Exception as e:
        print(f"❌ 발송 실패: {e}")

def main():
    laws = get_todays_laws(LAW_API_KEY, TARGET_DATE)
    if not laws: return
    
    important_laws = []
    print(f"\n🏎️ {len(laws)}건 정밀 분석(V11) 시작...")
    
    for idx, law in enumerate(laws):
        print(f"[{idx+1}/{len(laws)}] {law['법령명']}... ", end="")
        
        prompt = f"""
        당신은 한국산업인력공단의 국가기술자격 규제 심사 수석 연구원입니다. 
        매우 보수적인 잣대로 '활용도 분석'을 수행하십시오.

        [491개 자격 사전] {QNET_CERTS}
        [법령명] {law['법령명']}
        [내용] {law['원본']}

        [판정 가이드라인]
        1. 분류: 자격증 소지자의 의무 선임, 가점, 채용 요건에 직접적 변화가 있다면 '중요', 아니면 '일반'
        2. 활용도_구분: [대폭 증가, 소폭 증가, 변동 없음, 소폭 감소, 대폭 감소] 중 반드시 하나 선택
        3. 직접적 명시가 없는 추론은 모두 '변동 없음' 처리할 것.

        [출력 JSON]
        {{
            "분류": "중요 또는 일반",
            "요약": "법령 핵심 요약",
            "종목": "매칭된 자격증 명칭",
            "활용도_구분": "5단계 중 선택",
            "활용도_분석": "판단 근거"
        }}
        """
        try:
            response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            raw_text = response.text.strip()
            
            data = json.loads(raw_text)
            if isinstance(data, list): data = data[0]
            
            if data.get("분류") == "중요" and data.get("활용도_구분") != "변동 없음":
                important_laws.append({
                    "시행일자": law["시행일자"],
                    "법령명": law["법령명"],
                    "요약": data.get("요약"),
                    "관련종목": data.get("종목"),
                    "구분": data.get("활용도_구분"),
                    "심층분석": data.get("활용도_분석")
                })
                print("👉 [채택]")
            else: print("❌ [패스]")
        except: print("⚠️ [에러]")
        
    if important_laws:
        df_detail = pd.DataFrame(important_laws)
        df_summary = pd.DataFrame({"구분": ["총 시행법령", "유의미 법령"], "건수": [len(laws), len(important_laws)]})
        fname = f"HRD_Daily_Report_{TARGET_DATE}.xlsx"
        apply_excel_formatting(fname, df_summary, df_detail)
        send_naver_email(fname, len(laws), len(important_laws))

if __name__ == "__main__":
    main()
