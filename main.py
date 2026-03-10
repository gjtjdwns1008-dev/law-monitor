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
# 1. 환경 변수 세팅
# ==========================================
LAW_API_KEY = os.environ.get("LAW_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")

# ==========================================
# 2. 한국 시간(KST) 세팅
# ==========================================
KST = timezone(timedelta(hours=9))
today = datetime.now(KST)
TARGET_DATE = today.strftime("%Y%m%d")
FILE_PREFIX = today.strftime("%Y년_%m월_%d일")

genai.configure(api_key=GEMINI_API_KEY)

model = genai.GenerativeModel('gemini-2.5-flash') 

HEADERS = {'User-Agent': 'Mozilla/5.0'}
CURRENT_FOLDER = os.path.dirname(os.path.abspath(__file__))

session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

# ==========================================
# 3. 최신 국가기술자격 491개 마스터 도감
# ==========================================
QNET_CERTS = """
[사업관리] 공공조달관리사
[경영·회계·사무] 공장관리기술사, 포장기술사, 품질관리기술사, 포장기사, 품질경영기사, 포장산업기사, 품질경영산업기사, 사회조사분석사1급, 사회조사분석사2급, 소비자전문상담사1급, 소비자전문상담사2급, 컨벤션기획사1급, 컨벤션기획사2급
[교육·자연과학·사회과학] 이러닝운영관리사
[보건·의료] 국제의료관광코디네이터, 임상심리사1급, 임상심리사2급
[사회복지·종교] 직업상담사1급, 직업상담사2급
[문화·예술·디자인·방송] 제품디자인기술사, 시각디자인기사, 제품디자인기사, 컬러리스트기사, 시각디자인산업기사, 제품디자인산업기사, 컬러리스트산업기사, 웹디자인개발기능사, 제품응용모델링기능사, 컴퓨터그래픽기능사
[운전·운송] 철도운송산업기사, 농기계운전기능사
[영업·판매] 텔레마케팅관리사
[이용·숙박·여행·오락·스포츠] 미용장, 이용장, 스포츠경영관리사, 미용사(네일), 미용사(메이크업), 미용사(일반), 미용사(피부), 이용사
[음식서비스] 조리기능장, 복어조리산업기사, 양식조리산업기사, 일식조리산업기사, 중식조리산업기사, 한식조리산업기사, 복어조리기능사, 양식조리기능사, 일식조리기능사, 조주기능사, 중식조리기능사, 한식조리기능사
[건설] 금속재창호기능사, 플라스틱창호기능사, 건축구조기술사, 건축기계설비기술사, 건축시공기술사, 건축품질시험기술사, 교통기술사, 농어업토목기술사, 도로및공항기술사, 도시계획기술사, 상하수도기술사, 수자원개발기술사, 조경기술사, 지적기술사, 지질및지반기술사, 철도기술사, 측량및지형공간정보기술사, 토목구조기술사, 토목시공기술사, 토목품질시험기술사, 토질및기초기술사, 항만및해안기술사, 해양기술사, 건축목재시공기능장, 건축일반시공기능장, 배관기능장, 잠수기능장, 건설재료시험기사, 건축기사, 건축설비기사, 교통기사, 도시계획기사, 실내건축기사, 응용지질기사, 조경기사, 지적기사, 철도토목기사, 측량및지형공간정보기사, 콘크리트기사, 토목기사, 항로표지기사, 해양공학기사, 해양자원개발기사, 해양환경기사, 건설재료시험산업기사, 건축목공산업기사, 건축산업기사, 건축설비산업기사, 건축일반시공산업기사, 공간정보융합산업기사, 교통산업기사, 방수산업기사, 배관산업기사, 실내건축산업기사, 잠수산업기사, 조경산업기사, 지적산업기사, 측량및지형공간정보산업기사, 콘크리트산업기사, 토목산업기사, 항로표지산업기사, 해양조사산업기사, 거푸집기능사, 건설재료시험기능사, 건축도장기능사, 건축목공기능사, 공간정보융합기능사, 굴착기운전기능사, 기중기운전기능사, 도배기능사, 도화기능사, 로더운전기능사, 롤러운전기능사, 미장기능사, 방수기능사, 배관기능사, 불도저운전기능사, 비계기능사, 석공기능사, 실내건축기능사, 양화장치운전기능사, 온수온돌기능사, 유리시공기능사, 잠수기능사, 전산응용건축제도기능사, 전산응용토목제도기능사, 조경기능사, 조적기능사, 지게차운전기능사, 지도제작기능사, 지적기능사, 천공기운전기능사, 천장크레인운전기능사, 철근기능사, 철도토목기능사, 측량기능사, 컨테이너크레인운전기능사, 콘크리트기능사, 타워크레인운전기능사, 타일기능사, 항공사진기능사, 항로표지기능사
[광업자원] 화약류관리기술사, 화약류관리기사, 화약류관리산업기사, 화약취급기능사
[기계] 건설기계기술사, 공조냉동기계기술사, 금형기술사, 기계기술사, 산업기계설비기술사, 조선기술사, 차량기술사, 철도차량기술사, 항공기관기술사, 항공기체기술사, 건설기계정비기능장, 금형기능장, 기계가공기능장, 자동차정비기능장, 철도차량정비기능장, 건설기계설비기사, 건설기계정비기사, 공조냉동기계기사, 궤도장비정비기사, 그린전동자동차기사, 농업기계기사, 사출금형기사, 설비보전기사, 승강기기사, 일반기계기사, 자동차정비기사, 조선선체기사, 조선의장기사, 철도차량기사, 프레스금형기사, 항공기사, 건설기계설비산업기사, 건설기계정비산업기사, 공조냉동기계산업기사, 궤도장비정비산업기사, 기계설계산업기사, 기계조립산업기사, 농업기계산업기사, 사출금형산업기사, 설비보전산업기사, 스마트공장산업기사, 승강기산업기사, 자동차정비산업기사, 자동화설비산업기사, 정밀측정산업기사, 조선산업기사, 철도차량산업기사, 컴퓨터응용가공산업기사, 프레스금형산업기사, 항공산업기사, 건설기계정비기능사, 공조냉동기계기능사, 궤도장비정비기능사, 금형기능사, 기계가공조립기능사, 농업기계정비기능사, 반도체설비보전기능사, 선박기관정비기능사, 선체건조기능사, 선체설계기능사, 설비보전기능사, 스마트공장기능사, 승강기기능사, 이륜자동차정비기능사, 자동차보수도장기능사, 자동차정비기능사, 자동차차체수리기능사, 자동화설비기능사, 전산응용기계제도기능사, 정밀측정기능사, 철도차량정비기능사, 컴퓨터응용밀링기능사, 컴퓨터응용선반기능사, 표면실장장비기능사, 항공기정비기능사, 항공전기·전자정비기능사
[재료] 금속가공기술사, 금속재료기술사, 금속제련기술사, 세라믹기술사, 용접기술사, 표면처리기술사, 금속재료기능장, 압연기능장, 용접기능장, 제강기능장, 제선기능장, 주조기능장, 판금제관기능장, 표면처리기능장, 금속재료기사, 용접기사, 금속재료산업기사, 용접산업기사, 주조산업기사, 판금제관산업기사, 표면처리산업기사, 가스텅스텐아크용접기능사, 금속도장기능사, 금속재료시험기능사, 압연기능사, 열처리기능사, 이산화탄소가스아크용접기능사, 제강기능사, 제선기능사, 주조기능사, 축로기능사, 판금제관기능사, 표면처리기능사, 피복아크용접기능사
[화학] 화공기술사, 위험물기능장, 바이오화학제품제조기사, 정밀화학기사, 화공기사, 화약류제조기사, 화학분석기사, 바이오화학제품제조산업기사, 위험물산업기사, 화약류제조산업기사, 바이오공정기능사, 위험물기능사, 화학분석기능사
[섬유·의복] 섬유기술사, 의류기술사, 한복기능장, 섬유기사, 의류기사, 섬유디자인산업기사, 섬유산업기사, 신발산업기사, 패션디자인산업기사, 패션머천다이징산업기사, 남성복기능사, 봉제기능사, 세탁기능사, 신발제조기능사, 여성복기능사, 염색기능사(날염), 염색기능사(침염), 한복기능사
[전기·전자] 건축전기설비기술사, 발송배전기술사, 산업계측제어기술사, 전기응용기술사, 전기철도기술사, 전자응용기술사, 철도신호기술사, 전기기능장, 전자기능장, 광학기사, 로봇기구개발기사, 로봇소프트웨어개발기사, 로봇하드웨어개발기사, 의공기사, 임베디드기사, 전기공사기사, 전기기사, 전기철도기사, 전자기사, 철도신호기사, 3D프린터개발산업기사, 광학기기산업기사, 반도체커스텀레이아웃산업기사, 의공산업기사, 전기공사산업기사, 전기산업기사, 전기철도산업기사, 전자산업기사, 철도신호산업기사, 3D프린터운용기능사, 의료전자기능사, 임베디드기능사, 전기기능사, 전자기능사, 전자캐드기능사, 철도전기신호기능사
[정보통신] 정보관리기술사, 컴퓨터시스템응용기술사, 정보처리기사, 컴퓨터시스템기사, 사무자동화산업기사, 정보처리산업기사, 멀티미디어콘텐츠제작전문가, 정보기기운용기능사, 프로그래밍기능사
[식품가공] 수산제조기술사, 식품기술사, 제과기능장, 수산제조기사, 식육가공기사, 식품안전기사, 식품산업기사, 제과산업기사, 제빵산업기사, 떡제조기능사, 식품가공기능사, 제과기능사, 제빵기능사
[인쇄·목재·가구·공예] 귀금속가공기능장, 인쇄설계기사, 가구제작산업기사, 귀금속가공산업기사, 디지털인쇄산업기사, 보석감정산업기사, 보석디자인산업기사, 피아노조율산업기사, 가구제작기능사, 귀금속가공기능사, 도자공예기능사, 목공예기능사, 보석가공기능사, 보석감정사, 사진기능사, 석공예기능사, 인쇄기능사, 전자출판기능사, 피아노조율기능사
[농림어업] 농화학기술사, 산림기술사, 수산양식기술사, 시설원예기술사, 어업기술사, 종자기술사, 축산기술사, 산림기능장, 산림기사, 수산양식기사, 시설원예기사, 식물보호기사, 어업생산관리기사, 유기농업기사, 임산가공기사, 임업종묘기사, 종자기사, 축산기사, 화훼장식기사, 버섯산업기사, 산림산업기사, 수산양식산업기사, 식물보호산업기사, 어로산업기사, 유기농업산업기사, 종자산업기사, 축산산업기사, 화훼장식산업기사, 목재가공기능사, 버섯종균기능사, 산림기능사, 수산양식기능사, 식육처리기능사, 원예기능사, 유기농업기능사, 임업종묘기능사, 종자기능사, 축산기능사, 펄프종이제조기능사, 화훼장식기능사
[안전관리] 가스기술사, 건설안전기술사, 기계안전기술사, 비파괴검사기술사, 산업위생관리기술사, 소방기술사, 인간공학기술사, 전기안전기술사, 화공안전기술사, 가스기능장, 가스기사, 건설안전기사, 농작업안전보건기사, 누설비파괴검사기사, 방사선비파괴검사기사, 방재기사, 산업안전기사, 산업위생관리기사, 소방설비기사(기계분야), 소방설비기사(전기분야), 와전류비파괴검사기사, 인간공학기사, 자기비파괴검사기사, 초음파비파괴검사기사, 침투비파괴검사기사, 화재감식평가기사, 가스산업기사, 건설안전산업기사, 방사선비파괴검사산업기사, 산업안전산업기사, 산업위생관리산업기사, 소방설비산업기사(기계분야), 소방설비산업기사(전기분야), 자기비파괴검사산업기사, 초음파비파괴검사산업기사, 침투비파괴검사산업기사, 화재감식평가산업기사, 가스기능사, 방사선비파괴검사기능사, 자기비파괴검사기능사, 초음파비파괴검사기능사, 침투비파괴검사기능사
[환경·에너지] 기상예보기술사, 대기관리기술사, 소음진동기술사, 수질관리기술사, 자연환경관리기술사, 토양환경기술사, 폐기물처리기술사, 에너지관리기능장, 기상감정기사, 기상기사, 대기환경기사, 생물분류기사(동물), 생물분류기사(식물), 소음진동기사, 수질환경기사, 신재생에너지발전설비기사(태양광), 에너지관리기사, 온실가스관리기사, 자연생태복원기사, 토양환경기사, 폐기물처리기사, 환경위해관리기사, 대기환경산업기사, 소음진동산업기사, 수질환경산업기사, 신재생에너지발전설비산업기사(태양광), 에너지관리산업기사, 자연생태복원산업기사, 폐기물처리산업기사, 신재생에너지발전설비기능사(태양광), 에너지관리기능사, 환경기능사
"""

def get_todays_laws(api_key, target_date):
    all_laws_dict = {}
    search_date_range = f"{target_date}~{target_date}"
    
    print(f"\n📅 [{target_date}] 오늘 시행되는 법제처 법령 데이터를 수집합니다...")
    
    for target_type in ['law', 'histlaw']:
        page = 1
        while True:
            search_url = f"https://www.law.go.kr/DRF/lawSearch.do?OC={api_key}&target={target_type}&type=XML&efYd={search_date_range}&display=100&page={page}"
            try:
                response = session.get(search_url, headers=HEADERS, timeout=15)
                root = ET.fromstring(response.text)
            except Exception:
                break
                
            law_nodes = root.findall('.//law')
            if not law_nodes:
                break
                
            for law in law_nodes:
                law_id = law.find('법령일련번호').text if law.find('법령일련번호') is not None else ""
                law_name = law.find('법령명한글').text if law.find('법령명한글') is not None else "이름없음"
                enforce_date = law.find('시행일자').text if law.find('시행일자') is not None else "날짜없음"
                
                if not law_id or law_name in all_laws_dict:
                    continue
                    
                detail_url = f"https://www.law.go.kr/DRF/lawService.do?OC={api_key}&target={target_type}&MST={law_id}&type=XML"
                try:
                    detail_response = session.get(detail_url, headers=HEADERS, timeout=15)
                    detail_root = ET.fromstring(detail_response.text)
                    
                    reason_text = ""
                    for tag in ['.//개정이유', './/제개정이유']:
                        r_node = detail_root.find(tag)
                        if r_node is not None and r_node.text:
                            reason_text += r_node.text.strip() + "\n"
                            
                    jomuns = detail_root.findall('.//조문내용')
                    body_text = "\n".join([j.text.strip() for j in jomuns if j.text])
                    
                    stars = detail_root.findall('.//별표내용')
                    star_text = "\n".join([s.text.strip() for s in stars if s.text])
                    
                    full_text = f"[개정이유]\n{reason_text}\n[조문내용]\n{body_text}\n[별표내용]\n{star_text}"
                    full_text = full_text[:20000] 
                    
                except:
                    full_text = "서버 응답 지연"
                
                all_laws_dict[law_name] = {
                    "법령명": law_name,
                    "시행일자": enforce_date,
                    "주요 제·개정내용_원본": full_text
                }
                
                print(f"  📥 [가져오는 중...] {law_name}")
                time.sleep(0.5) 
                
            if len(law_nodes) < 100: 
                break
            page += 1 
            
    return list(all_laws_dict.values())

def apply_excel_formatting(filename, df_summary, df_detail):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='요약표', index=False)
        df_detail.to_excel(writer, sheet_name='상세분석', index=False)
        
    wb = load_workbook(filename)
    ws_detail = wb['상세분석']
    cols_to_resize = ['법령명', '주요 제·개정내용', '법령 관련 국가기술자격 종목', '활용도 심층분석']
    
    headers = [cell.value for cell in ws_detail[1]]
    for col_name in cols_to_resize:
        if col_name in headers:
            col_idx = headers.index(col_name) + 1
            col_letter = get_column_letter(col_idx)
            ws_detail.column_dimensions[col_letter].width = 45

    fill_even = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    for row_idx, row in enumerate(ws_detail.iter_rows(min_row=1), start=1):
        for cell in row:
            if row_idx == 1:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                if row_idx % 2 == 0:
                    cell.fill = fill_even
                else:
                    cell.fill = fill_odd

    ws_summary = wb['요약표']
    for col in ws_summary.columns:
        ws_summary.column_dimensions[col[0].column_letter].width = 20
    for row_idx, row in enumerate(ws_summary.iter_rows(min_row=1), start=1):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if row_idx == 1 or cell.column == 1:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    wb.save(filename)

def send_email_with_excel(filename, total_count, important_count):
    print("\n📧 이메일 발송을 준비합니다...")
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL 
    msg['Subject'] = f"🤖 [일일 모니터링] {FILE_PREFIX} 국가기술자격 관계 법령 분석"

    body = f"""
    자동화 서버에서 작성된 {FILE_PREFIX} 관계 법령 일일 모니터링 결과입니다.
    
    ▶ 오늘 시행되는 총 법령 수: {total_count}건
    ▶ 자격 규제 유의미(중요) 법령: {important_count}건
    """
    
    if important_count == 0:
        body += "\n※ 오늘은 국가기술자격 관련 행정 규제 변동(제·개정) 사항이 없습니다.\n"
        
    body += """
    상세한 심층 분석 결과는 첨부된 엑셀 파일을 확인해 주십시오.
    (본 메일은 GitHub Actions 클라우드 서버를 통해 매일 아침 자동 발송됩니다.)
    """
    
    msg.attach(MIMEText(body, 'plain'))

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= HRDKorea_Law_Report_{TARGET_DATE}.xlsx")
    msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.naver.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        receiver_list = [email.strip() for email in RECEIVER_EMAIL.split(',')]
        server.sendmail(SENDER_EMAIL, receiver_list, msg.as_string())
        server.quit()
        print(f"✅ 이메일 발송 성공! 사내 메일함을 확인하세요. (수신자: {RECEIVER_EMAIL})")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")

def main():
    print(f"\n[🚀 일일 자동화 모드] 가동 시작...")
    
    # 🚨 이 간판이 뜨면 무조건 2.5 모델 장착 완료!
    print(f"\n🚨 [V7 최종 완성본] 모델 교체 완료! (gemini-2.5-flash / 15초 쾌속 모드) 🚨")
    
    laws = get_todays_laws(LAW_API_KEY, TARGET_DATE)
    
    filename = os.path.join(CURRENT_FOLDER, f"HRDKorea_Law_Report_{TARGET_DATE}.xlsx")
    important_laws = []
    relation_count = 0
    
    if not laws:
        print("오늘은 새로 시행되는 전체 법령이 없습니다. (빈 보고서 발송 준비)")
    else:
        print(f"🏎️ 제미나이(Gemini) AI가 오늘 시행되는 {len(laws)}건의 법령을 심사합니다...")
        
        relation_keywords = ["자격", "기술", "면허", "기사", "기능", "안전", "환경", "폐기물", "시공", "관리", "검사"]
        
        for index, law in enumerate(laws):
            if any(kw in law["법령명"] + law["주요 제·개정내용_원본"] for kw in relation_keywords):
                relation_count += 1

            print(f"[{index+1}/{len(laws)}] {law['법령명']} 분석 중... (15초 대기) ", end="", flush=True)
            
            # 🚨 15초 쾌속 모드 (1.5 플래시 모델은 1분에 15개 허용)
            time.sleep(15) 
            
            prompt = f"""
            당신은 한국산업인력공단의 국가기술자격 규제 심사 수석 연구원입니다.
            매우 보수적이고 엄격한 잣대로 법령을 평가해야 합니다. AI의 과잉 추론이나 상상력을 절대 배제하십시오.

            [국가기술자격 사전]
            {QNET_CERTS}

            [분석 대상]
            법령명: {law['법령명']}
            법령내용 원문(개정이유, 조문, 별표 포함): {law['주요 제·개정내용_원본']}

            [🚨 절대 준수 사항 (금지어 및 제한 조건)]
            1. 간접 추론 금지: "정부 지원이 늘어나니 산업이 커질 것이고, 따라서 자격증 수요도 늘어날 것이다" 같은 간접적/연쇄적 추론은 절대 금지합니다.
            2. 무관한 법령 배제: 법령 원문에 자격증 소지자의 '의무 선임', '배치 기준 신설/강화/완화', '자격 요건'에 대한 "직접적인" 언급이나 행정 규제 변화가 없다면 무조건 "일반/무관" 및 "변동 없음"으로 판정하세요.
            
            [자격 활용도 판단 기준]
            1. 대폭 증가: 특정 국가기술자격증 소지자의 '의무 선임'이나 '배치 기준'이 법적으로 신설되거나 대폭 강화된 경우 (직접적 명시 필수).
            2. 소폭 증가: 자격증 취득 시 가산점 부여, 우대 조건 신설 등 직접적인 법적 혜택이 추가된 경우.
            3. 변동 없음: 단순 산업 진흥, 예산 지원, 행정 절차 변경, 타법 개정에 따른 부처명 변경 등 직접적인 자격 규제 변동이 없는 모든 경우.
            4. 소폭 감소: 특정 자격증 외에 다른 경력이나 학력으로도 선임될 수 있도록 대체 요건이 추가된 경우.
            5. 대폭 감소: 기존 자격증 선임 의무가 법적으로 완전히 폐지되거나 규제가 대폭 완화된 경우.

            [출력 JSON 형식]
            {{
                "분류": "중요 또는 일반/무관 (자격 규제와 관련된 '직접적인' 명시가 없다면 가차 없이 '일반/무관'으로 분류할 것)",
                "요약": "법령의 핵심 목적과 주요 내용을 개조식('~함', '~임')으로 요약 (줄바꿈 없이 작성)",
                "종목": "매칭된 국가기술자격증 나열 (없으면 '없음')",
                "활용도_구분": "대폭 증가, 소폭 증가, 변동 없음, 소폭 감소, 대폭 감소 중 택 1",
                "활용도_분석": "왜 그렇게 판단했는지 행정 규제 관점에서 매우 엄격하게 서술"
            }}
            """
            
            try:
                response = model.generate_content(
                    prompt,
                    generation_config={"response_mime_type": "application/json"}
                )
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]
                    
                ai_data = json.loads(raw_text.strip())
                
                judgement = ai_data.get("분류", "오류")
                summary = ai_data.get("요약", "요약 불가")
                related_certs = ai_data.get("종목", "없음")
                utility_category = ai_data.get("활용도_구분", "분류 안됨")
                utility_impact = ai_data.get("활용도_분석", "분석 불가")
                
            except Exception as e:
                print(f"❌ [에러 발생] {e} (10초 후 다음 법령 진행)")
                time.sleep(10)
                continue 
            
            if "중요" in judgement and "변동 없음" not in utility_category:
                important_laws.append({
                    "연번": len(important_laws) + 1,
                    "시행일자": law["시행일자"],
                    "법령명": law["법령명"],
                    "주요 제·개정내용": summary,  
                    "법령 관련 국가기술자격 종목": related_certs,
                    "활용도 구분": utility_category,
                    "활용도 심층분석": utility_impact
                })
                print(f"👉 [채택] {utility_category} | {related_certs}")
            else:
                print(f"❌ [패스]")

    if not important_laws:
        formatted_date = f"{TARGET_DATE[:4]}-{TARGET_DATE[4:6]}-{TARGET_DATE[6:]}"
        important_laws.append({
            "연번": "-",
            "시행일자": formatted_date,
            "법령명": "해당 없음",
            "주요 제·개정내용": "오늘의 국가기술자격 관련 법령 제·개정사항이 없습니다.",
            "법령 관련 국가기술자격 종목": "-",
            "활용도 구분": "-",
            "활용도 심층분석": "-"
        })
        
    real_important_count = len(important_laws) if important_laws[0]["연번"] != "-" else 0
    
    summary_data = {
        "구분": ["오늘의 시행법령 총계", "국가기술자격 관계 법령", "분석 및 관련 높은 법령"],
        "건수": [len(laws), relation_count, real_important_count]
    }
    
    df_summary = pd.DataFrame(summary_data)
    df_detail = pd.DataFrame(important_laws)
    
    apply_excel_formatting(filename, df_summary, df_detail)
    send_email_with_excel(filename, len(laws), real_important_count)
    
    print("\n==========================================")
    print(f"🎉 오늘의 엑셀 공장 가동이 완료되었습니다: {filename}")

if __name__ == "__main__":
    main()
