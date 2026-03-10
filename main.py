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
# 1. 환경 변수 (GitHub Secrets에서 안전하게 호출)
# ==========================================
LAW_API_KEY = os.environ.get("LAW_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")      # 네이버 아이디
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")  # 네이버 앱 비밀번호
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")  # 수신받을 이메일

# ==========================================
# 2. 날짜 및 AI 세팅 (가장 똑똑한 2.5-flash 모델 고정!)
# ==========================================
KST = timezone(timedelta(hours=9))
today = datetime.now(KST)
TARGET_DATE = today.strftime("%Y%m%d")
FILE_PREFIX = today.strftime("%Y년_%m월_%d일")

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-2.5-flash') 

HEADERS = {'User-Agent': 'Mozilla/5.0'}
session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

# ==========================================
# 3. 공단 전용 491개 자격 종목 (CSV 에러 방지용 내장형)
# ==========================================
QNET_CERTS = """
공공조달관리사, 공장관리기술사, 포장기술사, 품질관리기술사, 포장기사, 품질경영기사, 포장산업기사, 품질경영산업기사, 사회조사분석사1급, 사회조사분석사2급, 소비자전문상담사1급, 소비자전문상담사2급, 컨벤션기획사1급, 컨벤션기획사2급, 이러닝운영관리사, 국제의료관광코디네이터, 임상심리사1급, 임상심리사2급, 직업상담사1급, 직업상담사2급, 제품디자인기술사, 시각디자인기사, 제품디자인기사, 컬러리스트기사, 시각디자인산업기사, 제품디자인산업기사, 컬러리스트산업기사, 웹디자인개발기능사, 제품응용모델링기능사, 컴퓨터그래픽기능사, 철도운송산업기사, 농기계운전기능사, 텔레마케팅관리사, 미용장, 이용장, 스포츠경영관리사, 미용사(네일), 미용사(메이크업), 미용사(일반), 미용사(피부), 이용사, 조리기능장, 복어조리산업기사, 양식조리산업기사, 일식조리산업기사, 중식조리산업기사, 한식조리산업기사, 복어조리기능사, 양식조리기능사, 일식조리기능사, 조주기능사, 중식조리기능사, 한식조리기능사, 금속재창호기능사, 플라스틱창호기능사, 건축구조기술사, 건축기계설비기술사, 건축시공기술사, 건축품질시험기술사, 교통기술사, 농어업토목기술사, 도로및공항기술사, 도시계획기술사, 상하수도기술사, 수자원개발기술사, 조경기술사, 지적기술사, 지질및지반기술사, 철도기술사, 측량및지형공간정보기술사, 토목구조기술사, 토목시공기술사, 토목품질시험기술사, 토질및기초기술사, 항만및해안기술사, 해양기술사, 건축목재시공기능장, 건축일반시공기능장, 배관기능장, 잠수기능장, 건설재료시험기사, 건축기사, 건축설비기사, 교통기사, 도시계획기사, 실내건축기사, 응용지질기사, 조경기사, 지적기사, 철도토목기사, 측량및지형공간정보기사, 콘크리트기사, 토목기사, 항로표지기사, 해양공학기사, 해양자원개발기사, 해양환경기사, 건설재료시험산업기사, 건축목공산업기사, 건축산업기사, 건축설비산업기사, 건축일반시공산업기사, 공간정보융합산업기사, 교통산업기사, 방수산업기사, 배관산업기사, 실내건축산업기사, 잠수산업기사, 조경산업기사, 지적산업기사, 측량및지형공간정보산업기사, 콘크리트산업기사, 토목산업기사, 항로표지산업기사, 해양조사산업기사, 거푸집기능사, 건설재료시험기능사, 건축도장기능사, 건축목공기능사, 공간정보융합기능사, 굴착기운전기능사, 기중기운전기능사, 도배기능사, 도화기능사, 로더운전기능사, 롤러운전기능사, 미장기능사, 방수기능사, 배관기능사, 불도저운전기능사, 비계기능사, 석공기능사, 실내건축기능사, 양화장치운전기능사, 온수온돌기능사, 유리시공기능사, 잠수기능사, 전산응용건축제도기능사, 전산응용토목제도기능사, 조경기능사, 조적기능사, 지게차운전기능사, 지도제작기능사, 지적기능사, 천공기운전기능사, 천장크레인운전기능사, 철근기능사, 철도토목기능사, 측량기능사, 컨테이너크레인운전기능사, 콘크리트기능사, 타워크레인운전기능사, 타일기능사, 항공사진기능사, 항로표지기능사, 화약류관리기술사, 화약류관리기사, 화약류관리산업기사, 화약취급기능사, 건설기계기술사, 공조냉동기계기술사, 금형기술사, 기계기술사, 산업기계설비기술사, 조선기술사, 차량기술사, 철도차량기술사, 항공기관기술사, 항공기체기술사, 건설기계정비기능장, 금형기능장, 기계가공기능장, 자동차정비기능장, 철도차량정비기능장, 건설기계설비기사, 건설기계정비기사, 공조냉동기계기사, 궤도장비정비기사, 그린전동자동차기사, 농업기계기사, 사출금형기사, 설비보전기사, 승강기기사, 일반기계기사, 자동차정비기사, 조선선체기사, 조선의장기사, 철도차량기사, 프레스금형기사, 항공기사, 건설기계설비산업기사, 건설기계정비산업기사, 공조냉동기계산업기사, 궤도장비정비산업기사, 기계설계산업기사, 기계조립산업기사, 농업기계산업기사, 사출금형산업기사, 설비보전산업기사, 스마트공장산업기사, 승강기산업기사, 자동차정비산업기사, 자동화설비산업기사, 정밀측정산업기사, 조선산업기사, 철도차량산업기사, 컴퓨터응용가공산업기사, 프레스금형산업기사, 항공산업기사, 건설기계정비기능사, 공조냉동기계기능사, 궤도장비정비기능사, 금형기능사, 기계가공조립기능사, 농업기계정비기능사, 반도체설비보전기능사, 선박기관정비기능사, 선체건조기능사, 선체설계기능사, 설비보전기능사, 스마트공장기능사, 승강기기능사, 이륜자동차정비기능사, 자동차보수도장기능사, 자동차정비기능사, 자동차차체수리기능사, 자동화설비기능사, 전산응용기계제도기능사, 정밀측정기능사, 철도차량정비기능사, 컴퓨터응용밀링기능사, 컴퓨터응용선반기능사, 표면실장장비기능사, 항공기정비기능사, 항공전기·전자정비기능사, 금속가공기술사, 금속재료기술사, 금속제련기술사, 세라믹기술사, 용접기술사, 표면처리기술사, 금속재료기능장, 압연기능장, 용접기능장, 제강기능장, 제선기능장, 주조기능장, 판금제관기능장, 표면처리기능장, 금속재료기사, 용접기사, 금속재료산업기사, 용접산업기사, 주조산업기사, 판금제관산업기사, 표면처리산업기사, 가스텅스텐아크용접기능사, 금속도장기능사, 금속재료시험기능사, 압연기능사, 열처리기능사, 이산화탄소가스아크용접기능사, 제강기능사, 제선기능사, 주조기능사, 축로기능사, 판금제관기능사, 표면처리기능사, 피복아크용접기능사, 화공기술사, 위험물기능장, 바이오화학제품제조기사, 정밀화학기사, 화공기사, 화약류제조기사, 화학분석기사, 바이오화학제품제조산업기사, 위험물산업기사, 화약류제조산업기사, 바이오공정기능사, 위험물기능사, 화학분석기능사, 섬유기술사, 의류기술사, 한복기능장, 섬유기사, 의류기사, 섬유디자인산업기사, 섬유산업기사, 신발산업기사, 패션디자인산업기사, 패션머천다이징산업기사, 남성복기능사, 봉제기능사, 세탁기능사, 신발제조기능사, 여성복기능사, 염색기능사(날염), 염색기능사(침염), 한복기능사, 건축전기설비기술사, 발송배전기술사, 산업계측제어기술사, 전기응용기술사, 전기철도기술사, 전자응용기술사, 철도신호기술사, 전기기능장, 전자기능장, 광학기사, 로봇기구개발기사, 로봇소프트웨어개발기사, 로봇하드웨어개발기사, 의공기사, 임베디드기사, 전기공사기사, 전기기사, 전기철도기사, 전자기사, 철도신호기사, 3D프린터개발산업기사, 광학기기산업기사, 반도체커스텀레이아웃산업기사, 의공산업기사, 전기공사산업기사, 전기산업기사, 전기철도산업기사, 전자산업기사, 철도신호산업기사, 3D프린터운용기능사, 의료전자기능사, 임베디드기능사, 전기기능사, 전자기능사, 전자캐드기능사, 철도전기신호기능사, 정보관리기술사, 컴퓨터시스템응용기술사, 정보처리기사, 컴퓨터시스템기사, 사무자동화산업기사, 정보처리산업기사, 멀티미디어콘텐츠제작전문가, 정보기기운용기능사, 프로그래밍기능사, 수산제조기술사, 식품기술사, 제과기능장, 수산제조기사, 식육가공기사, 식품안전기사, 식품산업기사, 제과산업기사, 제빵산업기사, 떡제조기능사, 식품가공기능사, 제과기능사, 제빵기능사, 귀금속가공기능장, 인쇄설계기사, 가구제작산업기사, 귀금속가공산업기사, 디지털인쇄산업기사, 보석감정산업기사, 보석디자인산업기사, 피아노조율산업기사, 가구제작기능사, 귀금속가공기능사, 도자공예기능사, 목공예기능사, 보석가공기능사, 보석감정사, 사진기능사, 석공예기능사, 인쇄기능사, 전자출판기능사, 피아노조율기능사, 농화학기술사, 산림기술사, 수산양식기술사, 시설원예기술사, 어업기술사, 종자기술사, 축산기술사, 산림기능장, 산림기사, 수산양식기사, 시설원예기사, 식물보호기사, 어업생산관리기사, 유기농업기사, 임산가공기사, 임업종묘기사, 종자기사, 축산기사, 화훼장식기사, 버섯산업기사, 산림산업기사, 수산양식산업기사, 식물보호산업기사, 어로산업기사, 유기농업산업기사, 종자산업기사, 축산산업기사, 화훼장식산업기사, 목재가공기능사, 버섯종균기능사, 산림기능사, 수산양식기능사, 식육처리기능사, 원예기능사, 유기농업기능사, 임업종묘기능사, 종자기능사, 축산기능사, 펄프종이제조기능사, 화훼장식기능사, 가스기술사, 건설안전기술사, 기계안전기술사, 비파괴검사기술사, 산업위생관리기술사, 소방기술사, 인간공학기술사, 전기안전기술사, 화공안전기술사, 가스기능장, 가스기사, 건설안전기사, 농작업안전보건기사, 누설비파괴검사기사, 방사선비파괴검사기사, 방재기사, 산업안전기사, 산업위생관리기사, 소방설비기사(기계분야), 소방설비기사(전기분야), 와전류비파괴검사기사, 인간공학기사, 자기비파괴검사기사, 초음파비파괴검사기사, 침투비파괴검사기사, 화재감식평가기사, 가스산업기사, 건설안전산업기사, 방사선비파괴검사산업기사, 산업안전산업기사, 산업위생관리산업기사, 소방설비산업기사(기계분야), 소방설비산업기사(전기분야), 자기비파괴검사산업기사, 초음파비파괴검사산업기사, 침투비파괴검사산업기사, 화재감식평가산업기사, 가스기능사, 방사선비파괴검사기능사, 자기비파괴검사기능사, 초음파비파괴검사기능사, 침투비파괴검사기능사, 기상예보기술사, 대기관리기술사, 소음진동기술사, 수질관리기술사, 자연환경관리기술사, 토양환경기술사, 폐기물처리기술사, 에너지관리기능장, 기상감정기사, 기상기사, 대기환경기사, 생물분류기사(동물), 생물분류기사(식물), 소음진동기사, 수질환경기사, 신재생에너지발전설비기사(태양광), 에너지관리기사, 온실가스관리기사, 자연생태복원기사, 토양환경기사, 폐기물처리기사, 환경위해관리기사, 대기환경산업기사, 소음진동산업기사, 수질환경산업기사, 신재생에너지발전설비산업기사(태양광), 에너지관리산업기사, 자연생태복원산업기사, 폐기물처리산업기사, 신재생에너지발전설비기능사(태양광), 에너지관리기능사, 환경기능사
"""

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
                
                # [에러 방지망 1] 법제처 서버 응답이 정상이 아니면 바로 뱉어내기
                if response.status_code != 200:
                    print(f"❌ [서버 통신 에러] 상태 코드: {response.status_code}")
                    break
                    
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
                
            except Exception as e: 
                # [에러 방지망 2] 수집 중 튕기면 원인을 명확하게 화면에 출력
                print(f"❌ [데이터 수집/파싱 에러 발생] {e}")
                print(f"👉 원인(서버 메시지): {response.text[:200]}")
                break
                
    return list(all_laws_dict.values())

def apply_excel_formatting(filename, df_summary, df_detail):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='요약표', index=False)
        df_detail.to_excel(writer, sheet_name='상세분석', index=False)
    
    wb = load_workbook(filename)
    ws = wb['상세분석']
    
    # 엑셀 헤더: 선생님이 지시하신 '활용도 분석' 명칭 적용!
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
        # 네이버 메일 발송 세팅
        server = smtplib.SMTP('smtp.naver.com', 587)
        server.starttls()
        # 네이버는 로그인 시 보통 '@naver.com' 앞의 아이디만 사용합니다.
        server.login(SENDER_EMAIL.split('@')[0], EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("✅ 네이버 메일 발송 성공!")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")

def main():
    laws = get_todays_laws(LAW_API_KEY, TARGET_DATE)
    if not laws: 
        print("오늘은 새로 시행되는 법령이 없습니다. 푹 쉬십시오!")
        return
    
    important_laws = []
    print(f"\n🏎️ {len(laws)}건 정밀 분석(V13) 시작...")
    
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
            
            # [에러 방지망 3] 제미나이가 박스(리스트) 채로 던져도 알맹이만 꺼내는 패치
            data = json.loads(raw_text)
            if isinstance(data, list): 
                if len(data) > 0:
                    data = data[0]
                else:
                    data = {} # 빈 리스트 방어
            
            if data.get("분류") == "중요" and data.get("활용도_구분") != "변동 없음":
                important_laws.append({
                    "시행일자": law["시행일자"],
                    "법령명": law["법령명"],
                    "요약": data.get("요약", "요약 없음"),
                    "관련종목": data.get("종목", "없음"),
                    "구분": data.get("활용도_구분", "분류 불가"),
                    "심층분석": data.get("활용도_분석", "분석 불가")
                })
                print("👉 [채택]")
            else: print("❌ [패스]")
        except Exception as e: 
            print(f"⚠️ [에러] AI 분석 중 문제 발생: {e}")
        
    if important_laws:
        df_detail = pd.DataFrame(important_laws)
        df_summary = pd.DataFrame({"구분": ["총 시행법령", "유의미 법령"], "건수": [len(laws), len(important_laws)]})
        fname = f"HRD_Daily_Report_{TARGET_DATE}.xlsx"
        apply_excel_formatting(fname, df_summary, df_detail)
        send_naver_email(fname, len(laws), len(important_laws))

if __name__ == "__main__":
    main()
