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
# 1. 환경 변수 (GitHub 보안 금고에서 자동으로 불러옴)
# ==========================================
LAW_API_KEY = os.environ.get("LAW_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")

# ==========================================
# 2. 한국 시간(KST) 기준 오늘 날짜 자동 세팅
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

QNET_CERTS = """
[건설/건축/토목] 건축구조기술사, 건축시공기술사, 토목구조기술사, 토질및기초기술사, 건설안전기술사, 건축기사, 건축산업기사, 건축설비기사, 실내건축기사, 실내건축기능사, 토목기사, 건설재료시험기사, 측량및지형공간정보기사, 콘크리트기사, 방수산업기사, 전산응용건축제도기능사
[전기/전자/통신] 발송배전기술사, 건축전기설비기술사, 전기안전기술사, 전기응용기술사, 전기철도기술사, 철도신호기술사, 전기기능장, 전기기사, 전기산업기사, 전기공사기사, 전기공사산업기사, 전기기능사, 전자기사, 전자계산기기사, 무선설비기사, 정보통신기사, 통신설비기능장, 신재생에너지발전설비기사(태양광)
[안전관리] 산업안전지도사, 산업보건지도사, 산업안전기사, 산업안전산업기사, 건설안전기사, 건설안전산업기사, 소방설비기사(전기), 소방설비기사(기계), 가스기술사, 가스기능장, 가스기사, 산업위생관리기사, 인간공학기사, 화재감식평가기사
[환경/에너지] 대기관리기술사, 수질관리기술사, 대기환경기사, 수질환경기사, 폐기물처리기사, 소음진동기사, 자연생태복원기사, 온실가스관리기사, 환경기능사, 에너지관리기능장, 에너지관리기사
[기계/금속/재료] 기계안전기술사, 일반기계기사, 건설기계설비기사, 공조냉동기계기사, 설비보전기사, 승강기기사, 용접기능장, 용접기사, 자동차정비기사, 금속재료기사, 표면처리기사
[화학/위험물] 위험물기능장, 위험물산업기사, 위험물기능사, 화공안전기술사, 화공기사, 화학분석기사
[정보기술(IT)] 컴퓨터시스템응용기술사, 정보관리기술사, 정보처리기사, 정보처리산업기사, 정보보안기사, 빅데이터분석기사, 정보기기운용기능사
[농림/축산/어업] 산림기술사, 조경기술사, 산림기사, 조경기사, 종자기사, 유기농업기사, 식물보호기사, 축산기사, 식육처리기능사, 수산양식기사
[디자인/인쇄] 시각디자인기사, 제품디자인기사, 컬러리스트기사, 컴퓨터그래픽스운용기능사, 웹디자인기능사, 인쇄기사, 광고도장기능사
[서비스/기타] 직업상담사1급, 직업상담사2급, 임상심리사1급, 임상심리사2급, 사회조사분석사, 미용장, 미용사(일반,피부,네일,메이크업), 조리기능장, 한식조리기능사, 제과기능장
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
    cols_to_resize = ['법령명', '주요 제·개정내용', '법령 관련 국가기술자격 종목', '효용성 심층분석']
    
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
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("✅ 이메일 발송 성공! 사내 메일함을 확인하세요.")
    except Exception as e:
        print(f"❌ 이메일 발송 실패: {e}")

def main():
    print(f"\n[🚀 일일 자동화 모드] 가동 시작...")
    laws = get_todays_laws(LAW_API_KEY, TARGET_DATE)
    
    if not laws:
        print("오늘은 새로 시행되는 법령이 없습니다. 시스템을 종료합니다.")
        return
            
    filename = os.path.join(CURRENT_FOLDER, f"HRDKorea_Law_Report_{TARGET_DATE}.xlsx")
    important_laws = []
    
    print(f"\n🏎️ 제미나이(Gemini) 2.5 AI가 오늘 시행되는 {len(laws)}건의 법령을 심사합니다...")
    
    relation_keywords = ["자격", "기술", "면허", "기사", "기능", "안전", "환경", "폐기물", "시공", "관리", "검사"]
    relation_count = 0
    
    for index, law in enumerate(laws):
        if any(kw in law["법령명"] + law["주요 제·개정내용_원본"] for kw in relation_keywords):
            relation_count += 1

        print(f"[{index+1}/{len(laws)}] {law['법령명']} 분석 중... ", end="", flush=True)
        
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
        
        [자격 효용성 판단 기준]
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
            "효용성_구분": "대폭 증가, 소폭 증가, 변동 없음, 소폭 감소, 대폭 감소 중 택 1",
            "효용성_분석": "왜 그렇게 판단했는지 행정 규제 관점에서 매우 엄격하게 서술"
        }}
        """
        
        try:
            response = model.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"}
            )
            # 가끔 AI가 ```json 포맷을 붙이는 것을 방지
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
                
            ai_data = json.loads(raw_text.strip())
            
            judgement = ai_data.get("분류", "오류")
            summary = ai_data.get("요약", "요약 불가")
            related_certs = ai_data.get("종목", "없음")
            utility_category = ai_data.get("효용성_구분", "분류 안됨")
            utility_impact = ai_data.get("효용성_분석", "분석 불가")
            
        except Exception as e:
            print(f"❌ [에러] 통신 실패. 3초 대기")
            time.sleep(3)
            continue 
        
        if "중요" in judgement and "변동 없음" not in utility_category:
            important_laws.append({
                "연번": len(important_laws) + 1,
                "시행일자": law["시행일자"],
                "법령명": law["법령명"],
                "주요 제·개정내용": summary,  
                "법령 관련 국가기술자격 종목": related_certs,
                "효용성 구분": utility_category,
                "효용성 심층분석": utility_impact 
            })
            print(f"👉 [채택] {utility_category} | {related_certs}")
        else:
            print(f"❌ [패스]")
            
    summary_data = {
        "구분": ["오늘의 시행법령 총계", "국가기술자격 관계 법령", "활용 및 관련 높은 법령"],
        "건수": [len(laws), relation_count, len(important_laws)]
    }
    df_summary = pd.DataFrame(summary_data)
    df_detail = pd.DataFrame(important_laws)
    
    apply_excel_formatting(filename, df_summary, df_detail)
    send_email_with_excel(filename, len(laws), len(important_laws))
    
    print("\n==========================================")
    print(f"🎉 오늘의 엑셀 공장 가동이 완료되었습니다: {filename}")

if __name__ == "__main__":
    main()
