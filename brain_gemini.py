import json
import re
import time
from google import genai
from google.genai import types

# 💡 1단계 config 파일에서 제미나이 키와 자격증 리스트를 가져옵니다.
from config import GEMINI_API_KEY, QNET_CERTS

# 제미나이 클라이언트 초기화
client = genai.Client(api_key=GEMINI_API_KEY)

# ==========================================
# AI 분석 메인 함수
# ==========================================
def run_ai_analysis(law, attempt_count=5):
    """
    [V29 모듈화] 제미나이에게 마크다운 법령을 주고 JSON으로 결과를 받아옵니다.
    """
    prompt = f"""
    당신은 '한국산업인력공단(HRDK)'의 국가기술자격 정책 수석 연구원입니다.
    아래 [최신 법령 제·개정 내용]과 [국가기술자격 491개 종목 사전]을 바탕으로,
    이 법령이 국가기술자격의 수요, 활용도, 응시자격, 배치기준 등에 미칠 영향을 정밀하게 분석하십시오.

    [국가기술자격 491개 종목 사전]
    {QNET_CERTS}

    [최신 법령 제·개정 내용 (마크다운)]
    {law['원본']}

[🚨 수석 연구원 분석 지침 (신구비교 및 숲과 나무)]
    1. (숲 파악) '개정이유'와 '제1조(목적)'를 통해 이 법령이 어떤 산업/분야를 규제하거나 지원하는지 전체적인 배경을 먼저 파악하십시오.
    2. (나무 분석) 전체 조문이 아닌, '이번에 바뀐 핵심 조문(신설/개정)' 내용에 집중하여 규제 강화, 기준 완화, 인력 배치 기준 신설 등 실질적인 변화를 캐치하십시오.
    3. (파급 효과 추론) 위 분석을 바탕으로, 명시된 자격증뿐만 아니라 내용상 '간접적으로 연관된 자격증'도 모두 추론하십시오. 연관성이 전혀 없으면 "종목": "", "분류": "일반"으로 처리하십시오.

    [분석 및 판단 기준]
    1. 분류: '연관높음', '단순관련', '일반' 중 택1
    2. 활용도_구분: '연관높음'인 경우에만 [대폭 증가, 소폭 증가, 소폭 감소, 대폭 감소] 중 선택. 그 외는 빈칸("").
    3. 소관부처: 정부 부처명 추출.
    4. 개정유형: 제정, 일부개정, 전부개정 등 성격 추출.
    5. 조문리스트: 연관된 조문이 여러 개일 경우 **반드시 모두 추출**하여 배열 형태로 작성
       - [주의] 제O조 형태: {{"조문명": "제23조의2", "숫자": "23.2"}} 
       - [주의] 별표 형태: "별표1"이 아닌 "별표 1"과 같이 반드시 띄어쓰기를 지켜서 작성. (숫자는 빈칸 "")
    6. AI_신뢰도: 본 분석에 대한 AI의 객관적 확신도 ('높음', '보통', '낮음' 중 택1)
       - 높음: 법령(바뀐조문/별표)에 '국가기술자격 종목 명칭'이 정확히 텍스트로 명시된 경우
       - 보통: 명칭은 없으나 직무 내용상 연관성이 매우 높다고 강하게 추론되는 경우
       - 낮음: 연관성을 억지로 논리적 비약을 통해 연결해야 하는 경우
    7. 검토필요: 실무자의 교차 검증이 반드시 필요한 경우 'O', 아니면 'X'
       - [체크(O) 필수 조건]: ① 'AI_신뢰도'가 '보통/낮음'이거나, ② '활용도_구분'이 '대폭 증가/감소'로 파급력이 큰 경우
    8. 검토사유: '검토필요'가 'O'인 경우에 한해, 그 이유를 구체적으로 작성 (예: "자격 명칭이 직접 명시되지 않아 실무자 확인 요망"). '검토필요'가 'X'이면 빈칸("").

    🔥 [작성 가이드라인: 주요 제·개정내용 (요약)] 🔥
    - 실제 개정된 조항과 객관적인 팩트만 글머리 기호('-')를 사용하여 나열하십시오.

    🔥 [작성 가이드라인: 활용도 분석 상세] 🔥
    - [1000자 이내 제한] 1000자 이내로 간결하고 명확하게 분석하십시오.
    - ① 개정 배경, ② 방향성, ③ 파급효과에 집중하십시오.

    [🚨 JSON 작성 절대 규칙]
    1. 출력은 단 1개의 JSON 객체({{ }})만.
    2. (큰따옴표 전면 금지) 모든 텍스트 내부에 절대 큰따옴표(") 금지. 강조는 작은따옴표(') 사용.
    3. (실제 엔터키 금지) 텍스트 내부 실제 줄바꿈 대신 '\\n' 기호 사용.
    4. (종목 포맷팅) 각 직무분야 시작 시 'O ' 꼭지 사용 및 줄바꿈 기호('\\n') 사용.

    [출력 JSON 포맷]
    {{
        "소관부처": "...",
        "개정유형": "일부개정/전부개정/제정/폐지 중 택1",
        "요약": "개정내용 1문장 요약",
        "분류": "연관높음/단순관련/일반 중 택1",
        "종목": "관련 자격증 이름만 쉼표로 나열",
        "활용도_구분": "대폭 증가/소폭 증가/현상 유지/감소 중 택1",
        "활용도_분석": "자격증 수요 및 활용도 변화에 대한 분석 (3문장 이내)",
        "AI_신뢰도": "높음/보통/낮음 중 택1",
        "검토필요": "O/X 중 택1",
        "검토사유": "O인 경우 그 이유 (3문장 이내)",
        "조문리스트": [
            {{"조문명": "제1조(목적)", "숫자": "1"}},
            {{"조문명": "제2조의2(정의)", "숫자": "2:2"}},
            {{"조문명": "별표 1", "숫자": ""}}
        ]
    }}
    """

    for attempt in range(attempt_count):
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash', # 기본은 똑똑한 2.5로 유지!
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json", 
                    max_output_tokens=2048, # 🔥 글자 수 제한 (끊김 방지 및 속도 향상)
                    temperature=0.0  
                )
            )
            raw_text = response.text.strip().replace("```json", "").replace("```", "")
            data = json.loads(raw_text, strict=False)
            
            jomun_list = data.get("조문리스트", [])
            if not jomun_list or not isinstance(jomun_list, list):
                jomun_list = [{"조문명": "내용 확인", "숫자": ""}]
                
            links_str_list = []
            names_str_list = []
            
            for j in jomun_list:
                j_name = j.get("조문명", "확인불가")
                if "별표" in j_name:
                    j_name = re.sub(r'별표\s*(\d+)', r'별표 \1', j_name)
                    
                j_num = str(j.get("숫자", "")).strip().replace(".", ":")
                anchor = f"#J{j_num}" if j_num else ""
                
                if j_name == "내용 확인":
                    names_str_list.append("전체 (세부 조문 미지정)")
                    links_str_list.append(f"▶ {law['법령명']}\n{law['링크']}")
                else:
                    names_str_list.append(j_name)
                    links_str_list.append(f"▶ {law['법령명']} {j_name}\n{law['링크']}{anchor}")
                
            links_str = "\n\n".join(links_str_list)
            names_str = ", ".join(names_str_list)
            
            law_info = {
                "시행일자": law["시행일자"],
                "소관부처": data.get("소관부처", ""),
                "법령명": law["법령명"],
                "개정유형": data.get("개정유형", ""),
                "주요 제·개정내용": data.get("요약", ""),
                "법령 관련 국가기술자격 종목": data.get("종목", ""),
                "활용도 분석 구분": data.get("활용도_구분", ""),
                "활용도 분석 상세": data.get("활용도_분석", ""),
                "근거 조문": names_str,
                "AI 신뢰도": data.get("AI_신뢰도", ""),
                "검토 필요": data.get("검토필요", "X"),
                "검토 사유": data.get("검토사유", ""),
                "조문별 다이렉트 링크": links_str
            }
            return True, data.get("분류", ""), law_info
            
        except Exception as e:
            error_msg = str(e)
            # 🔥 [503 방어 로직] 서버가 뻗으면 길게 대기합니다!
            if "503" in error_msg or "high demand" in error_msg.lower():
                wait_time = 60 * (attempt + 1)
                print(f"  🚨 [서버 폭주 대피] {wait_time}초간 딥슬립 후 재시도... ", end="", flush=True)
            else:
                wait_time = 15 * (attempt + 1)
                print(f"  🚨 [일반 에러: {e}] {wait_time}초 대기... ", end="", flush=True)
                
            time.sleep(wait_time)
            
    return False, "", {"error": error_msg if 'error_msg' in locals() else "재시도 초과"}