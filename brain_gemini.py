import json
import re
from hrdk_law_core.llm_client import get_llm_client
from hrdk_law_core.certs import get_qnet_certs_text

# 🌟 [모델 추상화] Gemini 직접 호출 대신 통역 창구 사용. LLM_PROVIDER로 모델 교체 가능.
_llm = None
def _client():
    global _llm
    if _llm is None:
        _llm = get_llm_client()
    return _llm

# 🌟🌟🌟 [추가된 부분 1] 링크 조립 공장 (RESTful 포맷 생성기) 🌟🌟🌟
def generate_new_law_link(law_name, enforce_date, prom_num, prom_date, article_name):
    """별표/서식인지 일반 조항인지 구분해서 법제처 RESTful 링크를 완성합니다."""
    star_match = re.search(r'(별표|서식)\s*(\d+)', article_name)
    if star_match:
        target_id = f"{star_match.group(1)}{star_match.group(2)}" # 예: 별표2
        return f"https://www.law.go.kr/법령별표서식/({law_name},{enforce_date},{target_id})"
    
    jo_match = re.search(r'(제\d+조(?:의\d+)?)', article_name)
    if jo_match:
        target_id = jo_match.group(1) # 예: 제5조
        return f"https://www.law.go.kr/법령/{law_name}/({enforce_date},{prom_num},{prom_date})/{target_id}"
    
    # 조문 매칭 실패 시 그냥 기본 법령 링크로 보냄
    return f"https://www.law.go.kr/법령/{law_name}"


def run_ai_analysis(law, attempt_count=5):
    QNET_CERTS = get_qnet_certs_text(group_by_field=True)  # 🌟 코어 단일 출처에서 종목 로드
    prompt = f"""
    당신은 '한국산업인력공단(HRDK)'의 국가기술자격 정책 수석 연구원입니다.
    아래 [최신 법령 원본]을 읽고, [국가기술자격 491개 종목 사전] 중 어떤 종목에 영향을 미치는지 분석하십시오.

    [국가기술자격 491개 종목 사전]
    {QNET_CERTS}

    [최신 법령 제·개정 내용 (마크다운)]
    {law['원본']}

    [🚨 수석 연구원 분석 지침 (핵심 위주)]
    1. 억지로 연관성을 찾지 마십시오. 법령 내용과 직접적으로 연관이 있거나, 실무상 명백히 영향을 받는 자격증만 추출하십시오.
    2. 연관성이 없으면 "종목": "", "분류": "일반", "검토필요": "X" 로 빠르게 결론 내리십시오.
    3. 모든 텍스트(요약, 분석, 사유 등)는 조문별로 핵심만 5문장 이내로 아주 간결하게 작성하십시오.

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

    [🚨 치명적 에러 방지 규칙 - 반드시 지킬 것 (시스템 붕괴 방지)]
    1. **JSON 형태 유지**: JSON의 Key(예: "요약", "종목")는 반드시 큰따옴표(")를 사용해야 합니다.
    2. **내부 텍스트 큰따옴표 금지**: 단, 당신이 작성하는 한국어 내용 안에서 단어를 강조할 때는 절대 큰따옴표(")를 쓰지 말고, 무조건 작은따옴표(')를 쓰세요.
    3. **엔터키(줄바꿈) 절대 금지**: 모든 텍스트는 중간에 줄바꿈 없이 한 줄로 이어서 작성하세요.
    4. **순수 종목명만 추출 (가장 중요)**: 연관된 자격증을 나열할 때 '[안전관리]', 'ㅇ 직무분야:' 같은 분류명이나 기호를 절대 쓰지 마십시오. 오직 '자격증이름1, 자격증이름2, 자격증이름3' 처럼 자격증 이름만 쉼표로 연결해서 작성하세요.
    5. **모든 자격증 무제한 추출**: 3번의 규칙을 지키면서, 연관 자격증이 수백 개라도 단 하나도 누락하지 말고 모두 끝까지 다 작성하세요.

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

    try:
        raw_text = _client().generate_with_retry(
            prompt, attempt_count=attempt_count,
            max_output_tokens=32768, temperature=0.1,
        )
    except Exception as e:
        return False, "", {"error": str(e)}

    try:

            match = re.search(r'```json\s*(.*?)\s*```', raw_text, re.DOTALL | re.IGNORECASE)
            if match:
                json_str = match.group(1)
            else:
                json_str = raw_text.replace("```json", "").replace("```JSON", "").replace("```", "").strip()
            
            json_str = json_str.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
            
            try:
                data = json.loads(json_str, strict=False)
            except json.JSONDecodeError as je:
                print(f"\n    🚨 [AI 문법 파괴 발생! 범인 색출 블랙박스 로그]")
                print(f"    >> AI가 뱉은 날것의 텍스트:\n{json_str}\n")
                return False, "", {"error": f"JSON 문법 오류: {je}"}

            jomun_list = data.get("조문리스트", [])
            if not jomun_list or not isinstance(jomun_list, list):
                jomun_list = [{"조문명": "내용 확인", "숫자": ""}]
                
            links_str_list = []
            names_str_list = []
            
            # 🌟🌟🌟 [추가된 부분 2] 링크를 조립해서 리스트에 넣는 로직 🌟🌟🌟
            for j in jomun_list:
                j_name = j.get("조문명", "확인불가")
                if "별표" in j_name:
                    j_name = re.sub(r'별표\s*(\d+)', r'별표 \1', j_name)
                
                if j_name == "내용 확인":
                    names_str_list.append("전체 (세부 조문 미지정)")
                    links_str_list.append(f"▶ {law['법령명']}\n{law['링크']}")
                else:
                    names_str_list.append(j_name)
                    
                    # law_api.py에서 주머니에 넣어둔 재료(공포번호 등)를 꺼내서 링크 완성!
                    new_link = generate_new_law_link(
                        law_name=law.get('법령명', ''),
                        enforce_date=law.get('시행일자', ''),
                        prom_num=law.get('공포번호', ''),
                        prom_date=law.get('공포일자', ''),
                        article_name=j_name
                    )
                    links_str_list.append(f"▶ {law['법령명']} {j_name}\n{new_link}")
            # 🌟🌟🌟 (여기까지 변경됨) 🌟🌟🌟
                
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
        return False, "", {"error": str(e)}
