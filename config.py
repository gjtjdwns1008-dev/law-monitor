import os
from datetime import datetime, timedelta, timezone


# ==========================================
# 1. API 키 및 외부 연동 설정
# ==========================================
LAW_API_KEY = os.environ.get("LAW_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
# 🛠️ [D-3 패치] 웹훅 주소 하드코딩 제거 → GitHub Secrets(환경변수)에서 읽도록 변경
# ⚠️ 기존 주소는 저장소 이력에 노출되었으므로 Make.com에서 반드시 '재발급' 후 Secrets에 등록하세요.
WEBHOOK_URL = os.environ.get("WEBHOOK_URL")

# [V27 신규] 구글 시트 직접 제어용 환경 변수
GCP_SA_JSON = os.environ.get("GCP_SA_JSON")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")

# Phase 1 신규: SQLite 지식베이스 경로 (RADAR와 공유)
DB_PATH = os.environ.get("DB_PATH", "hrdk_law.db")

# ==========================================
# 2. 날짜 및 공통 변수 설정
# ==========================================
KST = timezone(timedelta(hours=9))
today = datetime.now(KST)

# 🌟 [D-1 로직 적용] 오늘(today)에서 1일을 뺀 어제 날짜를 계산합니다.
yesterday = today - timedelta(days=1)
TARGET_DATE = yesterday.strftime("%Y%m%d")

# 💡 만약 과거 데이터를 돌리고 싶다면 이 변수를 수동으로 바꿔서 쓰면 됩니다.
# 💡 TARGET_DATE = "20260429"
# 💡 TARGET_DATE = yesterday.strftime("%Y%m%d") # 💡 오전 5시에 돌면 어제 법령 전체를 다 가져옵니다!
# 💡 TARGET_DATE = today.strftime("%Y%m%d")

# 엑셀 및 구글 시트에 들어갈 컬럼명
COLUMNS = [
    "시행일자", "소관부처", "법령명", "개정유형", "주요 제·개정내용", 
    "법령 관련 국가기술자격 종목", "활용도 분석 구분", "활용도 분석 상세", 
    "근거 조문", "AI 신뢰도", "검토 필요", "검토 사유", "조문별 다이렉트 링크"
]

