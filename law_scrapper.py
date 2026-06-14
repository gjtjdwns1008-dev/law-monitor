"""
law_scrapper.py (호환 래퍼)
---------------------------
법령 수집 로직은 공유 코어(hrdk_law_core.scraper)로 이전되었습니다.
이 파일은 기존 main.py가 get_base_laws()를 인자 없이 호출하던 방식을
그대로 유지하기 위한 얇은 연결층(wrapper)입니다.

main.py는 수정할 필요 없이 그대로 동작합니다.
"""

from config import LAW_API_KEY, TARGET_DATE
from hrdk_law_core.scraper import get_base_laws as _core_get_base_laws


def get_base_laws(target_date=TARGET_DATE):
    """공유 코어의 수집 함수에 config의 API 키와 날짜를 자동으로 넘겨줍니다."""
    return _core_get_base_laws(api_key=LAW_API_KEY, target_date=target_date)
