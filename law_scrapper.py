# law-monitor 호환 래퍼
# main.py는 get_base_laws()를 인자 없이 호출합니다.
# 공유 코어는 api_key, target_date를 인자로 받으므로 여기서 연결해줍니다.
from config import LAW_API_KEY, TARGET_DATE
from hrdk_law_core.scraper import get_base_laws as _core_get_base_laws

def get_base_laws(target_date=TARGET_DATE):
    return _core_get_base_laws(api_key=LAW_API_KEY, target_date=target_date)
