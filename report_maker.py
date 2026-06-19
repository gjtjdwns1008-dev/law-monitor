import os
import json
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# 💡 1단계 config 파일에서 설정값들을 가져옵니다.
from config import COLUMNS, WEBHOOK_URL, GCP_SA_JSON, GOOGLE_SHEET_ID, TARGET_DATE

# ==========================================
# 1. 구글 시트 마스터 DB 적재 (V28 분류 로직 적용)
# ==========================================
def upload_to_google_sheet(total_len, high_list, simple_list, target_date=TARGET_DATE,
                           status="🟢 정상 작동", log=""):
    """[V29 최종] 3개 시트에 데이터를 분류하여 적재합니다.
    status/log: 총괄현황표에 통신·처리 상태를 함께 기록 (백필 추적용)."""
    if not GCP_SA_JSON or not GOOGLE_SHEET_ID:
        print("  ⚠️ 구글 시트 설정 정보가 없어 적재를 건너뜁니다.")
        return

    try:
        # JSON 파싱 및 인증
        creds_dict = json.loads(GCP_SA_JSON.strip(), strict=False)
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)

        # 🌟 날짜 가독성 변환 (20260428 -> 2026년_04월_28일)
        display_date = f"{target_date[:4]}년_{target_date[4:6]}월_{target_date[6:]}일"

        # 1) 총괄현황표 기록 [옵션 B] — 같은 날짜 행에 시도 이력 누적
        #    상태 칸 예: "04:13🔴 → 08:47🔴 → 12:31🟢" (한 줄로 그날 이력 전체 확인)
        try:
            from hrdk_law_core.sheets import upsert_daily_summary_row
            # status 문자열에서 심볼만 추출 (🟢/🔴/🟡)
            symbol = "🟢"
            for s in ("🔴", "🟡", "🟢"):
                if s in status:
                    symbol = s
                    break
            upsert_daily_summary_row(
                spreadsheet,
                sheet_name="총괄현황표",
                target_date_display=display_date,
                cols_before_status=[display_date, total_len, len(high_list), len(simple_list)],
                status_symbol=symbol,
                log=log,
            )
        except Exception as se:
            print(f"  ⚠️ '총괄현황표' 시트 기록 실패: {se}")

        # 데이터 변환 함수
        def prepare_rows(laws):
            return [[law.get(col, "") for col in COLUMNS] for law in laws]

        # 2) 연관 높은 법령 적재
        if high_list:
            try:
                ws_high = spreadsheet.worksheet("연관 높은 법령")
                ws_high.append_rows(prepare_rows(high_list), value_input_option="USER_ENTERED")
                print(f"  🔥 연관 높은 법령 {len(high_list)}건 적재 완료")
            except: print("  ⚠️ '연관 높은 법령' 시트를 찾을 수 없습니다.")

        # 3) 단순 관련 법령 적재
        if simple_list:
            try:
                ws_simple = spreadsheet.worksheet("국가기술자격 관계 법령(단순 관련)")
                ws_simple.append_rows(prepare_rows(simple_list), value_input_option="USER_ENTERED")
                print(f"  🟡 단순 관련 법령 {len(simple_list)}건 적재 완료")
            except: print("  ⚠️ '국가기술자격 관계 법령(단순 관련)' 시트를 찾을 수 없습니다.")

    except Exception as e:
        print(f"  ❌ 구글 시트 적재 중 치명적 오류: {e}")

# ==========================================
# 2. 엑셀 파일 생성 함수 (기존과 동일)
# ==========================================
def create_excel_report(high_impact_laws, simple_related_laws, target_date=TARGET_DATE):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "연관 높은 법령"
    ws2 = wb.create_sheet(title="국가기술자격 관계 법령(단순 관련)")
    
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for ws, data_list in [(ws1, high_impact_laws), (ws2, simple_related_laws)]:
        ws.append(COLUMNS)
        for row_idx, info in enumerate(data_list, 2):
            ws.append([info.get(c, "") for c in COLUMNS])
        # 기본 열 너비 설정
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20

    excel_filename = f"V29_법령모니터링_{target_date}.xlsx"
    wb.save(excel_filename)
    return excel_filename

# ==========================================
# 3. 메이크닷컴 웹훅 전송 (기존과 동일)
# ==========================================
def send_webhook_with_file(fname, total, high, simple, target_date=TARGET_DATE):
    if not WEBHOOK_URL: return
    # 🌟 [근본 원인 해결!] 메일/웹훅으로 보낼 때도 사람이 읽기 편한 날짜로 변환해서 쏩니다!
    display_date = f"{target_date[:4]}년 {target_date[4:6]}월 {target_date[6:]}일"
    
    # 이제 Make.com은 "20260428"이 아니라 "2026년 04월 28일" 이라는 데이터를 받게 됩니다!
    # 🏷️ system/source: 두 시스템(RADAR/monitor)을 구분하는 식별값 (메일 제목 분기용)
    summary_data = {
        "system": "law-monitor",
        "source": "monitor",
        "subject": f"[law-monitor] {display_date} 개정법령 활용도 모니터링 (연관 {high}건)",
        "date": display_date, "total": f"{total}건", "high": f"{high}건", "simple": f"{simple}건"
    }
    try:
        if fname and os.path.exists(fname):
            with open(fname, 'rb') as f:
                requests.post(WEBHOOK_URL, data=summary_data, files={'file': (os.path.basename(fname), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')})
        else:
            requests.post(WEBHOOK_URL, data=summary_data)
        print("  ✅ 웹훅 전송 성공!")
    except Exception as e: print(f"  ❌ 웹훅 에러: {e}")
