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
# 1. 구글 시트 적재 함수
# ==========================================
def upload_to_google_sheet(law_info_list, target_date=TARGET_DATE):
    """[V29 핵심 수정] GCP_SA_JSON 문자열을 안전하게 읽어 구글 시트에 적재합니다."""
    if not GCP_SA_JSON or not GOOGLE_SHEET_ID:
        print("  ⚠️ 구글 시트 키 또는 ID가 없어 적재를 건너뜁니다.")
        return False

    if not law_info_list:
        return True # 적재할 데이터가 없으면 그냥 넘어감

    try:
        # 🔥 여기서 아까 났던 에러(JSON 파싱 에러)를 완벽 차단합니다!
        creds_dict = json.loads(GCP_SA_JSON.strip(), strict=False)
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        # 시트 이름이 다르면 여기를 수정하세요 (기본값: 시트1)
        sheet = client.open_by_key(GOOGLE_SHEET_ID).worksheet("시트1") 
        
        # 여러 건의 데이터를 한 번에 적재 (API 호출을 줄여 속도 대폭 향상!)
        rows_to_insert = []
        for info in law_info_list:
            row = [info.get(col, "") for col in COLUMNS]
            rows_to_insert.append(row)
            
        sheet.append_rows(rows_to_insert, value_input_option="USER_ENTERED")
        print(f"  ✅ 구글 시트 마스터 DB 적재 완료! ({len(rows_to_insert)}건)")
        return True
        
    except Exception as e:
        print(f"  ❌ 구글 시트 적재 오류: {e}")
        return False

# ==========================================
# 2. 엑셀 파일 생성 함수
# ==========================================
def create_excel_report(high_impact_laws, simple_related_laws, target_date=TARGET_DATE):
    """분석된 데이터를 바탕으로 예쁜 서식의 엑셀 파일을 만듭니다."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "연관 높은 법령"
    ws2 = wb.create_sheet(title="국가기술자격 관계 법령(단순 관련)")
    
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for ws, data_list in [(ws1, high_impact_laws), (ws2, simple_related_laws)]:
        ws.append(COLUMNS)
        # 헤더 서식 지정
        for col_num in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
        # 데이터 삽입 및 서식 지정
        for row_idx, info in enumerate(data_list, 2):
            row_data = [info.get(c, "") for c in COLUMNS]
            ws.append(row_data)
            for col_idx in range(1, len(COLUMNS) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = center_align
                cell.border = thin_border
                
        # 열 너비 조정
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20

    excel_filename = f"V29_법령모니터링_{target_date}.xlsx"
    wb.save(excel_filename)
    print(f"  💾 엑셀 파일 생성 완료: {excel_filename}")
    return excel_filename

# ==========================================
# 3. 메이크닷컴(Make.com) 웹훅 전송 함수
# ==========================================
def send_webhook_with_file(fname, total, high, simple, target_date=TARGET_DATE):
    """🔥 Make.com에 안전하게 통계와 엑셀 전송 (0건 증발 완벽 방어)"""
    if not WEBHOOK_URL:
        print("  ⚠️ 웹훅 URL이 없어 전송을 생략합니다.")
        return

    # 💡 선생님 아이디어 적용: 0이 무시되지 않게 '건'을 붙여 문자열로 전송
    summary_data = {
        "date": str(target_date), 
        "total": f"{total}건", 
        "high": f"{high}건", 
        "simple": f"{simple}건"
    }
    
    print(f"\n🚀 Make.com 전송 시도: 총 {total}건 / 연관 {high}건 / 단순 {simple}건")
    try:
        if fname and os.path.exists(fname):
            with open(fname, 'rb') as f:
                files = {'file': (os.path.basename(fname), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                response = requests.post(WEBHOOK_URL, data=summary_data, files=files)
        else:
            response = requests.post(WEBHOOK_URL, data=summary_data)
        
        if response.status_code == 200:
            print("  ✅ 웹훅 전송 성공! (Make.com 서버가 데이터를 꽉 잡았습니다!)")
        else:
            print(f"  ❌ Make.com 수신 실패! 응답코드: {response.status_code}")
            
    except Exception as e: 
        print(f"  ❌ 웹훅 네트워크 에러: {e}")