"""
law-monitor - main.py (구조 통일판)
-----------------------------------
HRDK-LAW-RADAR와 동일한 흐름 구조를 따릅니다.
차이점(법-monitor 고유): 워크넷 매쉬업 없음 / 활용도 분석 / high·simple 2분류.

파일 역할 규칙 (RADAR와 동일):
  main.py        - 분석 흐름 전체 조율 (수집→종목준비→AI→저장→보고)
  brain_gemini.py- AI 프롬프트 + 응답 파싱만
  config.py      - 환경변수·컬럼 정의
  report_maker.py- 구글시트·엑셀·웹훅 출력
"""

import time
import sys

from config import TARGET_DATE, LAW_API_KEY, DB_PATH

# ── 공유 코어 임포트 (RADAR와 동일) ──────────────────────
from hrdk_law_core.scraper import get_base_laws
from hrdk_law_core.certs   import get_relevant_certs_text
from hrdk_law_core.db      import KnowledgeBase

# ── 레포 고유 모듈 ───────────────────────────────────────
from brain_gemini import run_ai_analysis
from report_maker import upload_to_google_sheet, create_excel_report, send_webhook_with_file


def main():
    print(f"🚀 [law-monitor] {TARGET_DATE} 데이터 수집 및 분석 시작...\n" + "=" * 50)
    start_time = time.time()

    # ── 지식베이스 로드 (보류 로그 기록용) ───────────────
    kb = KnowledgeBase(DB_PATH)
    print(f"📚 지식베이스 로드 완료 ({DB_PATH})")

    try:
        # ═══════════════════════════════════════════════════
        # 1. 법령 수집 (공유 코어)
        # ═══════════════════════════════════════════════════
        laws = get_base_laws(api_key=LAW_API_KEY, target_date=TARGET_DATE)

        if laws is None:
            err_msg = "법제처 서버 연결 완전 실패. 안전하게 종료합니다."
            print(f"❌ [결정적 오류] {err_msg}")
            upload_to_google_sheet(0, [], [])  # 총괄현황표에 에러 기록
            sys.exit(1)  # 종료코드 1 → 워크플로우 재시도 트리거

        if not laws:
            print(f"  ℹ️ {TARGET_DATE} 시행 법령 없음. (0건 처리)")
            upload_to_google_sheet(0, [], [])
            send_webhook_with_file(create_excel_report([], []), 0, 0, 0)
            return

        high_impact_laws, simple_related_laws = [], []
        failed_queue, all_results = [], []

        # ═══════════════════════════════════════════════════
        # 2. AI 정밀 분석 루프
        # ═══════════════════════════════════════════════════
        print(f"\n🏎️  총 {len(laws)}건 분석 시작...")
        for idx, law in enumerate(laws):
            print(f"  [{idx+1}/{len(laws)}] 🔍 {law['법령명']}")
            t0 = time.time()

            # 🌟 [버리지 않는 체] 스킵 법령은 사유와 함께 DB에 보존
            if law.get("스킵여부"):
                hold_reason = law.get("스킵사유", "조직/직제 관련")
                print(f"    ⏩ [보류: {hold_reason}] (삭제 아님, 보류 로그에 기록)")
                try:
                    kb.add_held_law(
                        law_name=law["법령명"],
                        enforce_date=law.get("시행일자", ""),
                        ministry=law.get("소관부처", ""),
                        hold_reason=hold_reason,
                        law_link=law.get("링크", ""),
                    )
                except Exception as he:
                    print(f"      ⚠️ 보류 로그 기록 실패: {he}")
                continue

            # 🌟 종목 슬림화 (누락 시 전체 폴백) — main에서 담당 (RADAR와 동일)
            certs_text = get_relevant_certs_text(law.get("원본", ""), group_by_field=True)
            success, classification, law_info = run_ai_analysis(law, certs_text)
            elapsed = time.time() - t0

            if success:
                if classification == "연관높음":
                    high_impact_laws.append(law_info)
                    print(f"    🌟 연관높음 ({elapsed:.1f}초)")
                elif classification == "단순관련":
                    simple_related_laws.append(law_info)
                    print(f"    🔹 단순관련 ({elapsed:.1f}초)")
                else:
                    print(f"    ❌ 해당없음 ({elapsed:.1f}초)")
                all_results.append(law_info)
            else:
                law["error_msg"] = law_info.get("error", "알 수 없음")
                failed_queue.append(law)
                print(f"    ⏩ [분석 실패: {law['error_msg']}] ({elapsed:.1f}초)")

        # ═══════════════════════════════════════════════════
        # 3. 패자부활전
        # ═══════════════════════════════════════════════════
        if failed_queue:
            print(f"\n🚑 패자부활전 {len(failed_queue)}건 시작... (20초 대기)")
            time.sleep(20)
            for law in failed_queue:
                print(f"  [재시도] {law['법령명']}... ", end="", flush=True)
                certs_text = get_relevant_certs_text(law.get("원본", ""), group_by_field=True)
                success, classification, law_info = run_ai_analysis(law, certs_text, attempt_count=3)
                if success:
                    if classification == "연관높음":
                        high_impact_laws.append(law_info)
                        print("🌟 연관높음")
                    elif classification == "단순관련":
                        simple_related_laws.append(law_info)
                        print("🔹 단순관련")
                    else:
                        print("❌ 해당없음")
                    all_results.append(law_info)
                else:
                    final_err = law.get("error_msg", "Gemini 크레딧 소진")
                    print(f"💀 [최종 실패] {final_err}")

        # ═══════════════════════════════════════════════════
        # 4. 구글 시트 & 보고서
        # ═══════════════════════════════════════════════════
        print("\n📝 구글 시트 마스터 DB 적재 시작...")
        upload_to_google_sheet(len(laws), high_impact_laws, simple_related_laws)

        print("\n📊 보고용 엑셀 파일 생성 중...")
        excel_filename = create_excel_report(high_impact_laws, simple_related_laws)

        print("\n🚀 Make.com 웹훅 전송 시작...")
        send_webhook_with_file(
            excel_filename, len(laws), len(high_impact_laws), len(simple_related_laws)
        )

        elapsed_total = time.time() - start_time
        print(f"\n🎉 [종료] 완료! (소요 시간: {elapsed_total / 60:.1f}분)")

    except Exception as e:
        fatal_msg = f"런타임 에러: {str(e)}"
        print(f"\n💥 [치명적 오류] {fatal_msg}")
        try:
            upload_to_google_sheet(0, [], [])
        except Exception as sheet_err:
            print(f"  ❌ 구글 시트 로깅도 실패: {sheet_err}")
        sys.exit(1)


if __name__ == "__main__":
    main()
