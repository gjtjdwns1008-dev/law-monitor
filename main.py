"""
law-monitor - main.py (백필 구조판)
-----------------------------------
HRDK-LAW-RADAR와 동일한 흐름 구조. 차이점: 워크넷 없음 / 활용도 분석 / high·simple 2분류.

🌟 백필(Backfill) 전략:
  법제처 IP 차단으로 며칠 건너뛰어도, 연결되는 날 밀린 날짜를 모두 따라잡습니다.
  - 시작 시 연결 확인 → 안 되면 즉시 종료 (재시도로 시간 낭비 안 함)
  - 마지막 성공일 다음날 ~ 어제까지의 모든 날짜를 과거→현재 순으로 처리
"""

import time
import sys
import os

from config import LAW_API_KEY, DB_PATH

from hrdk_law_core.scraper  import get_base_laws
from hrdk_law_core.certs    import get_relevant_certs_text
from hrdk_law_core.db       import KnowledgeBase
from hrdk_law_core.backfill import check_law_reachable, pending_dates, mark_done, is_valid_target_date

from brain_gemini import run_ai_analysis
from report_maker import upload_to_google_sheet, create_excel_report, send_webhook_with_file


def process_one_day(target_date: str, kb, run_note: str = "") -> bool:
    """
    하루치(target_date) 법령을 수집·분석·저장·보고합니다.
    반환: 처리 성공 여부 (수집 실패 시 False → 그날은 다음 기회에 재시도).
    run_note: 수동 실행 시 로그 접두어 (예: '[수동 6/17 실행] ').
    """
    print(f"\n{'='*50}\n📅 [{target_date}] 처리 시작\n{'='*50}")

    laws = get_base_laws(api_key=LAW_API_KEY, target_date=target_date)

    if laws is None:
        print(f"  ❌ [{target_date}] 법제처 수집 실패 (이 날짜는 다음 기회에 재시도)")
        return False

    if not laws:
        print(f"  ℹ️ [{target_date}] 시행 법령 없음 (0건)")
        upload_to_google_sheet(0, [], [], target_date=target_date,
            status="🟢 정상 작동 (공포 법령 없음)",
            log=f"{run_note}새로 시행되는 국가 법령이 없습니다.")
        return True  # 0건도 '처리 완료'로 간주 (밀린 목록에서 제거)

    high_impact_laws, simple_related_laws = [], []
    failed_queue, all_results = [], []

    print(f"\n🏎️  총 {len(laws)}건 분석 시작...")
    for idx, law in enumerate(laws):
        print(f"  [{idx+1}/{len(laws)}] 🔍 {law['법령명']}")
        t0 = time.time()

        # 🌟 [버리지 않는 체] 스킵 법령은 사유와 함께 DB에 보존
        if law.get("스킵여부"):
            hold_reason = law.get("스킵사유", "조직/직제 관련")
            print(f"    ⏩ [보류: {hold_reason}]")
            try:
                kb.add_held_law(
                    law_name=law["법령명"], enforce_date=law.get("시행일자", ""),
                    ministry=law.get("소관부처", ""), hold_reason=hold_reason,
                    law_link=law.get("링크", ""),
                )
            except Exception as he:
                print(f"      ⚠️ 보류 로그 기록 실패: {he}")
            continue

        certs_text = get_relevant_certs_text(law.get("원본", ""), group_by_field=True)
        success, classification, law_info = run_ai_analysis(law, certs_text)
        elapsed = time.time() - t0

        if success:
            if classification == "연관높음":
                high_impact_laws.append(law_info); print(f"    🌟 연관높음 ({elapsed:.1f}초)")
            elif classification == "단순관련":
                simple_related_laws.append(law_info); print(f"    🔹 단순관련 ({elapsed:.1f}초)")
            else:
                print(f"    ❌ 해당없음 ({elapsed:.1f}초)")
            all_results.append(law_info)
        else:
            law["error_msg"] = law_info.get("error", "알 수 없음")
            failed_queue.append(law)
            print(f"    ⏩ [분석 실패: {law['error_msg']}] ({elapsed:.1f}초)")

    # 패자부활전 (AI 분석 실패분)
    if failed_queue:
        print(f"\n🚑 패자부활전 {len(failed_queue)}건... (20초 대기)")
        time.sleep(20)
        for law in failed_queue:
            print(f"  [재시도] {law['법령명']}... ", end="", flush=True)
            certs_text = get_relevant_certs_text(law.get("원본", ""), group_by_field=True)
            success, classification, law_info = run_ai_analysis(law, certs_text, attempt_count=3)
            if success:
                if classification == "연관높음":
                    high_impact_laws.append(law_info); print("🌟 연관높음")
                elif classification == "단순관련":
                    simple_related_laws.append(law_info); print("🔹 단순관련")
                else:
                    print("❌ 해당없음")
                all_results.append(law_info)
            else:
                print(f"💀 [최종 실패] {law.get('error_msg', '크레딧 소진')}")

    # 저장 & 보고
    print("\n📝 구글 시트 적재...")
    log_text = (f"{run_note}총 {len(laws)}건 중 "
                f"연관높음 {len(high_impact_laws)}건, 단순관련 {len(simple_related_laws)}건")
    upload_to_google_sheet(len(laws), high_impact_laws, simple_related_laws,
                           target_date=target_date, status="🟢 정상 작동", log=log_text)
    print("📊 엑셀 보고서 생성...")
    excel_filename = create_excel_report(high_impact_laws, simple_related_laws, target_date=target_date)
    print("🚀 웹훅 전송...")
    send_webhook_with_file(excel_filename, len(laws), len(high_impact_laws),
                           len(simple_related_laws), target_date=target_date)
    return True


def main():
    print("🚀 [law-monitor] 시작\n" + "=" * 50)
    start_time = time.time()

    kb = KnowledgeBase(DB_PATH)
    print(f"📚 지식베이스 로드 완료 ({DB_PATH})")

    # ── [수동 실행 모드] 특정 일자만 처리 (연결 확인보다 먼저 — 대상 날짜를 알아야 함) ──
    manual_date = os.environ.get("MANUAL_DATE", "").strip()
    if manual_date:
        from datetime import datetime, timezone, timedelta
        run_day = datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d")
        if not is_valid_target_date(manual_date):
            print(f"❌ 잘못된 날짜: '{manual_date}'. YYYYMMDD 형식의 과거(또는 오늘) 날짜여야 합니다.")
            sys.exit(1)
        print(f"🔧 [수동 실행] {manual_date} 한 날짜만 처리합니다. (자동 백필 상태는 변경하지 않음)")
        # 수동 실행도 연결 확인 — 단, 실패해도 '대상 날짜(manual_date)' 행에 기록
        if not check_law_reachable(LAW_API_KEY):
            print(f"❌ [수동 실행] 법제처 연결 불가. {manual_date} 처리 실패.")
            try:
                upload_to_google_sheet(0, [], [], target_date=manual_date,
                    status="🔴 법제처 연결 불가 (IP 차단 추정)",
                    log="[수동 실행] 법제처 연결 실패. 되는 날 재시도 필요.")
            except Exception:
                pass
            sys.exit(1)
        ok = process_one_day(manual_date, kb, run_note="[수동 실행] ")
        # ⚠️ mark_done 호출하지 않음 — 수동 실행이 자동 백필을 꼬이게 하면 안 됨
        print(f"\n🎉 [수동 실행 종료] {manual_date} 처리 {'성공' if ok else '실패'}")
        if not ok:
            sys.exit(1)
        return

    # ── 1. 오늘이 '되는 날'인지 확인 (자동 실행) ──────────
    if not check_law_reachable(LAW_API_KEY):
        print("❌ 법제처 연결 불가 (오늘은 IP 차단일로 판단). 재시도 없이 종료합니다.")
        print("   → 밀린 날짜는 연결되는 다음 날 자동으로 따라잡습니다.")
        # 🌟 연결 실패도 총괄현황표에 기록 (통신 이력이 남도록)
        from datetime import datetime, timedelta, timezone
        kst_today = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d")
        try:
            upload_to_google_sheet(0, [], [], target_date=kst_today,
                status="🔴 법제처 연결 불가 (IP 차단 추정)",
                log="오늘 연결 실패. 밀린 날짜는 다음 연결일에 백필됩니다.")
        except Exception:
            pass
        sys.exit(1)
    print("✅ 법제처 연결 확인됨. 처리 시작.")

    # ── 2. 밀린 날짜 목록 계산 (마지막 성공일+1 ~ 어제) ──
    dates = pending_dates(kb)
    if not dates:
        print("ℹ️ 처리할 밀린 날짜가 없습니다 (이미 최신).")
        return
    print(f"📋 처리 대상 날짜 {len(dates)}일: {dates[0]} ~ {dates[-1]}")
    if len(dates) > 10:
        print(f"   ⚠️ 밀린 날짜가 {len(dates)}일로 많습니다. 순서대로 모두 처리합니다.")

    # ── 3. 과거→현재 순으로 따라잡기 ──────────────────────
    done, failed = 0, 0
    for d in dates:
        try:
            if process_one_day(d, kb):
                mark_done(kb, d)   # 성공 시 마지막 성공일 갱신
                done += 1
            else:
                failed += 1
                print(f"  ⏸️ [{d}] 수집 실패로 백필 중단. 다음 실행에서 이어서 처리합니다.")
                break
        except Exception as e:
            print(f"  💥 [{d}] 처리 중 오류: {e}")
            failed += 1
            break

    elapsed_total = time.time() - start_time
    print(f"\n🎉 [종료] 완료 {done}일 / 실패 {failed}일 (소요: {elapsed_total/60:.1f}분)")
    if failed and not done:
        sys.exit(1)


if __name__ == "__main__":
    main()
