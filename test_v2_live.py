# -*- coding: utf-8 -*-
"""v2 개조 라이브 테스트 — 실제 HWP COM으로 전 기능 검증 (스크래치 사본 사용)"""
import sys, os, shutil, tempfile
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import advanced_hwp_server as S

SRC = r"d:\Desktop\지원사업\예술경영_이교성\붙임1. [양식] 2026 예술분야 예비창업 프로그램 LBB.hwp"
WORK = os.path.join(tempfile.gettempdir(), "hwp_v2_test")
os.makedirs(WORK, exist_ok=True)
DOC = os.path.join(WORK, "test_form.hwp")
IMG = r"d:\Desktop\지원사업\예술경영_이교성\assets\acc2.png"

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  [PASS] " if cond else "  [FAIL] ") + name + (f" — {detail}" if detail else ""))


print("=== 준비: 프로세스 정리 + 사본 생성 ===")
print(S.kill_hwp_processes())
shutil.copy(SRC, DOC)

print("=== 1) open_document ===")
r = S.open_document(DOC)
check("open_document", "열었습니다" in r, r[:60])

print("=== 2) get_text_all (인코딩) ===")
t = S.get_text_all()
hangul = sum(1 for ch in t[:3000] if '가' <= ch <= '힣')
moji = sum(1 for ch in t[:3000] if 0x80 <= ord(ch) <= 0xFF)
check("한글 정상 추출", hangul > 200, f"한글 {hangul}자")
check("모지바케 없음", moji < hangul / 5, f"고위 Latin {moji}자")
check("핵심 문구 포함", "사업계획서" in t and "아이템명" in t)

print("=== 3) find_text (보고형) ===")
r = S.find_text("아이템명", show_context=True)
check("find_text 한글 검색", "발견" in r and "찾을 수 없습니다" not in r, r.split(chr(10))[0])

print("=== 4) find_text_and_move + fill_cell_after_label ===")
r = S.find_text_and_move("아이템명", occurrence=1)
check("find_text_and_move", "이동했습니다" in r, r[:70])
r = S.fill_cell_after_label("아이템명", "테스트 안전관제 서비스")
check("fill_cell_after_label", "기입했습니다" in r, r[:60])
t2 = S.get_text_all()
check("기입 반영 확인", "테스트 안전관제 서비스" in t2)

print("=== 5) format_text_occurrence (RGB) ===")
r = S.format_text_occurrence("테스트 안전관제 서비스", color="#FF0000", bold=True)
check("format_text_occurrence", "변경했습니다" in r, r[:60])

print("=== 6) find_and_replace(all) + 검증 ===")
r = S.find_and_replace("테스트 안전관제 서비스", "교체된 서비스명", replace_all=True)
check("find_and_replace all", "바꾸었습니다" in r, r[:70])
check("치환 잔존 검증", "교체된 서비스명" in S.get_text_all())

print("=== 7) delete_paragraph_containing ===")
r = S.delete_paragraph_containing("교체된 서비스명")
check("delete_paragraph_containing", "삭제했습니다" in r, r[:60])
check("삭제 반영", "교체된 서비스명" not in S.get_text_all())

print("=== 8) clear_cell_after_label ===")
S.fill_cell_after_label("아이템명", "비울 내용 A")
r = S.clear_cell_after_label("아이템명")
check("clear_cell_after_label", "비웠습니다" in r, r[:60])
check("셀 비움 반영", "비울 내용 A" not in S.get_text_all())

print("=== 9) move_to_page + get_text_by_page ===")
r = S.move_to_page(2)
check("move_to_page", "이동했습니다" in r, r[:50])
t3 = S.get_text_by_page(2)
check("get_text_by_page", len(t3) > 20 and "실패" not in t3[:20], f"{len(t3)}자")

print("=== 10) create_table (폭·데이터·탈출) ===")
S.move_to_document_end()
hwp = S.hwp_controller.hwp
hwp.HAction.Run("BreakPara")
r = S.create_table(3, 3, width_mm=120,
                   data_json='[["구분","항목","값"],["A","가나다","1"],["B","라마바","2"]]',
                   col_widths="[1,2,1]")
check("create_table", "생성했습니다" in r, r[:80])
pos = hwp.GetPos()
check("표 탈출 (커서 본문)", pos[0] == 0, f"list={pos[0]}")
S.insert_text("표 다음 문장입니다")
check("표 다음 입력 정상", "표 다음 문장입니다" in S.get_text_all())

print("=== 11) insert_image (mm) + list/resize ===")
r = S.insert_image(IMG, width=60)
check("insert_image", "삽입했습니다" in r, r[:70])
r = S.list_document_images()
check("list_document_images", "그림" in r and "mm" in r, r.split(chr(10))[0])
n_pics = r.count("#")
r = S.resize_document_image(n_pics - 1, 40)
check("resize_document_image", "->" in r and "40" in r, r[:70])

print("=== 12) render_content_blocks ===")
S.move_to_document_end()
hwp.HAction.Run("BreakPara")
blocks = ('[{"type":"h","text":"렌더 테스트 제목"},'
          '{"type":"b","text":"불릿 하나"},'
          '{"type":"t","headers":["열1","열2"],"rows":[["값1","값2"]],"width_mm":100},'
          '{"type":"note","text":"주석입니다"}]')
r = S.render_content_blocks(blocks)
check("render_content_blocks", "4개 블록" in r, r[:60])
t4 = S.get_text_all()
check("렌더 반영", "렌더 테스트 제목" in t4 and "불릿 하나" in t4)

print("=== 13) measure_layout ===")
r = S.measure_layout('["렌더 테스트 제목","아이템명"]')
check("measure_layout", "총 " in r and "페이지" in r, r.split(chr(10))[0])
check("마커 실측", "'렌더 테스트 제목': p" in r)

print("=== 14) export_to_pdf ===")
pdf = os.path.join(WORK, "test_out.pdf")
r = S.export_to_pdf(pdf)
check("export_to_pdf", os.path.exists(pdf), r[:60])

print("=== 16) 리뷰 지적 회귀 체크 ===")
r = S.find_and_replace("이_문서에_절대_없는_문자열_XYZ", "무언가", replace_all=True)
check("부재 텍스트 all-replace 거짓성공 방지", "찾을 수 없습니다" in r, r[:50])
S.find_text_and_move("표 다음 문장입니다", select=True)
r = S.apply_font_format(font_size=12, underline=True, color="pink")
check("apply_font_format(밑줄+미지의 색)", "적용했습니다" in r, r[:60])
r = S.find_and_replace("표 다음 문장입니다", "표 뒤 문장으로 교체", replace_all=False)
check("단건 치환(새 경로)", "바꾸었습니다" in r, r[:50])
check("단건 치환 반영", "표 뒤 문장으로 교체" in S.get_text_all())
r = S.set_page_margins(20, 20, 18, 18)
check("set_page_margins 무예외", "설정했습니다" in r, r[:50])
r = S.find_and_replace("ai테스트", "AI테스트", replace_all=True)  # 대소문자 치환 오판 방지 (부재해도 OK)
check("대소문자 all-replace 안전", "실패" not in r, r[:50])

print("=== 15) 저장·종료 ===")
r = S.save_document()
check("save_document", "저장했습니다" in r)
r = S.quit_hwp()
check("quit_hwp", "종료했습니다" in r, r[:50])

print()
print(f"결과: PASS {len(PASS)} / FAIL {len(FAIL)}")
if FAIL:
    print("실패 목록:", FAIL)
    sys.exit(1)
print("ALL GREEN")
