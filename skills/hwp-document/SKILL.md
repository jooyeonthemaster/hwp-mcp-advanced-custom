---
name: hwp-document
description: 한컴 한글(HWP·HWPX) 문서 작업을 한 흐름으로 처리한다 — 정부지원사업 신청서·중간보고서·결과보고서 양식 채우기, 한글 문서 새로 쓰기·편집·읽기·PDF 변환. 한컴 MCP(mcp__hwp__*)의 안전한 사용 순서와 함정 회피, 양식 표 빈칸은 HWPX XML 수술(scripts/hwpx_fill.py)로 채우기, AI가 쓴 서술형 본문은 im-not-ai(humanize-korean)로 AI 티를 벗긴 뒤 넣기, 한글에서 다시 열어 검증하기까지 다룬다. "한글 파일", "hwp", "hwpx", "한컴", "양식 채워줘", "신청서 작성", "보고서 한글로", "사업계획서", "중간보고서", "빈칸 채우기", "한글 문서 읽어줘" 요청에 사용.
---

# 한컴 문서 작업 (hwp-document)

한글 문서 작업은 세 도구를 나눠 쓴다.

| 도구 | 맡는 일 | 약점 |
|---|---|---|
| **한컴 MCP** (`mcp__hwp__*`, 76개) | 열기·읽기·문구 치환·새 문서 조판·PDF 내보내기·검증 | 병합 셀 표 이동이 불안정, 경고창이 뜨면 30분 멈춤 |
| **HWPX 수술** (`scripts/hwpx_fill.py`) | 양식 표 빈칸 채우기, 대량 치환, 안내 상자 지우기, 문단 복제 | .hwpx만 됨, 한글이 없어도 되지만 결과는 한글로 열어 봐야 함 |
| **im-not-ai** (`humanize-korean` 플러그인) | AI가 쓴 서술형 문단의 번역투·기계적 대구·당위 남발 걷어내기 | 내용은 안 바꾼다 — 사실 오류는 못 잡음 |

## 0. 길 고르기

| 요청 | 길 |
|---|---|
| 문서 읽기·요약·표 추출 | MCP: `open_document` → `get_text_all` / `get_document_structure` / `get_table_as_csv` |
| 문구 몇 군데 고치기 | MCP: `find_and_replace` / `batch_replace` |
| **양식 표 빈칸 채우기** (신청서·보고서) | **HWPX 수술** (§3). MCP `fill_cell_after_label` 은 hops=1(바로 오른쪽 칸)만 |
| 새 문서 만들기 | MCP: `create_document` → `render_content_blocks` (제목·불릿·표·이미지 JSON) |
| 분량 제한(예: 10쪽 이내) 확인 | MCP: `export_to_pdf` → `measure_layout` |

**.hwp 양식을 수술해야 할 때**: HWPX 수술은 .hwpx만 다룬다. 사용자에게 한글에서 「다른 이름으로 저장 → 파일 형식 HWPX」로 바꿔 달라고 요청한다. MCP `save_document` 는 확장자와 상관없이 **항상 HWP 형식으로 저장**하므로 이름만 `.hwpx` 로 바꿔도 HWPX가 되지 않는다.

## 1. 철칙

1. **원본 양식은 절대 덮어쓰지 않는다.** 결과는 항상 새 파일명으로 저장한다. `hwpx_fill.py` 도 원본 경로 저장을 거부한다.
2. **"저장했습니다"를 믿지 않는다.** MCP `save_document` 는 실제로 저장하지 못했는데 성공 메시지를 돌려준 적이 있다. 저장 뒤에는 반드시 파일이 있는지, 수정 시각이 방금인지 확인한다.
   ```powershell
   Get-Item -LiteralPath "D:\경로\[별첨] 결과.hwpx" | Select-Object FullName, Length, LastWriteTime
   ```
   대괄호(`[별첨]`)가 든 파일명은 `-LiteralPath` 가 필수다. 빼면 와일드카드로 해석돼 "없음"으로 나온다.
3. **사실은 사용자 자료에서만 가져온다.** 매출·인원·날짜·특허 건수·연락처를 지어내지 않는다. 모르는 칸은 비워 두고 목록으로 보고한다.
4. **실데이터는 공개 저장소에 커밋하지 않는다.** 채우기용 spec.json·빌드 스크립트에는 연락처·매출이 들어간다. 작업 폴더(예: `D:\Desktop\지원사업\...\tools\`)에 두고, 이 저장소에는 올리지 않는다.
5. **성공을 단정하지 않는다.** 마지막에 "한글에서 직접 열어 확인해 달라"고 요청한다.

## 2. 한컴 MCP 안전한 사용 순서

```
1) set_automation_mode(true)      ← 파일을 열기 "전에". 경고창 자동 처리
2) list_all_hwp_windows()          ← 이미 열린 한글 창 확인
3) open_document("실제 전체 경로")  ← 빈 문서가 아니라 실제 파일
4) 작업 (읽기·치환·조판)
5) save_document("새 경로")         ← 그리고 §1-2 파일 확인
```

**함정과 대응** (2026-08 실제 작업에서 확인)

| 함정 | 증상 | 대응 |
|---|---|---|
| 경고창 블로킹 | 손상·변조 경고 같은 모달이 뜨면 COM 호출이 무한 대기, MCP 전체가 1800초 타임아웃 | 열기 전 `set_automation_mode(true)`. 이미 멈췄으면 PowerShell `Stop-Process -Name Hwp -Force` (MCP `kill_hwp_processes` 도 같이 멈춰 있어 안 통함) |
| 셀 이동 오작동 | `fill_cell_after_label` hops>1 이 병합 셀에서 라벨 칸에 착지해 기존 내용 파괴 | hops=1만 쓴다. 그 외 표 채우기는 §3 HWPX 수술 |
| 침묵 저장 실패 | `save_document(경로)` 가 성공을 돌려주고 파일이 없음 | §1-2 파일 확인 |
| 타임아웃 후 상태 불명 | 타임아웃이 났는데 명령은 이미 적용돼 있음 | 다시 실행하기 전에 `get_text_all` 로 현재 상태부터 확인. 안 하면 두 번 적용된다 |
| 특수문자 검색 실패 | `ⓐ`·`※` 같은 문자가 든 라벨을 `find_text` 가 못 찾음 | 특수문자를 뺀 부분 문자열로 검색 |
| 여러 한글 프로세스 | COM은 한 프로세스에만 붙는다 | 대상 파일을 `open_document` 로 MCP 쪽 인스턴스에서 연다 |

## 3. HWPX 수술로 양식 채우기

HWPX는 ZIP 안의 XML이다. XML을 다시 직렬화하지 않고 문자열 그대로 고쳐서 양식의 서식을 보존한다. 이 방법은 2026-08 초기관광벤처 중간보고서 작업에서 검증했다. `hwpx_fill.py` 는 그때 쓴 스크립트와 바이트 단위로 같은 결과를 낸다.

### 3-1. 살펴보기

```bash
python scripts/hwpx_fill.py inspect "양식.hwpx"
python scripts/hwpx_fill.py inspect "양식.hwpx" --grep "삭제 후 제출"
```

- **빈 run 목록**: 빈칸은 `<hp:run charPrIDRef="N"/>` 처럼 글자 없이 닫힌 run이다. 번호마다 바로 앞에 보이는 글자가 나온다. `[1] …중간보고서 기업명` 이면 1번이 기업명 칸이다.
- **`--grep`**: 문구가 몇 번 나오는지 센다. 치환의 `expect` 와 `nth` 를 정할 때 쓴다.
- 섹션이 여럿이면(`Contents/section1.xml` …) `--section` 으로 하나씩 본다.

### 3-2. spec.json 쓰기 (작업 폴더에)

```json
{
  "section": "Contents/section0.xml",
  "ops": [
    {"op": "replace", "old": "제00회", "new": "제18회"},
    {"op": "replace", "old": "<hp:t>1건</hp:t>", "new": "<hp:t>진행 중</hp:t>", "nth": 3},
    {"op": "replace_text", "old": "예시) 홈페이지 구축", "new": "유료 고객 확보"},
    {"op": "fill_empty", "cells": [[1, "회사명"], [2, "대표자"], [38, "40,000,000원"]]},
    {"op": "clone_after", "anchor": "    - 기존 항목", "lines": ["    - 추가 항목 1", "    - 추가 항목 2"]},
    {"op": "remove_tables", "marker": "삭제 후 제출", "forbid": ["협약시 계량목표"]}
  ]
}
```

| op | 하는 일 | 주의 |
|---|---|---|
| `replace` | XML 문자열 그대로 치환. `expect`(기본 1)와 실제 등장 횟수가 다르면 오류 | `old`·`new` 는 XML 표기. `&` 는 `&amp;`. 같은 글자가 많으면 `<hp:t>…</hp:t>` 까지 넣어 좁히거나 `nth` |
| `replace_text` | 일반 글자 치환(`&`·`<`·`>` 자동 변환) | 한 `<hp:t>` 안에 든 글자만 찾는다. 서식이 바뀌는 지점에서 끊긴 글자는 못 찾음 |
| `fill_empty` | 빈 run 번호에 글자 넣기 | 번호는 **inspect 로 본 원본 기준**. 도구가 큰 번호부터 채워 번호 밀림을 막는다. 빈 run이 아닌 칸(안내 문구가 든 칸)은 `replace` 로 |
| `clone_after` | anchor 문단을 복제해 글자만 바꾼 문단들을 뒤에 붙임 | anchor 는 그 문단의 `<hp:t>` 글자 전체와 정확히 같아야 함 |
| `remove_tables` | marker 가 든 가장 안쪽 표를 통째로 삭제(「삭제 후 제출」 안내 상자) | `forbid` 에 실데이터 문구를 넣어 엉뚱한 표 삭제를 막는다 |
| `use_section` | 이후 op를 다른 섹션에 적용 | |

**`nth` 를 여러 번 쓸 때**: 치환하면 뒤쪽 등장 순서가 바뀐다. 같은 문구의 여러 칸을 채울 때는 **뒤에서부터**(큰 nth부터) 적거나, 매번 `nth: 0` 으로 앞에서부터 하나씩 채운다.

### 3-3. 시험 실행 → 저장

```bash
python scripts/hwpx_fill.py fill "양식.hwpx" "결과.hwpx" spec.json --dry-run
python scripts/hwpx_fill.py fill "양식.hwpx" "결과.hwpx" spec.json
```

- 오류가 하나라도 있으면 저장하지 않고 전부 나열한다(등장 횟수 불일치, 없는 번호, 가드 걸림, XML 문법).
- 저장할 때 `<hp:linesegarray>`(줄 배치 캐시)를 전부 지운다. **글자 길이가 크게 바뀐 채로 캐시가 남아 있으면 한글이 "손상되었거나 변조되었을 가능성" 경고를 띄우며 열기를 거부한다.** 짧은 치환만 하면 증상이 안 나서 원인을 찾기 어렵다. 한글은 열 때 캐시를 다시 계산한다.
- `mimetype` 은 첫 항목에 무압축으로, 나머지 항목은 원래 순서대로 다시 묶는다.
- 결과 파일이 이미 있으면 거부한다. `--overwrite` 를 주면 덮어쓴다(원본 양식은 이것으로도 못 덮어쓴다).

복잡한 작업은 라이브러리로 직접 짜도 된다.

```python
import sys; sys.path.insert(0, r"<이 스킬 폴더>/scripts")
from hwpx_fill import HwpxDoc
doc = HwpxDoc("양식.hwpx")
doc.replace("제00회", "제18회")
doc.fill_empty_many([(1, "회사명"), (2, "대표자")])
print(doc.remove_tables("삭제 후 제출", forbid=["회사명"]))
doc.save("결과.hwpx")   # doc.errors 가 있으면 예외
```

## 4. 본문 윤문 — im-not-ai

양식의 서술형 칸(사업 개요, 추진 경과, 향후 계획, 미달성 사유 등)을 AI가 썼다면 넣기 전에 AI 티를 걷어낸다.

**대상**: 세 문장 이상 이어지는 서술형 문단.
**대상 아님**: 숫자·금액·날짜·이름·연락처·짧은 칸 값, 개조식 한 줄 항목(`◦`·`-`), 법령·양식 고정 문구. 개조식 항목에 돌리면 문장 잇기 규칙(E-1) 때문에 항목이 합쳐질 수 있다.

### 4-1. 설치 확인

이 스킬은 im-not-ai(`epoko77-ai/im-not-ai`, MIT)를 번들하지 않고 플러그인으로 연결한다. 원본 업데이트를 그대로 받기 위해서다.

```bash
claude plugin details humanize-korean@im-not-ai   # 설치 확인
# 없으면:
claude plugin marketplace add epoko77-ai/im-not-ai
claude plugin install humanize-korean@im-not-ai
```

설치돼 있지 않으면 **조용히 건너뛰지 않는다.** 사용자에게 설치 명령을 알리고, 윤문 없이 진행할지 묻는다.

### 4-2. 절차

1. 서술형 칸을 한 파일로 모은다. 칸마다 마크다운 제목을 달아 둔다. 제목이 사라지거나 본문에 흡수되면 humanize-korean 의 구조 검사(`heading_lost`·`heading_absorbed`)가 잡아내므로 칸 경계가 보존된다.
   ```markdown
   ## [F01] 서비스 한줄 소개
   (본문)

   ## [F02] 추진 경과
   (본문)
   ```
2. (선택) 고칠 값어치가 있는지 먼저 본다. `Skill` 도구로 `humanize-korean:humanize-scan` 에 파일 경로를 넘긴다. 대화 안에서 끝나고 파일을 만들지 않는다.
3. 윤문한다. `Skill` 도구로 `humanize-korean:humanize-korean` 를 부르고 인자에 파일 경로와 옵션을 적는다.
   - 장르: 사업계획서·보고서는 `리포트`, 공문·신청서는 `공적`
   - 강도: 정부 양식은 `보수` (격식체 유지, 원문에 없던 표현 금지)
   - 사용자의 `/humanize` 명령은 사용자 전용이다(모델이 부를 수 없음). 모델은 위 스킬 이름으로 부른다.
4. 결과(`_workspace/<run_id>/final.md`)에서 `## [Fxx]` 제목으로 칸을 다시 나눈다. **제목 개수가 원본과 다르면 그 칸은 원문을 쓰고 보고한다.**
5. 숫자·고유명사가 원문과 같은지 대조한다. 다르면 원문을 쓴다.
6. 등급이 C·D로 나오면 사용자에게 알리고 정밀(heavy) 재실행을 권한다. 자동으로 다시 돌리지 않는다.

`_workspace/` 는 **현재 작업 폴더**에 생긴다. 저장소 안에서 돌렸다면 커밋 전에 지우거나 `.gitignore` 를 확인한다.

## 5. 검증 — 완료라고 말하기 전에

1. **XML 문법**: `hwpx_fill.py` 가 저장 전에 검사한다. 통과해도 한글이 열어 준다는 보장은 없다(linesegarray 함정).
2. **한글에서 열기**: `set_automation_mode(true)` → `open_document(결과)`. 경고창 없이 열려야 한다.
3. **구조 비교**: `get_document_structure` 로 쪽수·표 수를 원본 양식과 비교한다. 안내 상자를 지웠다면 그만큼 줄어든 것이 맞다.
4. **채운 값 확인**: `get_text_all` 에서 넣은 값이 보이는지, 양식 자리표시(`00`, `예시)`, `<삭제 후 제출>`, `     원`)가 남았는지 찾는다.
5. **분량**: 쪽수 제한이 있으면 `export_to_pdf` → `measure_layout` 으로 실제 쪽수를 잰다.
6. **파일 확인**: §1-2 PowerShell 명령으로 결과 파일의 존재와 시각을 확인한다.
7. 보고: 채운 칸 수, 비워 둔 칸(자료 없음) 목록, 윤문 결과(등급·변경률), 검증 결과를 알린다. 그리고 "한글에서 직접 열어 확인해 주세요"로 끝낸다.
