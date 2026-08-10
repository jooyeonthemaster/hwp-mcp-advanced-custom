# Advanced HWP MCP Server 설치 가이드 (v2)

Windows + 한글(HWP) 정품 설치 환경에서 동작합니다. COM 자동화 기반이라 **한글 프로그램이 반드시 설치**돼 있어야 합니다.

## 1. 요구사항

- Windows 10/11
- 한글 2010 이상 (한컴오피스)
- Python 3.10 이상

## 2. 패키지 설치

```bash
cd C:\tools\hwp-mcp-advanced-custom
pip install -r requirements.txt
```

`pymupdf`와 `Pillow`는 선택 사항이지만 설치를 권장합니다.
- `pymupdf` 없으면 `measure_layout`(PDF 실측 페이지 측정) 사용 불가
- `Pillow` 없으면 `insert_image`의 원본 비율 자동 계산이 근사값으로 대체됨

## 3. 보안 모듈 등록 (중요)

이 단계를 건너뛰면 **파일에 접근할 때마다 "접근 허용" 대화상자가 떠서 자동화가 멈춥니다.**

저장소에 `security_module/FilePathCheckerModule.dll`이 동봉돼 있고, 서버가 시작할 때 레지스트리
`HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule`을 **자동으로 점검·복구**합니다.
따로 할 일은 없지만, 수동으로 확인하려면:

```bash
reg query "HKCU\Software\HNC\HwpAutomation\Modules"
```

값이 DLL의 절대 경로(`...\security_module\FilePathCheckerModule.dll`)여야 정상입니다.
빈 문자열이나 `FilePathCheckerModule` 같은 이름만 들어 있으면 대화상자가 뜹니다.

## 4. MCP 등록

### Claude Code (CLI / VSCode 확장)

`%USERPROFILE%\.claude.json`의 최상위 `mcpServers`에 추가하면 **모든 프로젝트에서 공용**으로 쓸 수 있습니다.

```json
{
  "mcpServers": {
    "hwp": {
      "type": "stdio",
      "command": "python",
      "args": ["C:/tools/hwp-mcp-advanced-custom/advanced_hwp_server.py"],
      "env": {}
    }
  }
}
```

`command`는 실제 python 실행 파일 경로를 쓰는 편이 안전합니다 (`where python`으로 확인).

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`에 같은 내용을 넣고 앱을 완전히 종료 후 재시작합니다.

## 5. 동작 확인

```bash
python test_v2_live.py
```

실제 한글을 띄워 36개 항목을 검증합니다. 마지막에 `ALL GREEN`이 나오면 정상입니다.
(테스트는 임시 폴더에 사본을 만들어 작업하므로 원본 문서를 건드리지 않습니다.)

MCP 등록 후에는 대화창에서 이렇게 확인할 수 있습니다:

```
새 한글 문서를 만들고 3행 3열 표에 데이터를 채워줘
```

## 6. 문제 해결

| 증상 | 원인·조치 |
|---|---|
| 파일 접근 대화상자가 계속 뜸 | 3번 보안 모듈 등록 확인 |
| `PermissionError` / 문서가 안 열림 | 한글이 파일을 잠근 상태 — `kill_hwp_processes` 도구 실행 |
| 한글이 `¡à ¿¹¼ú`처럼 깨짐 | v2에서 자동 복구됨. 계속되면 구버전 서버를 보고 있는지 경로 확인 |
| MCP 도구가 목록에 안 보임 | 설정 저장 후 Claude를 **완전히 재시작** (서버는 시작 시 1회 로드됨) |
| 표 만든 뒤 내용이 표 안에 들어감 | v2에서 해결됨 (`create_table`이 커서를 표 밖으로 빼냄) |

## 7. 개발 참고

- `check_hwp_api.py` — 설치된 한글 COM 객체의 사용 가능한 메서드·파라미터셋을 덤프하는 진단 스크립트.
  API 이름이 버전마다 달라 확신이 없을 때 이걸로 실물을 확인하세요.
- `test_v2_live.py` — 기능 추가 시 여기에 검증 항목을 함께 추가합니다.
