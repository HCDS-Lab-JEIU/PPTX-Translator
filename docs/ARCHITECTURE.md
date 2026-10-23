# 아키텍처

Python + PySide6 네이티브 데스크톱 앱입니다. `ppt-ui-handoff/UI_SPEC.md`는 기능·상태 전이 기준이며 `layout-preview.html`은 배치·색상·상호작용 참고입니다.

## 실행 흐름

`app.py` → `MainWindow` → `FileRunner(QThread)` → `DesktopEngine`

엔진은 PPT 분석 → AI 번역 → 원본 사본에 번역 적용 → 선택적 AI/Python 레이아웃 조정 → 저장 순서로 실행합니다. 번역·조정 노드의 설정은 독립적이고 계정·한도는 같은 도구를 사용하는 노드끼리 공유합니다. 현재 파일의 설정은 시작할 때 복사하며 일반 속성 변경은 다음 파일부터 반영합니다.

## GUI 모듈


- `app.py`: QApplication 생성·테마·메인 창을 조합하는 진입점.
- `pptx_gui/main_window.py`: 화면 조합, 파일 큐, 노드 선택, 시작 전 검증.
- `pptx_gui/widgets/files_panel.py`, `pipeline_panel.py`, `inspector.py`, `quota_panel.py`, `connections_dialog.py`, `progress_panel.py`: 배치별 독립 위젯.
- `pptx_gui/models.py`: 파일·노드·저장 이름·모델·한도 상태 데이터.
- `pptx_gui/connections.py`: 도구별 공유 계정/한도 캐시와 비동기 조회.
- `pptx_gui/adapters/codex.py`, `gemini.py`, `antigravity.py`: 서로 다른 공식 도구의 인증·지원 기능·모델/한도 조회와 작업 어댑터.
- `pptx_gui/runner.py`, `engine.py`: 백그라운드 파일 큐, 페이지 체크포인트, 중지·재개·저장.


## 공용 처리 모듈

| 경로 | 책임 |
| --- | --- |
| `main.py` | PPTX 분석·추출·좌표·해시 및 독립 추출 CLI |
| `pptx_text.py` | 텍스트·Office Math XML 처리 |
| `translate_texts.py` | 번역 배치·응답 검증·보호 표시·텍스트 저장 |
| `apply_translation.py` | 원본 ZIP 구조를 보존하며 번역 텍스트 적용 |
| `gemini_cli.py` | 기존 Antigravity CLI 호출·사용량·stream-json 처리 |
| `cli_discovery.py` | 도구 실행 파일 자동 탐색 |
| `pptx_pipeline/ai.py` | AI 응답·provider 계약·레지스트리 |
| `pptx_pipeline/codex_provider.py` | Codex App Server stdio JSON-RPC |
| `pptx_pipeline/stages.py` | 재사용 가능한 분석·번역·적용·조정 단계 |
| `pptx_pipeline/layout.py` | 렌더링 결과 검토·검증된 레이아웃 변경 |
| `pptx_pipeline/rendering.py`, `render_powerpoint.ps1` | PowerPoint COM 렌더링 |

위 모듈은 GUI 의존성이므로 Legacy로 취급하지 않습니다. 특히 `main.py`는 GUI 진입점은 아니지만 분석 엔진으로 사용합니다. `gemini_cli.py`의 기존 이름은 공식 Google Gemini CLI를 의미하지 않으며, 공식 Gemini 어댑터는 `pptx_gui/adapters/gemini.py`입니다.

## Legacy 경계

`legacy/pipeline.py`와 `legacy/copy_pptx.py`는 독립 CLI 도구입니다. Legacy가 활성 공용 코드를 가져오는 것은 허용하지만, GUI·공용 코드는 `legacy`를 가져오지 않습니다. 원본 보존 방식의 번역 적용과 추출 데이터만으로 근사 복원하는 `copy_pptx`는 별개 기능입니다. 테스트는 회귀 검증을 위해 `tests/`에 함께 유지합니다.

## 확장과 안전한 실행

AI 확장은 `AIProvider.generate(prompt, model=None, images=())` 계약과 `AIResponse`를 따릅니다. GUI 연결·인증·모델/한도 조회는 도구별 연결 어댑터가 담당하고, 실제 모델 호출 provider를 엔진에 제공합니다. CLI 레지스트리 등록과 GUI 어댑터 연결은 별도입니다.

이미지 검토 provider는 실제 이미지 첨부 또는 성공한 이미지 열람 근거를 제공해야 합니다. Python이 허용된 수정안을 검증하고 적용하며, 조정 단계에서 본문·수식을 다시 작성하지 않습니다.

파일 처리·AI·렌더링은 작업자 스레드에서 실행합니다. UI는 신호로 갱신하며 모델 호출 중지도 소유한 프로세스에만 적용합니다. 원본과 기존 결과는 덮어쓰지 않고, 페이지별 체크포인트를 저장합니다. 인증은 각 공식 도구에 맡기며 토큰·키·인증 URL을 로그나 프로젝트 파일에 저장하지 않습니다.

## Windows 배포

[packaging/windows](../packaging/windows/README.md)에서 PyInstaller 빌드를 관리합니다. 전용 `launcher.py`만 frozen 실행의 작업 폴더를 사용자별 쓰기 가능한 위치로 조정하며, UI·엔진 코드를 복제하지 않습니다. 렌더링 PowerShell 스크립트는 패키지의 동일 모듈 경로에 데이터로 포함합니다.
