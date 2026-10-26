# 아키텍처

Python + PySide6 네이티브 데스크톱 앱입니다. `ppt-ui-handoff/UI_SPEC.md`는 기능·상태 전이 기준이며 `layout-preview.html`은 배치·색상·상호작용 참고입니다.

## 실행 흐름

`app.py` → `MainWindow` → `FileRunner(QThread)` → `DesktopEngine`

엔진은 PPT 분석 → AI 번역 → 원본 사본에 번역 적용 → 선택적 AI/Python 레이아웃 조정 → 저장 순서로 실행합니다. 번역·조정 노드의 설정은 독립적이고 계정·한도는 같은 도구를 사용하는 노드끼리 공유합니다. 현재 파일의 설정은 시작할 때 복사하며 일반 속성 변경은 다음 파일부터 반영합니다.

번역은 추출한 문서 전체를 우선 한 요청에 묶습니다. 페이지 경계로 나누지 않으며, 서식 표시와 JSON을 포함한 입력이 기본 60,000자 예산을 넘을 때만 항목 단위로 분할합니다. 이는 모델의 컨텍스트 한도를 의미하지 않는 앱의 분량 기준이며, 단일 항목은 잘라서 손상시키지 않습니다. 완료한 묶음의 항목 ID별 번역을 체크포인트에 저장하므로 기존 페이지별 체크포인트도 재사용하고 미완료 항목만 요청합니다. 페이지 진행률은 각 페이지의 모든 항목이 처리됐는지로 계산합니다.

GUI 번역은 서식·줄바꿈 문제를 내용 검증 실패로 취급하지 않습니다. 알려진 의미별 서식 구간은 일부 누락돼도 활용하고, 대응할 수 없는 서식 표시는 제거한 뒤 원문 기본 서식으로 번역을 적용합니다. 문자 위치에 따라 강조 서식을 추정하지 않습니다. 수식 표시의 순서·개수는 지키며, 줄바꿈만 달라졌다면 수식 사이의 번역문을 원래 문단의 수식 위치에 맞춰 배치합니다. 서식·문단 재배치는 `translation_deferred`와 적용 보고서의 `formatting_deferred`에 남기며, 이를 이유로 모든 해당 페이지를 상세 조정으로 보내지 않습니다.

텍스트 적용은 DrawingML의 `a:endParaRPr`가 문단의 마지막 요소가 되도록 삽입 위치를 제한하고 결과 순서를 검사합니다. XML에 텍스트가 존재해도 순서가 잘못되면 PowerPoint에서 표시되지 않을 수 있으므로 단순 텍스트 비교만으로 검증하지 않습니다. `application_version`이 이전인 체크포인트는 저장된 번역을 원본에 재적용하고 이전 검토 이력을 보존합니다.

GUI 조정은 **전체 슬라이드의 단일 PDF 출력 → PDF 일괄 읽기·페이지별 조정 종류 예측 → 해당 페이지 조정** 순서로 시작합니다. PowerPoint에서 현재본 전체를 `document-pdf/slides.pdf`로 내보내고 PDFium으로 모든 PDF 페이지를 읽습니다. 숨긴 슬라이드도 내보내 PDF 페이지 번호와 PPT 슬라이드 번호를 일치시킵니다. 이미지 입력 provider에는 PDF 페이지에서 만든 원본/현재본 JPG 비교 모음을 한 요청에 첨부합니다. PDF 자체의 네이티브 첨부 지원을 가정하지 않습니다. 일부 페이지만 조정할 때도 전체 PDF를 먼저 읽고, 실제 변경은 요청 범위로 제한합니다. PDF·해시·페이지별 예측은 체크포인트에 보존합니다.

각 페이지의 `predicted_adjustments`는 `font/spacing/images/compression` 중 복수 종류를 가질 수 있고, 수정 불필요 페이지는 빈 배열입니다. 해당 예측과 실제 편집 가능 여부로 단계별 요청을 제한합니다. 줄 수 증가·사소한 넘침·측정 누락은 강제 조정 조건이 아닙니다. 자동 후보는 새 넘침 중 3pt와 가용 크기 6%를 모두 초과하는 경우이며 시각 후보는 `overflow/clipping/overlap/unreadable`의 구체적 근거를 요구합니다. 필요한 페이지만 기본 6페이지씩 상세 조정하고, 바뀐 페이지만 재렌더링·재검토합니다. 예측이 실패하면 측정상 명확한 넘침만 폰트·간격으로 처리하고 나머지는 미검토로 기록합니다. 전체 상세 검토로 자동 확대하지 않습니다. 입력 해시·정책·예측 규칙 버전이 같은 결과는 재개 시 재사용합니다. 공용 modifier의 비단계식 모드는 기존 PNG/JPG 선별을 유지합니다.

GUI의 상세 조정은 **폰트 크기 → 공백·간격 → 그림 크기·위치 → 문장 간소화** 순서로 실행합니다. 각 단계의 필드 제한을 모델 지시와 적용 검증 양쪽에서 강제합니다. 기본은 단계당 수정 1회이며, 바뀐 페이지만 렌더링해 읽기 전용 검토로 해결 여부를 확인합니다. 해결된 페이지는 후보에서 제외하고, 해당 방법으로 해결할 수 없거나 검토에 실패한 페이지만 다음 단계로 넘깁니다. 수정 가능한 대상이 없는 단계는 요청을 생략합니다. 문장 간소화 옵션을 끄면 마지막 단계도 생략합니다. 단계별 후보·해결·잔여 페이지와 검토 캐시를 기록합니다. 공용 modifier의 기본 혼합 조정 모드는 기존 CLI 호환을 위해 유지하며 GUI가 `staged=True`를 지정합니다.

공백 조정은 반복 공백·탭·빈 문단과 문단/줄 간격·내부 여백을 줄일 수 있습니다. 단어 사이의 구분 공백·문자·숫자·비어 있지 않은 줄의 순서는 유지합니다. 코드로 식별되는 텍스트의 공백은 보호하고, 수식·명시적 줄바꿈이 있는 프레임의 줄 경계도 유지합니다. 비어 있지 않은 문단을 재사용해 강조·목록 수준을 보존하고 간격·글자 크기·그림 변경에는 누적 제한을 적용합니다.

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
| `pptx_styles.py` | 혼합 서식 구간 표시·검증과 원본 run 서식 대응 |
| `translate_texts.py` | 번역 배치·응답 검증·보호 표시·텍스트 저장 |
| `translation_validation.py` | 번역문을 유지하는 서식 복구·수식 위치 복원·실패 사유 분류 |
| `apply_translation.py` | 원본 ZIP 구조를 보존하며 번역 텍스트 적용 |
| `gemini_cli.py` | 기존 Antigravity CLI 호출·사용량·stream-json 처리 |
| `cli_discovery.py` | 도구 실행 파일 자동 탐색 |
| `pptx_pipeline/ai.py` | AI 응답·provider 계약·레지스트리 |
| `pptx_pipeline/codex_provider.py` | Codex App Server stdio JSON-RPC |
| `pptx_pipeline/stages.py` | 재사용 가능한 분석·번역·적용·조정 단계 |
| `pptx_pipeline/layout.py` | 렌더링 결과 검토·검증된 레이아웃 변경 |
| `pptx_pipeline/layout_adjustments.py` | 조정 단계별 허용 작업·대상 선별·공백/간격 보존 처리 |
| `pptx_pipeline/layout_overview.py` | 전체 JPG 비교 모음·문제 페이지 선별·줄바꿈 변화 감지 |
| `pptx_pipeline/layout_pdf.py` | 전체 PDF 페이지 수 검증·순서 유지·PDFium 렌더링 |
| `pptx_pipeline/layout_text.py` | 선택적 문장 압축의 숫자·단위·수식·서식 보호 검증 |
| `pptx_pipeline/layout_cache.py` | 동일 입력·렌더러의 렌더링 재사용 |
| `pptx_pipeline/rendering.py`, `render_powerpoint.ps1` | PowerPoint COM 렌더링 |

위 모듈은 GUI 의존성이므로 Legacy로 취급하지 않습니다. 특히 `main.py`는 GUI 진입점은 아니지만 분석 엔진으로 사용합니다. `gemini_cli.py`의 기존 이름은 공식 Google Gemini CLI를 의미하지 않으며, 공식 Gemini 어댑터는 `pptx_gui/adapters/gemini.py`입니다.

도형 탐색은 python-pptx가 지원하는 일반·그룹·그림·표/차트 프레임·연결선만 래핑합니다. 지원되지 않는 `p:contentPart`는 원본 트리에 남겨 ZIP 패치 시 보존하며, 호환 분기의 편집 가능한 도형·그림이 있으면 해당 분기를 사용합니다. 그룹 내부에도 같은 규칙을 적용합니다. 추출 데이터와 최종 보고서의 `preserved_unsupported_shapes`에 페이지와 종류를 기록하고 GUI는 경고를 표시하며 `needs_review`로 저장합니다. 삽입 개체 자체의 텍스트/OCR 편집을 지원한다는 의미는 아닙니다.

## Legacy 경계

`legacy/pipeline.py`와 `legacy/copy_pptx.py`는 독립 CLI 도구입니다. Legacy가 활성 공용 코드를 가져오는 것은 허용하지만, GUI·공용 코드는 `legacy`를 가져오지 않습니다. 원본 보존 방식의 번역 적용과 추출 데이터만으로 근사 복원하는 `copy_pptx`는 별개 기능입니다. 테스트는 회귀 검증을 위해 `tests/`에 함께 유지합니다.

## 확장과 안전한 실행

AI 확장은 `AIProvider.generate(prompt, model=None, images=())` 계약과 `AIResponse`를 따릅니다. GUI 연결·인증·모델/한도 조회는 도구별 연결 어댑터가 담당하고, 실제 모델 호출 provider를 엔진에 제공합니다. CLI 레지스트리 등록과 GUI 어댑터 연결은 별도입니다.

이미지 검토 provider는 실제 이미지 첨부 또는 성공한 이미지 열람 근거를 제공해야 합니다. Python이 허용된 수정안을 검증하고 적용합니다. GUI의 문장 압축 옵션을 켜면 필요한 프레임만 의미를 유지하는 짧은 표현을 제안할 수 있습니다. 숫자·수식·수식 줄바꿈·문단 경계·혼합 서식 표시와 보호된 코드·단위·기호·부정 표현을 검사하고 Office Math 구조의 동일성을 확인합니다. 의미 동등성 전체를 기계적으로 보장할 수 없으므로 압축문은 시각·의미 재검토 후에도 `needs_review`와 보고서에 남깁니다. 공용 `LayoutPolicy`의 기본은 압축 불허이며 GUI는 사용자 옵션을 전달합니다.

파일 처리·AI·렌더링은 작업자 스레드에서 실행합니다. UI는 신호로 갱신하며 모델 호출 중지도 소유한 프로세스에만 적용합니다. 원본과 기존 결과는 덮어쓰지 않고, 페이지별 체크포인트를 저장합니다. 인증은 각 공식 도구에 맡기며 토큰·키·인증 URL을 로그나 프로젝트 파일에 저장하지 않습니다.

선택적 조정 실패는 파일 전체 실패로 전파하지 않고 안전한 직전 번역본을 저장하며, 발생 단계·예외 종류와 가린 데이터 오류 설명을 보고합니다. 후보 적용/렌더링 실패는 아직 채택하지 않은 후보를 버려 앞선 검증 결과를 유지합니다. 중지·한도 제어·자원 고갈은 예외이며, 필수 최종 저장의 실제 실패는 계속 전파합니다.

## Windows 배포

[packaging/windows](../packaging/windows/README.md)에서 PyInstaller 빌드를 관리합니다. 전용 `launcher.py`만 frozen 실행의 작업 폴더를 사용자별 쓰기 가능한 위치로 조정하며, UI·엔진 코드를 복제하지 않습니다. 렌더링 PowerShell 스크립트는 패키지의 동일 모듈 경로에 데이터로 포함합니다.
