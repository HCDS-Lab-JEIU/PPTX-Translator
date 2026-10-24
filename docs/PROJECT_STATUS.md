# 프로젝트 상태

## 구현·검증한 범위

모의 화면의 노드 설정 보존·공유 잔량·저장 이름·정지 중 UI 반응과, 페이지 체크포인트 재개 및 영어→한국어 적용을 자동 검증합니다. 현재 계정에서 Codex 모델/한도를 조회하고 임시 1장 PPT의 실제 번역→적용→PowerPoint/Codex 검토→저장 및 원본 보존을 확인했습니다. 로그인·로그아웃은 사용자 계정을 바꾸는 작업이므로 자동 테스트에서 실행하지 않았습니다.

남은 범위: 공식 Gemini CLI 설치 환경에서 실제 인증/모델 실행 검증, Gemini의 직접 로그아웃·정밀 계정 잔량·추론 옵션·이미지 검토 연동, 앱 재시작 후 작업 큐/체크포인트 가져오기, 한도 초기화 시 자동 재개, 별도 PC에서의 배포 검증입니다. 용어·문체 재작성은 현재 조정 엔진이 지원하지 않으므로 UI에 비활성 사유를 표시합니다. 코드 보존은 번역 지시를 사용하며 수식 표시·숫자는 별도로 검증합니다. 조정은 텍스트·수식을 재작성하지 않습니다. 영상·이미지 속 글자 OCR은 하지 않습니다.

공식 참고: [Qt QThread](https://doc.qt.io/qtforpython-6/PySide6/QtCore/QThread.html), [Codex App Server 인증·모델·한도](https://learn.chatgpt.com/docs/app-server), [Gemini 인증](https://geminicli.com/docs/get-started/authentication/), [Gemini headless](https://geminicli.com/docs/cli/headless/), [Gemini 모델 선택](https://geminicli.com/docs/cli/model/).


## 문서·파일 정리

혼합 서식 번역은 원문의 서식 구간을 `[[STYLE_n]]…[[/STYLE_n]]`으로 모델에 전달하고, 번역된 의미 구간에 원본 run의 굵기·색상·기울임 등 서식을 적용합니다. 어순 변경 시 문자 위치로 서식을 배분하지 않습니다. 표시 누락·중복·중첩 및 문단·수식 경계 이동은 저장 전에 거부합니다. 의미 대응의 품질은 번역 모델에 의존하며 실제 문서 검토가 필요합니다. 기존 표시 없는 혼합 서식 번역·체크포인트는 원본을 새 작업으로 추출·번역해야 합니다.

추출 스키마 3의 `texts.txt`는 일반 항목에 기존 JSON 문자열을 유지하고, 혼합 서식 항목에 `{"text": "원문", "styled_text": "서식 표시 포함 원문"}` 객체를 사용합니다. `content.json`의 `text`는 표시 없는 원문이며, 수식 내부 텍스트는 별도 항목으로 유지합니다.

- GUI 진입점은 `app.py`입니다.
- GUI가 사용하지 않는 기존 파이프라인 CLI와 근사 복원 도구를 `legacy/`로 분리했습니다.
- 공용 처리 모듈과 `pptx_pipeline/`는 GUI에서도 사용하므로 활성 코드로 유지합니다.
- README는 사용자용 빠른 시작, 개발 규칙은 `AGENTS.md`, 상세 문서는 `docs/`로 구분합니다.
- 기존 검증 결과는 구현 당시의 기록이며, 각 PC의 설치·계정·모델 지원은 실행 시 다시 조회합니다.

## Windows 패키징

`packaging/windows/`에 PyInstaller spec, 격리 빌드 스크립트와 EXE 전용 진입점을 분리했습니다. Python 런타임을 포함한 폴더·ZIP 배포를 사용합니다. AI CLI와 PowerPoint는 별도 설치 항목입니다. 설치 프로그램·자동 업데이트·코드 서명은 미구현입니다.

Windows x64 빌드와 약 59MB ZIP 생성을 확인했습니다. PATH에 Python이 없는 상태의 EXE 데모 실행·종료와 렌더러 리소스 포함, ZIP CRC·개인 자료 제외를 확인했습니다. 새 빌드 환경에서 회귀 테스트 78개가 통과했습니다. Python이 설치되지 않은 별도 PC의 실제 번역·조정 검증은 남아 있습니다.
