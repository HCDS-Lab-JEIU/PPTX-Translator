# 프로젝트 개발 지침

## 읽을 문서

- 사용자 소개와 실행: [README.md](README.md)
- 구조·모듈 책임: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- 테스트·산출물 관리: [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)
- 현재 제약·남은 작업: [docs/PROJECT_STATUS.md](docs/PROJECT_STATUS.md)
- Windows 배포 변경 시: [packaging/windows/README.md](packaging/windows/README.md)
- GUI 변경 시: [ppt-ui-handoff/UI_SPEC.md](ppt-ui-handoff/UI_SPEC.md), [시각 참고 HTML](ppt-ui-handoff/layout-preview.html)

## 구현 규칙

- Python + PySide6 네이티브 데스크톱 앱을 유지한다. GUI 진입점은 `app.py`다.
- 위젯은 배치와 책임별로 분리하고 메인 창은 조합·상태 연결에 집중한다.
- 파일 처리·모델 호출·PowerPoint 렌더링으로 GUI 스레드를 블로킹하지 않는다.
- 번역과 조정의 도구·모델·추론 설정은 독립적으로 유지한다. 계정 연결·한도는 도구별로 공유한다.
- Codex App Server, 공식 Gemini CLI, Antigravity CLI를 구분한다. 실제 조회·검증한 모델과 기능만 노출한다.
- 원본 PPTX와 기존 결과를 덮어쓰지 않는다. Office Math 구조와 보호 표시를 유지한다.
- 자격증명 파일을 직접 읽거나 복사하지 않는다. 키·토큰·인증 URL·계정 인증 응답을 로그나 프로젝트 파일에 남기지 않는다.
- 데모 데이터는 명확히 표시하고 실제 호출·저장과 섞지 않는다.

## 활성 코드와 Legacy

- `pptx_gui/`, `pptx_pipeline/` 및 루트의 공용 처리 모듈은 GUI가 사용하는 활성 코드다.
- `main.py`, `translate_texts.py`, `apply_translation.py`, `gemini_cli.py`는 CLI 기능도 있지만 GUI 의존성이므로 Legacy로 옮기지 않는다.
- `legacy/`에는 GUI가 사용하지 않는 독립 CLI 도구를 둔다. 활성 코드에서 `legacy`를 import하지 않는다.
- Legacy 명령은 프로젝트 루트에서 `python -m legacy.pipeline`, `python -m legacy.copy_pptx`로 실행한다.
- 경로를 변경하면 import, mock patch 경로, 문서 명령도 함께 갱신한다.

## 검증과 문서

```powershell
python -B -m unittest discover -v
python app.py --demo
python -m legacy.pipeline --help
python -m legacy.copy_pptx --help
```

- 동작·경로 변경 후 관련 회귀 테스트를 실행한다. 인증 변경이나 유료 모델 호출은 일반 단위 테스트에 포함하지 않는다.
- README는 GitHub 사용자용으로 짧게 유지한다. 상세 설계·개발 기록·미구현 사항은 `docs/`의 해당 문서에 기록한다.
- `output/`, `.venv/`, 캐시는 버전 관리 대상이 아니다. 사용자 원본·환경과 현재 실행의 체크포인트를 임의로 삭제하지 않는다.
