# PPT 번역 도구

이 프로젝트는 Codex (GPT 6.1 Sol-Medium)을 사용한 바이브코딩으로 개발되었습니다.

PowerPoint 파일의 텍스트를 번역하고, 원본 서식과 수식을 보존해 새 PPTX로 저장하는 Python + PySide6 데스크톱 앱입니다. 번역 후 레이아웃 조정을 선택적으로 실행할 수 있습니다.

## 설치와 실행

Python 3.10 이상이 필요합니다. 레이아웃 조정은 Windows와 Microsoft PowerPoint가 필요합니다.

```powershell
python -m pip install -r requirements-gui.txt
python app.py
```

화면만 확인하려면 데모 모드를 실행합니다. 파일·계정·모델·한도는 모의 데이터이며 실제 AI 호출이나 PPT 저장을 하지 않습니다.

```powershell
python app.py --demo
```

## 사용 순서

1. PPTX 파일을 추가합니다.
2. 연결 관리에서 사용할 AI 도구를 탐색하고 자신의 계정으로 연결합니다.
3. 번역·조정 노드의 도구·모델·추론 수준을 각각 설정합니다. 조정이 필요 없으면 조정 노드의 사용 여부를 해제합니다.
4. PPT 저장 노드에서 폴더와 파일 이름을 지정하고 실행합니다.

기본 결과 경로는 `output/exports/{원본 이름}_{최종 언어}.pptx`입니다. 원본과 기존 결과는 덮어쓰지 않습니다. 중간 결과는 선택한 출력 폴더의 `.pptx-work/<작업 ID>/`에 저장합니다.

## AI 도구

- **Codex App Server**: 실제 모델·지원 추론 수준·계정 한도를 조회하며 텍스트 번역과 이미지 검토를 지원합니다.
- **Antigravity CLI**: 실제 모델 목록·한도를 조회하며 번역과 검증된 모델의 이미지 검토를 지원합니다.
- **공식 Gemini CLI**: 별도 어댑터를 제공합니다. 실제 연결 테스트를 통과한 자동 모드로 번역하며, 이미지 검토는 아직 지원하지 않습니다.

AI 도구는 별도 설치·로그인이 필요하며 각 사용자가 자신의 PC에서 자신의 계정으로 사용합니다. 앱은 설치 경로를 자동 탐색하고 연결 관리에서 수동 경로도 지정할 수 있습니다. 지원하지 않거나 조회하지 못한 모델·한도는 임의로 표시하지 않습니다.

## 문서

| 문서 | 내용 |
| --- | --- |
| [GUI 사용법](docs/USAGE.md) | 연결, 노드 설정, 저장, 중지·재개, 제공 기능 |
| [아키텍처](docs/ARCHITECTURE.md) | GUI·공용 처리·AI 어댑터의 책임과 확장 |
| [개발·검증](docs/DEVELOPMENT.md) | 테스트 실행과 산출물 관리 |
| [Windows EXE 빌드](packaging/windows/README.md) | Python 없는 PC용 실행 파일 빌드·배포 |
| [프로젝트 상태](docs/PROJECT_STATUS.md) | 구현 범위·제약·남은 작업 |
| [AI 개발 지침](AGENTS.md) | 저장소 수정 규칙 |
| [Legacy CLI](legacy/README.md) | 기존 명령행 파이프라인과 근사 복원 도구 |
| [UI 명세](ppt-ui-handoff/UI_SPEC.md) | 기능·상태 전이 기준 |

슬라이드 이미지 속 글자의 OCR, 구형 `.ppt` 입력, 앱 재시작 후 작업 큐 가져오기는 현재 지원하지 않습니다. 최종 번역과 레이아웃은 사용 전에 확인하세요.
