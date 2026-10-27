# Windows 실행 파일 빌드

빌드 설정·전용 진입점·환경·산출물을 이 폴더에서 관리합니다. 앱 구현은 기존 `app.py`와 공용 모듈을 사용합니다.

## 빌드

Windows에 Python 3.10 이상을 설치한 개발 PC에서 프로젝트 루트 기준으로 실행합니다. 일반적인 Windows x64 배포에는 x64 Python을 사용하세요.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows/build.ps1
# 설치한 빌드 환경 재사용
powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows/build.ps1 -SkipInstall
# 진단용 콘솔을 표시하는 빌드
powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows/build.ps1 -SkipInstall -Console
```

`.build-venv/`에 별도 환경을 만들고 의존성을 설치합니다. 기존 개발용 `.venv`나 전역 환경을 수정하지 않습니다. 처음에는 패키지 다운로드가 필요합니다.

`-SkipInstall`을 사용해도 PDF 모듈(`pypdfium2`, `pypdfium2_raw`)과 네이티브 렌더링을 빌드 전에 검사합니다. 누락되면 설치를 생략하지 않고 다시 실행하세요. spec에는 PDF 모듈을 명시적으로 포함합니다. 소스 실행 환경은 별도이므로 `python -m pip install -r requirements-gui.txt`도 해당 Python으로 실행해야 합니다.

## 배포

- `dist/PPTTranslator/PPTTranslator.exe`: GUI 실행 파일
- `dist/PPTTranslator/_internal/`: Python 런타임·라이브러리·리소스
- `dist/PPTTranslator-windows.zip`: 배포용 전체 폴더 ZIP
- `build/`: PyInstaller 분석·캐시

**ZIP 압축을 풀어 폴더 전체를 전달합니다. EXE만 복사하면 실행되지 않습니다.** 설치 프로그램·자동 업데이트·코드 서명은 아직 제공하지 않습니다. 빌드 환경과 산출물은 Git에서 제외합니다.

배포 PC에는 Python 설치가 필요 없습니다. AI CLI는 사용자가 별도로 설치·로그인해야 하며, 조정을 사용할 때만 Microsoft PowerPoint가 필요합니다. Node.js 기반 CLI에는 해당 CLI가 요구하는 Node.js도 필요합니다.

## 포함 범위와 저장 위치

앱 import를 기준으로 Python·Qt 의존성을 묶고 `render_powerpoint.ps1`을 `pptx_pipeline/` 아래에 포함합니다. 원본 PPTX, `output/`, 로컬 설정, 인증 정보, 외부 AI CLI와 Legacy 도구는 수집하지 않습니다. 저장소 전체를 데이터로 복사하지 않습니다.

패키지 리소스는 모듈의 `__file__` 기준으로 찾습니다. EXE 전용 `launcher.py`는 작업 폴더를 `%LOCALAPPDATA%/PPTTranslator`로 설정하므로 기본 결과는 그 아래 `output/exports/`에 저장합니다. PPT 저장 노드에서 다른 폴더를 선택할 수 있습니다. Python으로 `app.py`를 실행하는 동작은 바뀌지 않습니다.

## 확인

```powershell
& .\packaging\windows\dist\PPTTranslator\PPTTranslator.exe --demo
# PATH에서 Python을 제거한 상태로 자동 데모 실행·스크린샷·렌더러 포함 확인
powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows/verify.ps1
```

Python이 없는 별도 Windows PC에서도 압축 해제·GUI 실행·파일 선택을 확인하세요. 실제 CLI 자동 탐색·로그인과 작은 PPT의 번역·조정 생략·선택적 PowerPoint 조정·저장·원본 보존도 확인해야 합니다. 개발 PC의 데모 성공만으로 별도 PC 검증을 대체하지 않습니다.

확인한 빌드: `v0.3.0-alpha`, Windows x64, Python 3.13, PyInstaller 6.22.3. 격리 환경에서 회귀 테스트 150개가 통과했으며, PATH·PYTHONHOME·PYTHONPATH를 비운 상태에서 EXE 데모의 화면 생성과 정상 종료를 확인했습니다. 약 63MiB ZIP의 CRC·배포 폴더와의 일치, PDF 전체 예측·단계별 조정·영문 병기 재사용 모듈 포함과 번들 PDFium DLL의 로드·문서 생성·종료를 확인했습니다. 실제 사용자 원본·인증 파일·외부 CLI는 포함하지 않으며, 라이브러리 리소스인 `pptx/templates/default.pptx`는 포함합니다. 배포 ZIP 옆의 `SHA256SUMS.txt`로 무결성을 확인할 수 있습니다. 별도 PC의 실제 AI·PowerPoint 전체 흐름은 아직 검증하지 않았습니다.

Windows에서 빌드하며 [폴더 배포](https://pyinstaller.org/en/stable/operating-mode.html), [spec 데이터 포함](https://pyinstaller.org/en/stable/spec-files.html), [리소스 경로](https://pyinstaller.org/en/stable/runtime-information.html)는 PyInstaller 공식 문서를 따릅니다. 공개 배포 전 포함 라이브러리의 라이선스와 배포 조건도 확인해야 합니다.
