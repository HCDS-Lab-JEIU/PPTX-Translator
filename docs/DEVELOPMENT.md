# 개발과 검증

프로젝트 루트에서 실행합니다. Python 3.10 이상을 사용하고 GUI 의존성을 설치합니다.

```powershell
python -m pip install -r requirements-gui.txt
python -B -m unittest discover -v
python app.py --demo
python -m legacy.pipeline --help
python -m legacy.copy_pptx --help
```

GUI 자동 화면 확인은 `python app.py --demo --screenshot output/gui-demo.png --quit-after 1800`으로 실행합니다. 데모 모드는 실제 AI 호출·인증 변경·PPT 저장을 하지 않습니다.

`tests/`는 추출·수식·번역 적용·AI 프로토콜·GUI·자동 탐색·Legacy CLI의 회귀 테스트를 포함합니다. 실제 계정 로그인/로그아웃은 자동 테스트하지 않습니다. 이미지 검토를 실제 검증하려면 Windows PowerPoint와 연결된 이미지 지원 AI가 필요합니다.

`tests/test_recovery.py`는 잘못된 응답의 항목별 재시도·원문 유지, 수식 줄바꿈 보호, 적용 도중 부분 변경의 복원, 조정 예외·임시 파일 접근 실패 후 저장과 단계/종류 로그, 사용자 중지·자원 고갈 전파를 모의 응답으로 검증합니다. 필수 저장 오류는 계속 전파하며 복구 결과는 `needs_review`로 보고합니다.

47장 문서가 한 번의 번역 요청으로 처리되는지, 분할 예산이 서식 포함 JSON 크기를 계산하는지 검증합니다. 체크포인트 재개 테스트는 작은 예산을 주입해 묶음 사이의 중단과 완료 항목 재사용을 확인합니다.

`tests/test_layout_overview.py`는 전체 문서 JPG 선별, 후보 페이지만 수정·재렌더링, 줄 수 증가·미리보기 불확실성에 의한 과잉 선별 억제, 큰 넘침만 강제 선별, 전체 검토 실패 시 자동 수정 확대 방지, 중지 후 캐시 재사용과 선택적 압축의 보호 검증을 포함합니다. 실제 PowerPoint 렌더링과 모델의 시각·의미 판단 품질은 별도로 확인합니다.

`tests/test_translation_tolerance.py`는 서식 표시 누락·비어 있는 조사 구간·문단 수 변화에서도 번역과 가능한 의미별 서식을 유지하는지 확인합니다. 숫자·수식 보호, native 수식 위치·XML 유지, 적용 보고서 및 이전 체크포인트의 실패 항목만 재시도하는 동작도 검증합니다.

문단 재구성 회귀는 실제 원본처럼 `a:endParaRPr`가 있는 문단으로 검사합니다. 일반 텍스트·서식 표시 경로에서 문단 끝 서식 뒤에 텍스트가 생기지 않아야 하며, 이전 적용 결과는 모델 호출 없이 저장된 번역으로 다시 생성해야 합니다. 렌더링 후보 실패 시 직전 검증본을 채택하는 회귀도 포함합니다. 텍스트가 XML에 있다는 것과 PowerPoint에 표시된다는 것을 구분해 실제 렌더링을 확인합니다.

`tests/test_unsupported_shapes.py`는 미지원 `p:contentPart`로 발생하는 `has_ph_elm`/`rot` 오류의 회귀를 검사합니다. 직접 삽입·그룹·AlternateContent 호환 그림 선택과 번역·레이아웃 적용 후 개체 XML 구조 및 관련 ZIP 부품 보존을 확인합니다.

`tests/test_layout_stages.py`는 폰트→공백·간격→그림→문장 간소화 순서, 렌더링 후 검토된 해결 페이지의 후보 제거, 단계 경계에서 금지된 작업 차단, 검토 실패 페이지 유지, 축약 비활성화, 단계별 캐시 재개를 검증합니다. 공백 수정 시 숫자·단어 구분·코드 보호와 빈 문단 제거 후 강조·목록 속성 보존도 검사합니다. 모델 응답은 모의 데이터이며 유료 호출은 하지 않습니다.

`tests/test_pdf_planning.py`는 전체 PDF가 상세 조정보다 먼저 존재하는지, 모든 PDF 페이지의 예측·복수 조정 종류·불필요 단계 생략·일부 페이지만 변경하는 실행에서도 전체 PDF 검토·예측 누락 시 무분별한 조정 방지·페이지 수 불일치·PDF 삭제 후 캐시 복구를 확인합니다. PDFium은 동시 호출을 잠금으로 제한하며 외부 Poppler 설치 없이 앱에서 PDF 페이지를 읽습니다. 실제 PowerPoint의 PDF 출력은 별도 통합 검증이 필요합니다.

## 산출물 관리

원본 자료와 사용자 환경 `.venv`, `.vscode`는 임의로 정리하지 않습니다. 실행 결과는 `output/` 또는 사용자가 지정한 폴더에 저장하며 버전 관리에서 제외합니다.

GUI는 출력 폴더의 `.pptx-work/<작업 ID>/`에 중간 데이터를 저장합니다. 실행 중인 체크포인트·번역·렌더링 파일을 삭제하지 마세요. 최종 PPT, 결과 보고서와 보고서가 참조하는 자료를 함께 유지합니다.

Legacy CLI는 `pipeline.json`, `modified.layout.json` 등에 파일 경로를 기록합니다. 완료 후 사용하지 않는 중간 PPT·이미지·원본 이벤트를 정리할 때 보고서의 참조도 확인해야 합니다. `needs_review` 결과의 최종 이미지와 검토 자료는 남깁니다. 재개에 필요한 추출·번역 파일을 삭제하면 해당 단계부터 다시 처리해야 합니다.

## 문서 관리

- 사용자 설치·빠른 시작: [README](../README.md)
- GUI 상세 사용법과 제공 기능: [USAGE](USAGE.md)
- 모듈 책임과 확장: [ARCHITECTURE](ARCHITECTURE.md)
- 구현 상태·제약·다음 작업: [PROJECT_STATUS](PROJECT_STATUS.md)
- 기존 CLI 사용법: [legacy/README](../legacy/README.md)

이전 검증 당시의 계정 모델 수·한도 수치를 제품의 고정 사양으로 문서화하지 않습니다. 지원 상태가 바뀌면 사용자 안내와 상태 문서를 함께 갱신합니다.

## Git 공개 전 확인

`.gitignore`는 원본 자료 `ppt resource/`, 산출물 `output/`, `.venv/`, 로컬 편집기·도구 설정, 환경 파일, 일반적인 인증 파일·개인 키·로그를 제외합니다. 원본 PPTX에는 작성자·최종 수정자 메타데이터가 있을 수 있으므로 공개 예제는 별도로 검토한 자료를 사용합니다.

CLI 경로는 PATH·환경 변수·사용자 홈을 기준으로 탐색합니다. 개발자 PC의 사용자 이름·설치 버전·절대경로를 코드나 문서에 추가하지 않습니다. `Path.home()` 기반의 VS Code 확장 탐색은 각 사용자 PC의 설치를 찾는 동작입니다.

Git 저장소 생성 후 커밋 전 `git status --short`와 `git diff --cached`로 실제 스테이징 내용을 확인합니다. `.gitignore`는 이미 추적 중인 파일이나 과거 커밋을 제거하지 않습니다. 인증 토큰이 포함된 파일을 과거에 커밋했다면 제외 규칙 추가만으로 해결되지 않습니다.
