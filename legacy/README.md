# Legacy CLI 사용법

> GUI에서 사용하지 않는 독립 CLI 도구를 보관하는 폴더입니다. 모든 명령은 프로젝트 루트에서 실행하세요.
> `python legacy/pipeline.py` 대신 `python -m legacy.pipeline`을 사용해야 공용 모듈을 정상적으로 찾습니다.

| 파일 | 기능 |
| --- | --- |
| `pipeline.py` | 기존 명령행 파이프라인 |
| `copy_pptx.py` | 추출 JSON·이미지로 근사 PPTX 재구성 |

아래 `main.py`, `translate_texts.py`, `apply_translation.py`, `gemini_cli.py`는 GUI에서도 사용하는 공용 모듈이므로 루트에 유지합니다. 독립 실행 방법도 함께 기록합니다. Legacy CLI의 기본 도구·출력 경로·재개 동작은 GUI와 다릅니다.


Python 3.10 이상. GUI와 PowerPoint 설치 없이 `python-pptx`로 읽습니다.

## 실행

```powershell
python -m pip install -r requirements.txt
python main.py "ppt resource\[SemiconductorCircuits] 02_전기회로소자.pptx"
python main.py "자료.pptx" --output "output/자료"
```

기본 저장 위치는 실행 디렉터리의 `output`입니다.

- `output/texts.txt`: 필터링된 텍스트. 터미널 표준 출력에도 같은 내용을 출력합니다.
- `output/content.json`: 페이지별 텍스트·이미지 파일 경로·위치·제외 건수·경고 목록.
- `output/resources/image_<sha256>.<확장자>`: PPTX에 포함된 원본 이미지. 동일 이미지는 한 파일로 저장하고 각 배치의 페이지·위치는 JSON에 남깁니다.

진행 결과와 경고는 표준 오류(stderr)에 출력합니다. 다른 PPTX를 같은 출력 폴더로 실행하면 `texts.txt`와 `content.json`을 갱신합니다. 기존 이미지 파일은 삭제하지 않으므로 현재 문서의 이미지 목록은 `content.json`을 기준으로 읽으세요. 여러 문서는 `--output`으로 각각 폴더를 지정할 수 있습니다.

## 텍스트 형식

```text
[page=2][x=1.2000,y=2.3000,w=8.0000,h=1.5000 cm; shape=5]["전기 회로\n저항과 전류"]
[page=2][x=1.2000,y=4.0000,w=3.0000,h=0.8000 cm; shape=8; table=8,row=1,col=2]["전류(A)"]
```

페이지는 1부터 시작합니다. 위치는 슬라이드 왼쪽 위가 원점인 cm 단위입니다.
`x,y`는 도형의 왼쪽 위, `w,h`는 크기이며 회전한 도형은 회전 후 외접 사각형입니다.
텍스트는 도형 단위, 표는 셀 단위로 기록합니다. 문자열은 JSON 방식으로 이스케이프하므로 줄바꿈·따옴표를 안전하게 구분할 수 있습니다. 그룹 도형은 확대·축소·회전을 반영하여 슬라이드 좌표로 변환합니다. 표의 행·열도 1부터 시작하며 병합 셀의 범위는 JSON의 `row_span`, `column_span`에 기록합니다.

위에서 아래, 같은 높이에서는 왼쪽에서 오른쪽 순으로 정렬합니다. 복잡한 다단 편집에서 의미상의 읽기 순서까지 추론하지는 않습니다. JSON의 `slides[].texts[]`는 `page`, `shape_id`, `kind`, `is_title`, `position_cm`, `text`를 포함합니다. 빈 페이지도 JSON에는 남습니다.

## 불필요한 텍스트 제외

- 머리말·꼬리말·슬라이드 번호·날짜 placeholder는 제외합니다.
- 슬라이드 상단/하단 10%에 완전히 들어가는 일반 텍스트 상자 중 `3`, `- 3 -`, `3 / 35`, `Page 3`, `3쪽` 같은 페이지 번호를 제외합니다.
- 같은 여백 위치에서 동일한 짧은 문구가 전체 페이지의 60% 이상이면서 최소 3개 페이지에서 반복되면 머리말·꼬리말로 보고 제외합니다.
- 제목 placeholder, 표 셀, 본문 영역의 숫자는 위의 휴리스틱 필터에서 유지합니다.

일반 텍스트로 작성한 머리말·꼬리말은 위치와 반복으로 판별하므로 모든 PPTX에서 완벽하게 구분할 수는 없습니다. 삭제한 텍스트 자체는 결과에 저장하지 않습니다.

## 지원 범위

슬라이드와 그룹 내부의 텍스트, 표, 그림 도형을 추출합니다. 이미지는 화면 렌더링이 아닌 원본 파일이며 자르기 비율과 회전은 JSON에 별도로 기록합니다. 슬라이드 배경·마스터·레이아웃의 그림, 도형 채우기에 사용된 그림, 차트·SmartArt 렌더링, 이미지 OCR은 지원하지 않습니다. 발표자 노트와 자동 생성 description은 이번 CLI 출력에서 제외합니다. 구형 `.ppt`는 `.pptx`로 변환해야 합니다.

API 문서: [python-pptx](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)

검증 테스트: `python -m unittest -v`

## 추출 결과로 PPTX 재구성

```powershell
python -m legacy.copy_pptx
python -m legacy.copy_pptx "output/content.json" -o "output/copy.pptx"
```

기본 입력은 `output/content.json`, 출력은 같은 폴더의 `copy.pptx`입니다.
JSON과 상대 경로의 `resources` 이미지만 읽으며 원본 PPTX는 사용하지 않습니다.
`copy.report.json`에는 복원한 페이지·텍스트·이미지 수와 복원 한계를 기록합니다.

- 슬라이드 크기·페이지 순서·빈 페이지를 유지합니다.
- 텍스트와 표 셀은 저장된 위치·크기의 편집 가능한 텍스트 상자로 생성합니다.
- 이미지는 원본 바이트, 위치·크기, 자르기 비율과 개별 회전을 반영합니다.
- 새로 추출한 JSON은 `z_order`를 사용해 텍스트·이미지 겹침 순서를 유지합니다. 기존 JSON은 이미지 위에 텍스트를 배치합니다.
- 글꼴 정보가 없는 기존 추출 결과에는 맑은 고딕, 본문 18pt·제목 28pt를 적용합니다. `--font`, `--body-size`, `--title-size`로 변경할 수 있습니다. 글자 크기 자동 맞춤은 PowerPoint에서 처리합니다.

**원본과 완전히 동일한 PPTX를 복원하는 기능은 아닙니다.** 추출 시 보존하지 않은 글꼴·색상·문단 서식·배경·일반 도형·차트·SmartArt·애니메이션, 삭제한 머리말·꼬리말·페이지 번호는 복원할 수 없습니다. 표는 셀 텍스트만 남아 있어 테두리와 빈 셀은 복원하지 않습니다. 텍스트 회전, 그룹 이미지의 회전·뒤집기·기울임도 완전 복원 대상에서 제외됩니다. 이미지의 회전 전 크기를 외접 사각형에서 결정할 수 없는 45도 등의 경우는 회전 없이 배치하고 보고서에 기록합니다.

완전한 시각적 복원을 원한다면 추출 단계에서 서식·테마·도형·전체 변환 정보까지 저장하도록 형식을 확장해야 합니다.

## Gemini 호출 클래스 (현재 Antigravity CLI)

`gemini_cli.py`의 `GeminiCLI` 클래스는 현재 사용 중인 Antigravity CLI(`agy`)를 호출합니다. 추가 Python 패키지는 필요하지 않습니다. PATH 및 Windows의 `%LOCALAPPDATA%/agy/bin/agy.exe` 등을 탐색하며 기존 CLI 로그인 상태를 사용합니다.

테스트용 기본 실행은 `agy --print /usage --output-format text`를 호출합니다.

```powershell
python gemini_cli.py
python gemini_cli.py --json
python gemini_cli.py --executable "$env:LOCALAPPDATA/agy/bin/agy.exe" --timeout 60
python -m unittest tests.test_gemini_cli -v
```

```python
from gemini_cli import GeminiCLI

cli = GeminiCLI()
usage = cli.get_usage()
for quota in usage.quotas:
    print(quota.model_group, quota.remaining_percent, quota.reset_at)

# 일반 프롬프트 호출이 필요한 경우 (모델 호출 발생):
# response = cli.run("질문 내용")
# print(response.stdout)
```

`get_usage()`는 모델 그룹별 남은 사용량 비율(`remaining_percent`), 제한 이름(`limit_name`), 리셋 시각(`reset_at`, 시간대 포함 ISO 8601), 원본 보고서(`raw_output`)를 반환합니다. `--json`은 이를 JSON으로 출력합니다. 실행 파일은 `ANTIGRAVITY_CLI_EXECUTABLE` 환경변수로도 지정할 수 있습니다.

Antigravity CLI 1.4.6의 4열 TSV 응답을 기준으로 파싱합니다. CLI 오류·시간 초과·알 수 없는 보고서 형식은 예외로 반환하고, 실패 응답은 `GeminiCLIError.result`에서 확인할 수 있습니다. 새로운 사용량을 임의로 추정하거나 일반 생성 응답의 토큰 수를 계정 잔여량으로 사용하지 않습니다. 로그인되지 않은 경우 먼저 터미널에서 `agy`로 로그인하세요. 구형 Gemini CLI의 `/stats`와는 별도 기능이며 자동으로 대체하지 않습니다.

공식 문서: [Antigravity 사용량](https://www.antigravity.google/docs/cli/commands/usage), [headless 호출](https://www.antigravity.google/docs/cli/headless).

## 추출 텍스트를 영어로 번역

```powershell
python translate_texts.py
python translate_texts.py "output/texts.txt" -o "output/translated.txt"
```

기본 입력은 `output/texts.txt`, 출력은 입력 파일과 같은 폴더의 `translated.txt`입니다.
`PPTXTextTranslator`가 추출 텍스트를 읽어 `GeminiCLI.generate()`로 Gemini에 전달하고, 응답을 검증한 뒤 UTF-8로 저장합니다. 기존 CLI 로그인과 모델 사용량을 사용합니다.

- `[page=...][위치·크기]` 부분은 Python이 그대로 유지하고 마지막 텍스트 문자열만 영어 번역으로 바꿉니다. 표의 행·열 정보도 유지합니다.
- 모델에는 ID·페이지·본문을 전달하고, 수식·숫자·기호·단위·이름·기존 영어·문단을 유지하도록 요청합니다.
- ID별 누락·중복·빈 번역을 검증하고 입력 순서로 재조립합니다. 긴 입력은 배치로 나누고 stdin으로 전달하여 Windows 명령줄 길이 제한을 피합니다. 개별 텍스트 상자는 분할하지 않으므로 큰 상자 하나는 배치 목표 크기를 넘을 수 있습니다.
- 원문의 숫자 토큰과 개수도 비교합니다. 숫자를 생략하거나 단어로 풀어쓴 항목은 해당 항목만 한 번 재번역하며, 그래도 일치하지 않으면 저장하지 않습니다.
- 배치마다 `translated.partial.txt`와 `translated.progress.json`을 저장합니다. 모든 배치를 처리한 뒤 최종 파일을 교체합니다. 빈 응답이 지속된 항목은 원문으로 유지하고 `needs_review`로 기록합니다. 인증 오류·검증 실패 등으로 중단되면 기존 최종 파일은 유지하며, 재실행하면 처음부터 번역합니다.
- 기본 모델은 설치 환경에서 확인한 `gemini-3.8-flash-medium`입니다. `--model`로 변경할 수 있으며, 모델 목록은 `agy models`에서 확인합니다. `--timeout`은 배치별 제한 시간(기본 300초), `--batch-chars`는 입력 배치 목표 문자 수(기본 8000)입니다.

```python
from gemini_cli import GeminiCLI
from translate_texts import PPTXTextTranslator

translator = PPTXTextTranslator(GeminiCLI(timeout=300))
report = translator.translate_file("output/texts.txt")
print(report["output"])
```

PPTX에서 시작하는 전체 실행 순서:

```powershell
python main.py "자료.pptx"
python translate_texts.py
```

이번 단계는 번역 텍스트 파일을 생성하며 `content.json`과 PPTX는 변경하지 않습니다.
번역 품질과 수식 보존은 모델 출력에 의존하므로 최종 사용 전에 내용을 확인하세요.

## 원본 PPTX에 번역 적용

```powershell
python apply_translation.py "ppt resource/[SemiconductorCircuits] 02_전기회로소자.pptx"
python apply_translation.py "자료.pptx" --manifest output/content.json --translations output/translated.txt -o output/translated.pptx
```

기본 출력은 번역 텍스트와 같은 폴더의 `translated.pptx`이며 처리 건수는 `translated.report.json`에 기록합니다.

원본을 임시 사본으로 복사한 뒤, 한글이 포함된 추출 텍스트를 영어로 교체합니다. 영어·숫자·기호만 있는 텍스트는 원본 그대로 유지합니다. 이미 제외되어 번역하지 않은 페이지 번호·머리말·꼬리말도 원본 사본에 남습니다.

- `content.json`의 페이지·도형 ID·표 행/열로 원본을 찾고, `translated.txt`의 페이지·위치 및 도형 ID를 대조합니다. 순서가 바뀐 번역도 위치와 ID로 연결합니다.
- 새 추출 파일에는 `shape=도형ID`와 원본 파일 SHA-256을 기록합니다. 기존 번역 파일에 도형 ID가 없으면 위치를 통해 `content.json`의 ID를 찾습니다. 같은 위치가 중복되어 대상을 확정할 수 없으면 중단합니다.
- 원문 텍스트·도형 위치·슬라이드 수·크기를 검증합니다. SHA-256이 있는 새 데이터는 원본 파일 일치 여부도 검사합니다. 누락·중복·원본 불일치 시 기존 출력 파일을 덮어쓰지 않습니다.
- 슬라이드의 텍스트 XML만 수정하고, 이미지·테마·마스터·배경·노트·관계 파일 등 나머지 ZIP 항목은 원본 바이트 그대로 유지합니다. 도형 위치·크기·회전과 텍스트 프레임 설정도 유지합니다.
- 문단과 run의 글꼴·크기·색상 설정을 보존하며, 번역 전후 공통 부분은 원래 run에 남깁니다. 단어 순서나 길이가 바뀌므로 한 문장 안의 혼합 서식이 번역 단어에 대응하는 위치까지 완벽히 같지는 않습니다. 줄 수가 달라지면 기존 문단 서식을 복제하여 문단을 재구성하고 보고서에 기록합니다.

글자 크기나 상자 크기는 자동으로 바꾸지 않습니다. 영어가 길어질 수 있으므로 생성한 PPTX의 줄바꿈·넘침은 PowerPoint에서 확인하세요.

전체 흐름은 `main.py 원본.pptx` → `translate_texts.py` → `apply_translation.py 원본.pptx`입니다. 기존 `legacy/copy_pptx.py`는 추출 데이터만으로 새 PPTX를 만드는 별도의 기능입니다.

## 수식 포함 텍스트

`AlternateContent/Choice` 내부의 텍스트 상자·표도 추출합니다. 중복 추출을 피하기 위해 현대 PowerPoint용 Choice를 선택하며, 일반 도형과 그룹 내부에서도 같은 방식으로 처리합니다.

네이티브 Office Math 수식은 `texts.txt`에서 `[[EQ_1]]` 같은 보호 표시로 나타납니다. 번역 시 표시·순서·줄바꿈을 검증하고, PPTX에 반영할 때는 표시 대신 그 자리에 있던 수식 XML을 그대로 유지합니다. 수식 내부의 한글 설명(`m:t`)은 별도 `kind=math_text` 항목으로 추출하며 `math_text=번호`로 대응시켜 해당 글자만 번역합니다. 숫자·분수·위첨자 등 수식 구조는 재생성하지 않습니다.

같은 텍스트가 편집 가능한 Fallback 표·텍스트 상자에도 있으면 함께 교체합니다. 구버전 호환용 비트맵 Fallback과 이미지 속 글자는 다시 렌더링하거나 OCR하지 않으므로 원본 이미지로 남습니다. 네이티브 수식이 지원되는 PowerPoint에서 결과를 확인하세요.

수식 지원 전에 만든 `content.json`/`texts.txt`에는 수식 포함 상자가 빠져 있습니다. 수정된 분석기로 다시 추출하고 번역한 뒤 PPTX에 반영해야 합니다. `text_schema_version=2`로 새 형식을 구분하며, 이전 데이터로 수식 포함 원본에 적용하면 재추출 안내와 함께 중단합니다.

## 컴포넌트 파이프라인과 레이아웃 조정

전체 파이프라인:

```powershell
python -m legacy.pipeline "원본.pptx"
```

기존 추출·번역·적용 결과에서 Modify만 실행:

```powershell
python -m legacy.pipeline "원본.pptx" --start-at modify
```

단계 또는 검토 범위를 지정할 수도 있습니다.

```powershell
python -m legacy.pipeline "원본.pptx" --stop-after analyze
python -m legacy.pipeline "원본.pptx" --start-at translate --stop-after apply
python -m legacy.pipeline "원본.pptx" --start-at modify --pages 2,5,12 --max-passes 2
```

| 컴포넌트 | 담당 |
| --- | --- |
| `AnalysisStage` | Python으로 원본 분석, `texts.txt`·`content.json`·이미지 저장 |
| `TranslationStage` | 주입한 AI로 영어 번역, `translated.txt` 저장 |
| `ApplicationStage` | 원본 사본에 번역 적용, `translated.pptx` 생성 |
| `ModificationStage` | `PPTXLayoutModifier`를 호출해 `modified.pptx` 생성 |
| `PresentationPipeline` | 단계 순서·중간 재개·실패 상태 관리 |
| `AIProvider` | AI 공통 계약: `generate(prompt, model=None, images=())` |
| `GeminiCLIProvider` | 현재 동작하는 Antigravity/Gemini 어댑터 |
| `CodexAppServerProvider` | Codex App Server JSON-RPC로 이미지 첨부·검토 |
| `PresentationRenderer` | 렌더링 계약. 기본 구현은 `PowerPointRenderer` |

각 Stage는 `run(PipelineContext) -> dict` 계약을 따릅니다. 기존 스크립트도 계속 사용할 수 있으며, 파이프라인은 해당 기능을 컴포넌트로 감싸서 실행합니다.

```python
from pathlib import Path
from pptx_pipeline import (
    GeminiCLIProvider, CodexAppServerProvider, PowerPointRenderer, PPTXLayoutModifier,
    AnalysisStage, TranslationStage, ApplicationStage, ModificationStage,
    PipelineContext, PresentationPipeline,
)

translation_ai = GeminiCLIProvider(model="gemini-3.8-flash-medium", timeout=300)
layout_ai = CodexAppServerProvider(effort="low", timeout=300)
modifier = PPTXLayoutModifier(layout_ai, PowerPointRenderer())
flow = PresentationPipeline([
    AnalysisStage(),
    TranslationStage(translation_ai),
    ApplicationStage(),
    ModificationStage(modifier),
])
context = PipelineContext(Path("원본.pptx"), Path("output"), progress=print)
flow.run(context)  # 또는 start_at="modify"
```

### Modify 동작

1. Windows PowerPoint를 읽기 전용·문서 창 없이 열고 원본과 번역본의 PNG를 생성합니다. `TextRange2`로 실제 글자 영역과 상자 크기, 글꼴 크기도 측정합니다. 사용자가 열어 둔 프레젠테이션은 닫지 않습니다.
2. AI에 원본·번역 이미지, 페이지·도형 ID, 텍스트·좌표·글꼴·측정값을 제공합니다. Codex에는 `localImage`로 PNG를 직접 첨부합니다. Gemini를 선택하면 파일 보기 도구로 이미지를 읽으며 성공한 이미지 열람 이벤트가 없으면 수정하지 않습니다.
3. AI가 JSON 수정안을 반환합니다. Python이 수정 대상·최소 글자 크기·누적 축소율·그림 비율·이동 거리·슬라이드 경계를 검사합니다. 잘못된 수정안은 오류 이유를 전달해 한 번 재요청합니다.
4. 글꼴 크기 또는 허용된 도형 위치·크기만 수정합니다. 텍스트·수식 내용, 이미지 원본·자르기, 테마·마스터 등은 유지합니다. 그룹이나 회전 도형의 좌표 변경은 지원하지 않으며 해당 텍스트의 글자 크기는 조절할 수 있습니다. 표는 셀별 글자 크기를 조절합니다.
5. 수정 후 바뀐 페이지만 다시 PowerPoint로 렌더링하고 AI가 재검토합니다. 수정하지 않은 페이지의 통과 판정은 재사용합니다. 기본 2번까지 적용하며 마지막 검토 후에도 문제가 있으면 `needs_review`로 보고합니다. 시각 판단은 AI에 의존하므로 최종 발표 전 사람이 확인하는 것이 좋습니다.

`LayoutPolicy`로 최소 글자 크기(기본 10pt), 원래 글자 크기의 하한(65%), 그림 배율(65~110%), 원래 위치에서 이동 한도(3cm), 적용 횟수를 바꿀 수 있습니다. 원래 10pt 미만인 글자는 더 줄이지 않습니다. 제목·본문의 원래 서식과 수식 안의 상대 서식을 최대한 유지합니다.

PowerPoint가 설치된 Windows가 기본 요구사항입니다. Python `pywin32`는 필요하지 않으며 PowerShell COM으로 렌더링합니다. 다른 환경은 `PresentationRenderer` 계약을 구현한 렌더러를 주입하세요. 렌더러가 없거나 이미지 검증이 실패하면 좌표만 보고 임의 수정하는 대체 경로는 실행하지 않습니다.

기본 조합은 번역 `gemini-cli`의 `gemini-3.8-flash-medium`, 레이아웃 검토 `codex-app-server`의 reasoning `low`입니다. Codex 모델은 `model/list`에서 계정의 기본 이미지 지원 모델을 조회합니다. `--layout-model`로 고정하거나 `--layout-effort`로 reasoning 설정을 바꿀 수 있습니다. Gemini 검토를 선택할 경우 기본 모델은 `gemini-3.8-flash-low`입니다. 실제 속도는 CLI·서버 상태에 따라 달라집니다. `--model`은 두 단계 공통 모델 옵션이므로 서로 다른 backend를 사용하는 기본 구성에서는 `--translation-model`과 `--layout-model`을 각각 사용하세요.

### Codex 연결

현재 Codex에서 `codex mcp-server`는 제거되었습니다. 따라서 MCP 클라이언트를 흉내 내지 않고 공식 후속 인터페이스인 App Server의 stdio JSON-RPC를 구현했습니다. App Server는 실험적 인터페이스입니다. [공식 제거·이전 안내](https://learn.chatgpt.com/docs/mcp-server), [App Server 프로토콜](https://learn.chatgpt.com/docs/app-server).

`codex login`으로 로그인한 기존 계정을 사용합니다. Python 추가 의존성이나 별도 API 키는 필요하지 않습니다. 실행 파일은 PATH 또는 `CODEX_EXECUTABLE` 환경 변수에서 찾으며 `--codex-executable`로 실제 `codex.exe`를 지정할 수 있습니다. 로그인 정보와 사용자 전역 설정은 프로그램에서 수정하지 않습니다.

```powershell
codex login status
python -m legacy.pipeline "원본.pptx" --start-at modify
python -m legacy.pipeline "원본.pptx" --start-at modify --layout-provider codex-app-server --layout-effort low
# Gemini 검토로 변경
python -m legacy.pipeline "원본.pptx" --start-at modify --layout-provider gemini-cli
```

각 요청은 새로운 읽기 전용 임시 thread를 사용합니다. 도구 실행·파일 수정 대신 이미지와 JSON을 주고받으며 실제 PPT 수정은 Python이 담당합니다. 빈 응답은 다음 페이지로 진행하고 인증·프로토콜 오류는 진단과 함께 중단합니다. 요청별 전체 이벤트·응답·첨부 경로는 `output/codex-events/*.json`에 저장합니다. 이 기록에는 문서 내용이 들어갈 수 있습니다.

기본 AI 검토 요청은 페이지별로 실행하며 `--pages-per-request`로 묶는 페이지 수를 바꿀 수 있습니다. 빈 응답이면 해당 요청의 페이지를 미검토로 기록하고 다음 요청을 진행합니다. 검토하지 못한 페이지가 남으면 결과를 생성하되 `needs_review`(종료 코드 2)로 보고합니다. 번역에서 여러 페이지를 묶은 요청이 빈 응답이면 페이지별로 나누어 시도하며, 그래도 빈 응답인 텍스트는 원문을 유지하고 실패 페이지·항목 ID를 기록합니다. 인증·라이선스 오류, 잘못된 번역/수정안, 이미지 열람 검증 실패 등은 중단하며 `pipeline.json`과 가능한 경우 `failed-ai-response.jsonl`에 진단을 남깁니다.

### 결과와 상태

- `output/modified.pptx`: Modify 결과. 원본과 `translated.pptx`는 유지합니다.
- `output/texts.txt`, `output/content.json`, `output/resources/`: 추출 텍스트·위치·이미지.
- `output/translated.txt`, `output/translated.pptx`: 번역 텍스트와 레이아웃 검토 전 번역본.
- `output/translated.partial.txt`, `output/translated.progress.json`: 번역 배치별 중간 결과와 진행·실패 상태. 빈 응답으로 유지한 원문도 포함하므로 실패 항목 목록을 함께 확인하세요.
- `output/modified.layout.json`: 검토·미검토 페이지, 적용 이력, 최종 이미지 경로, `verified`/`needs_review` 상태.
- `output/layout/preview.pptx`: 각 검토 요청 완료 후 갱신되는 중간 PPT. 아직 재렌더링·재검토하지 않은 수정안을 포함할 수 있습니다. PowerPoint에서 파일을 열어 잠그면 갱신이 실패할 수 있으므로 사본을 열어 확인하세요.
- `output/layout/progress.json`: 현재 검토 회차·대기 페이지·실패 내역·중간 PPT 경로. 종료 시 최종 보고서로 갱신됩니다.
- `output/layout/review_회차_페이지.json`: 요청별 검토 결과와 수정안. 빈 응답 원본 이벤트는 `empty_회차_페이지.jsonl`에 저장합니다.
- `output/layout/reference/`, `before/`, `after_회차/`: 원본·수정 전·수정 후 슬라이드 PNG.
- `output/layout/candidate_회차.pptx`, `plan_회차.json`: 적용 회차별 PPT와 수정안.
- `output/pipeline.json`: 완료 단계 또는 실패한 단계·오류.
- `output/codex-events/`: Codex 요청별 원본 이벤트·최종 응답·이미지 경로·사용 모델.

`--pages`를 지정하면 지정 페이지에만 시각 검토를 적용하며 미검토 페이지를 보고서에 명시합니다. 기본은 전체입니다. CLI 종료 코드는 완료 `0`, 실행 실패 `1`, 결과는 만들었지만 추가 검토 필요 `2`입니다. `verified`는 요청 페이지에 대한 AI의 시각 검토 통과를 뜻하며 모든 상황에서의 완벽한 레이아웃을 보증하는 값은 아닙니다.

출력은 기본적으로 현재 작업 폴더의 `output/`에 저장하며 `-o 경로`로 변경합니다. 실행 시작 시 절대 경로를 출력합니다. `--start-at modify`는 기존 `translated.pptx`에서 검토를 다시 시작하며 저장된 검토 체크포인트를 자동으로 건너뛰지는 않습니다. 이전 실행의 파일이 남을 수 있으므로 현재 `pipeline.json`과 `layout/progress.json`의 상태·경로·수정 시각을 함께 확인하세요.

### AI 모델 확장

현재 등록된 실제 backend는 `gemini-cli`와 `codex-app-server`입니다. 별도 ChatGPT API 어댑터는 향후 추가할 대상이며, 등록되지 않은 backend를 선택하면 명확한 오류를 반환합니다. Stage와 Modifier는 특정 서비스에 의존하지 않습니다.

다른 모델을 추가하려면 `AIProvider.generate()` 계약을 구현하고 `ProviderRegistry.register(name, factory)`에 등록하거나, 생성한 객체를 `TranslationStage`와 `PPTXLayoutModifier`에 직접 주입합니다. 이미지 요청은 실제 이미지 바이트를 모델에 전달하거나 열람한 뒤 `AIResponse.images_read`에 해당 경로를 기록해야 합니다. `AIResponse.text`에는 모델의 최종 텍스트를 반환합니다(기존 번역기와 호환되는 `stdout` 별칭도 제공합니다). API 키·인증·CLI 호출 방식은 어댑터 내부에 둡니다.

CLI는 `--provider`, `--translation-provider`, `--layout-provider`와 `--model`, `--translation-model`, `--layout-model`을 지원합니다. 따라서 번역과 레이아웃 조정에서 서로 다른 backend나 모델을 선택할 수 있도록 구성되어 있습니다. 새 backend를 CLI에서 사용하려면 CLI가 만드는 registry에도 그 어댑터 factory를 등록하세요.

참고 문서: [PowerPoint PNG 내보내기](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export), [실제 글자 영역 측정](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.boundheight), [Antigravity headless](https://antigravity.google/docs/cli/headless/).

