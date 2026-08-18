# PowerPoint Translator using Amazon Bedrock

Amazon Bedrock Mantle을 통해 OpenAI 및 Anthropic 모델을 호출하면서
PowerPoint의 서식과 구조를 보존하는 번역 도구입니다. CLI 또는 FastMCP
서버로 사용할 수 있습니다.

**현재 릴리스: v1.1.0**

[English](README.md) · [구성 및 MCP 참조](docs/reference_ko.md) ·
[CLI 치트시트](docs/cheatsheet.md)

## 주요 기능

- GPT-5.6 Sol/Terra/Luna/Cyber 및 Claude Opus 5/Sonnet 5/Haiku 4.5 지원
- 기본 모델 GPT-5.6 Terra와 선택적 Luna fallback
- AWS 기본 자격 증명 체인, Mantle 단기 토큰 및 SigV4 지원
- 서식, 레이아웃, 색상, 언어별 폰트, 노트 및 차트 보존
- 대형 프레젠테이션과 폴더 배치의 병렬 번역
- 선택적 웹 검증을 지원하는 PowerPoint 구조화 Markdown 내보내기
- SQLite 또는 메모리 번역 캐시
- 원본 언어 감지, YAML 용어집 및 비용 dry-run
- 일시적 오류 자동 재시도와 429/503 발생 시 모델 fallback
- CLI 및 FastMCP 인터페이스

## 예제

복잡한 슬라이드 레이아웃을 유지하며 번역합니다:

<table>
<tr>
<td><img src="imgs/original-en-complex.png" alt="영어 원본" width="450"/></td>
<td><img src="imgs/translated-ko-complex.png" alt="한국어 번역본" width="450"/></td>
</tr>
<tr>
<td align="center"><em>영어 원본 슬라이드</em></td>
<td align="center"><em>레이아웃이 보존된 한국어 번역본</em></td>
</tr>
</table>

### Claude Code MCP

<table>
<tr>
<td><img src="imgs/mcp-cc1.png" alt="MCP 연결" width="450"/></td>
<td><img src="imgs/mcp-cc2.png" alt="MCP 번역" width="450"/></td>
</tr>
<tr>
<td align="center"><em>MCP 연결 확인</em></td>
<td align="center"><em>MCP를 통한 번역</em></td>
</tr>
</table>

## 빠른 시작

### 요구사항

- Python 3.11 이상
- Amazon Bedrock 액세스 권한이 있는 AWS 계정
- AWS 기본 자격 증명 체인에서 사용할 수 있는 자격 증명
- 선택한 AWS 리전의 모델 액세스 권한

### 설치

```bash
git clone https://github.com/daekeun-ml/ppt-translator
cd ppt-translator
uv sync
```

AWS 자격 증명을 설정합니다:

```bash
aws configure
export AWS_REGION=us-east-1
# export AWS_PROFILE=default
```

GPT 모델은 단기 토큰을 자동 생성하고 Claude 모델은 SigV4를 사용합니다.
`AWS_BEARER_TOKEN_BEDROCK`은 선택적인 장기 API 키 재정의 설정입니다.

일반적인 사용에는 기본 설정만으로 충분합니다. 설정을 바꾸려면:

```bash
cp .env.example .env
```

전체 설정은 [구성](docs/reference_ko.md#구성)을 참고하세요. README에는
`.env` 전문을 중복해서 표시하지 않습니다.

## CLI 사용법

### 전체 프레젠테이션 번역

```bash
uv run ppt-translate translate samples/en.pptx -t ko
```

30페이지를 초과하는 프레젠테이션은 여러 청크로 나누어 병렬 처리합니다:

```bash
# 기본값: 최대 4개 슬라이드 워커, 청크당 30페이지
uv run ppt-translate translate large-deck.pptx -t ko

# 80페이지 -> 20+20+20+20
uv run ppt-translate translate large-deck.pptx -t ko \
  --slide-workers 4 --slides-per-worker 20
```

![단일 프레젠테이션 병렬 번역](imgs/standalone.png)

### 폴더 내 모든 PPT 파일 일괄 번역

```bash
# 기본적으로 하위 폴더까지 재귀 처리
uv run ppt-translate batch-translate samples/ -t ko

# 최상위 폴더만 처리
uv run ppt-translate batch-translate samples/ -t ko --no-recursive

# 출력 폴더와 동시 처리할 PPT 파일 수 지정
uv run ppt-translate batch-translate samples/ -t ko -o translated_ko/ -w 10
```

`--workers`는 동시에 처리할 PowerPoint 파일 수를, `--slide-workers`는 각
프레젠테이션 내부의 병렬 슬라이드 청크 수를 제어합니다.

![폴더 병렬 번역](imgs/batch-translate.png)

### 특정 슬라이드 번역

```bash
uv run ppt-translate translate-slides samples/en.pptx -s "1,3,5" -t ko
uv run ppt-translate translate-slides samples/en.pptx -s "2-4" -t ko
```

### PowerPoint를 Markdown으로 정리

```bash
# 슬라이드 출처가 표시된 AI 구조화 한국어 노트
uv run ppt-translate export-markdown samples/en.pptx -l ko

# 모델 호출 없이 원문을 그대로 추출
uv run ppt-translate export-markdown samples/en.pptx \
  --mode extract --language source

# 폴더 구조를 유지하며 모든 PPT를 일괄 변환
uv run ppt-translate batch-export-markdown samples/ -l ko

# 필요한 외부 주장만 제한적으로 웹 검증
uv run ppt-translate export-markdown samples/en.pptx -l ko --web-verify
```

구조화 모드는 슬라이드를 병렬 청크로 요약하고, 슬라이드 출처가 포함된 전체 개요를 만들며 표·차트 데이터·발표자 노트를 보존합니다. 각 청크가 요청한 출력 언어인지 검증하며, 언어가 다르거나 슬라이드가 누락되면 자동으로 재생성하고 필요하면 더 작은 청크로 다시 분할합니다. 대상 언어가 지정된 구조화 문서에 원문을 fallback으로 섞지 않습니다. 웹 검증은 명시적으로 켠 경우에만 실행됩니다. 자세한 내용은 [Markdown 내보내기 참조](docs/reference_ko.md#markdown-내보내기)를 참고하세요.

### Dry-Run 및 캐시

```bash
# 모델을 호출하지 않고 토큰과 비용 추정
uv run ppt-translate translate samples/en.pptx -t ko --dry-run

# 캐시는 기본적으로 ~/.ppt-translator/cache.db에 저장
uv run ppt-translate translate samples/en.pptx -t ko

# 캐시 비활성화
uv run ppt-translate translate samples/en.pptx -t ko --no-cache
```

![Dry-run 비용 추정](imgs/dry-run.png)

다음 화면은 모든 번역이 캐시된 폴더 실행과 모델을 실제로 호출하는 첫
실행을 비교합니다:

![배치 캐시 비교](imgs/batch-translate-cache.png)

### 기타 주요 옵션

```bash
# 원본 언어를 지정하면 자동 감지 생략
uv run ppt-translate translate samples/en.pptx --source-language en -t ko

# ./glossary.yaml은 자동 탐색되며 직접 지정할 수도 있음
uv run ppt-translate translate samples/en.pptx -t ko -g glossary.yaml

# 차트 텍스트 번역 제외
uv run ppt-translate translate samples/en.pptx -t ko --no-charts

# 슬라이드 정보 확인
uv run ppt-translate info samples/en.pptx
```

일상적으로 사용하는 전체 명령은 [CLI 치트시트](docs/cheatsheet.md)를
참고하세요.

## MCP 서버

FastMCP 서버를 시작합니다:

```bash
uv run mcp_server.py
```

연결된 AI 어시스턴트에 자연어로 요청할 수 있습니다:

```text
samples/en.pptx를 한국어로 번역해줘
samples/ 폴더를 일본어로 일괄 번역하되 dry-run부터 실행해줘
3번 슬라이드 내용을 보여줘
samples/en.pptx를 구조화된 한국어 Markdown으로 정리해줘
```

Claude Code와 Kiro 설정 및 전체 MCP 도구 목록은
[구성 및 MCP 참조](docs/reference_ko.md#mcp-서버)를 참고하세요.

## 문서

- [구성 및 MCP 참조](docs/reference_ko.md)
- [CLI 치트시트](docs/cheatsheet.md)
- [후처리 예제](docs/post_processing_examples.md)
- [PowerPoint 처리 구조](docs/ppt_handler_structure.md)

## 라이선스

이 프로젝트는 MIT 라이선스를 따릅니다. [LICENSE](LICENSE)를 참고하세요.
