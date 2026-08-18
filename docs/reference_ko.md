# 구성 및 MCP 참조

[README로 돌아가기](../README_ko.md) · [English](reference.md) ·
[CLI 치트시트](cheatsheet.md)

## 구성

애플리케이션은 저장소 루트의 `.env`를 읽습니다. 기본 설정이 내장되어
있으므로 파일 생성은 선택 사항입니다:

```bash
cp .env.example .env
```

가능하면 AWS 기본 자격 증명 체인을 사용하세요:

```bash
aws configure
export AWS_REGION=us-east-1
# export AWS_PROFILE=default
```

IAM role, IAM Identity Center 세션, 컨테이너 자격 증명 및 인스턴스
프로파일도 사용할 수 있습니다. `AWS_BEARER_TOKEN_BEDROCK`은 선택적인
장기 API 키 재정의 설정입니다.

### Mantle 및 모델 설정

| 환경변수 | 기본값 | 설명 |
|---|---|---|
| `AWS_REGION` | `us-east-1` | Bedrock Mantle 리전 |
| `AWS_PROFILE` | 미설정 | 선택적 AWS 프로파일 |
| `AWS_BEARER_TOKEN_BEDROCK` | 미설정 | 선택적 Bedrock 장기 API 키 |
| `MANTLE_MODEL_ID` | `openai.gpt-5.6-terra` | 기본 번역 모델 |
| `MANTLE_ENABLE_MODEL_FALLBACK` | `true` | 재시도 소진 후 모델 fallback 활성화 |
| `MANTLE_FALLBACK_MODEL_ID` | `openai.gpt-5.6-luna` | 일시적 오류 시 fallback 모델 |
| `BEDROCK_MAX_RETRIES` | `5` | 동일 모델 최대 재시도 횟수 |
| `BEDROCK_MANTLE_TIMEOUT_SECONDS` | `300` | provider 요청 타임아웃 |
| `BEDROCK_MANTLE_OPENAI_BASE_URL` | 자동 | 선택적 OpenAI 호환 endpoint 재정의 |
| `OPENAI_REASONING_EFFORT` | `none` | GPT-5.6 reasoning effort |

`MANTLE_MODEL_ID`가 없으면 `BEDROCK_MODEL_ID`를 하위 호환 별칭으로
사용합니다.

동일 모델 재시도가 모두 실패하면 HTTP 429 rate limit과 HTTP 503 service
unavailable 오류에서 `MANTLE_FALLBACK_MODEL_ID`로 전환합니다. quota 또는
billing 관련 429 오류는 fallback 대상이 아닙니다.

### 번역 및 병렬 처리

| 환경변수 | `.env.example` 값 | 설명 |
|---|---|---|
| `DEFAULT_TARGET_LANGUAGE` | `ko` | 기본 대상 언어 |
| `MAX_TOKENS` | `2000` | 요청당 최대 출력 토큰 |
| `TEMPERATURE` | `0.1` | GPT sampling temperature. Claude 5에는 전달하지 않음 |
| `ENABLE_POLISHING` | `true` | 자연스러운 번역 다듬기 |
| `BATCH_SIZE` | `20` | 모델 요청당 텍스트 수 |
| `CONTEXT_THRESHOLD` | `5` | 문맥 기반 처리 임계값 |
| `BATCH_WORKERS` | `10` | 폴더 배치의 동시 PowerPoint 파일 수 |
| `SLIDE_WORKERS` | `4` | 프레젠테이션당 최대 병렬 슬라이드 청크 수 |
| `SLIDES_PER_WORKER` | `30` | 청크당 최대 슬라이드 수 |

`BATCH_WORKERS`와 `SLIDE_WORKERS`는 별개입니다. 기본 설정에서 84페이지
자료는 `30+30+24`의 세 슬라이드 워커로 처리됩니다.

CLI 옵션으로 환경변수를 재정의할 수 있습니다:

```bash
uv run ppt-translate batch-translate samples/ -t ko \
  --workers 5 --slide-workers 4 --slides-per-worker 20
```

### 캐시, 폰트 및 후처리

| 환경변수 | 기본값 | 설명 |
|---|---|---|
| `CACHE_BACKEND` | `sqlite` | `sqlite`, `memory` 또는 `none` |
| `CACHE_PATH` | `~/.ppt-translator/cache.db` | SQLite 캐시 파일 |
| `FONT_KOREAN` | `맑은 고딕` | 한국어 출력 폰트 |
| `FONT_JAPANESE` | `Yu Gothic UI` | 일본어 출력 폰트 |
| `FONT_ENGLISH` | `Amazon Ember` | 영어 출력 폰트 |
| `FONT_CHINESE` | `Microsoft YaHei` | 중국어 출력 폰트 |
| `FONT_DEFAULT` | `Arial` | fallback 폰트 |
| `ENABLE_TEXT_AUTOFIT` | `true` | 출력 텍스트 자동 맞춤 |
| `TEXT_LENGTH_THRESHOLD` | `10` | 자동 맞춤 처리 최소 텍스트 길이 |
| `DEBUG` | `false` | 디버그 로그 활성화 |

번역 프로세스를 종료한 뒤 기본 캐시를 삭제할 수 있습니다:

```bash
rm -f ~/.ppt-translator/cache.db*
```

캐시 키에는 원문, 원본 및 대상 언어, 모델, 폴리싱 여부와 용어집 해시가
포함됩니다.

### 용어집

`./glossary.yaml`은 자동 탐색됩니다. 다른 파일은 `--glossary` 또는 MCP
`glossary_file` 매개변수로 지정합니다.

```yaml
ko:
  "API": "API"
  "Foundation Model": "파운데이션 모델"
  "Observability": "Observability"
ja:
  "Cloud": "クラウド"
```

## Markdown 내보내기

`export-markdown`은 프레젠테이션 하나를 Markdown 문서로 만듭니다.
`batch-export-markdown`은 입력 폴더 구조를 유지하면서 `.pptx`마다
하나의 `.md` 파일을 생성합니다. 부모 폴더 검색 시 생성된
`translated_*`와 `markdown_*` 하위 폴더는 제외합니다.

```bash
# 구조화된 한국어 노트
uv run ppt-translate export-markdown presentation.pptx -l ko

# 모델 호출 없이 원문 추출
uv run ppt-translate export-markdown presentation.pptx \
  --mode extract --language source

# 폴더 일괄 처리
uv run ppt-translate batch-export-markdown presentations/ -l ko \
  --workers 5 --chunk-workers 4 --slides-per-chunk 10
```

구조화 모드는 다음 내용을 만듭니다.

- `[Slide N]` 출처가 포함된 핵심 요약과 주요 주제
- 자료에 실제로 존재하는 결정 사항, 실행 항목, 위험 및 미해결 질문
- 슬라이드별 구조화 노트
- Markdown 표, 차트 레이블과 값, 발표자 노트

이미지는 개수를 표시하고 분석되지 않은 콘텐츠로 명시합니다. OCR 또는
비전 기반 이미지 해석은 수행하지 않습니다.

웹 검증은 명시적으로 켜야 합니다.

```bash
uv run ppt-translate export-markdown presentation.pptx \
  -l ko --web-verify --max-web-queries 3
```

Amazon Bedrock Mantle의 OpenAI 호환 endpoint는 서버 측 내장 웹검색을
제공하지 않습니다. 따라서 애플리케이션이 `ddgs`로 제한된 검색 결과와
URL을 수집하고, 이를 선택한 Mantle 모델에 전달한 뒤 클릭 가능한 출처를
Markdown에 기록합니다. 검색 내용은 신뢰할 수 없는 증거로 취급하며
명령으로 실행하지 않습니다.

생성된 검색어는 `ddgs`가 사용하는 외부 검색 backend로 전달됩니다. 기밀
또는 민감한 프레젠테이션은 해당 데이터 처리가 허용된 경우가 아니라면
`--web-verify`를 사용하지 마세요.

| 환경변수 | 기본값 | 설명 |
|---|---|---|
| `MARKDOWN_WORKERS` | `4` | 프레젠테이션당 동시 슬라이드 요약 청크 |
| `MARKDOWN_SLIDES_PER_CHUNK` | `10` | 구조화 요약 요청당 슬라이드 수 |
| `MARKDOWN_MAX_TOKENS` | `3000` | 슬라이드 청크 최대 출력 토큰 |
| `MARKDOWN_OVERVIEW_MAX_TOKENS` | `2000` | 전체 개요 최대 출력 토큰 |
| `MARKDOWN_WEB_MAX_TOKENS` | `1500` | 웹 검증 최대 출력 토큰 |
| `MARKDOWN_MAX_WEB_QUERIES` | `3` | 프레젠테이션당 최대 검색 횟수 |
| `MARKDOWN_SEARCH_RESULTS_PER_QUERY` | `5` | 검색어당 모델에 전달할 결과 수 |
| `MARKDOWN_REASONING_EFFORT` | `low` | Markdown 구조화에 사용할 GPT reasoning effort |

슬라이드 청크와 전체 개요는 기존 캐시를 사용합니다. 외부 정보는 바뀔 수
있으므로 웹 검색 결과와 검증 섹션은 캐시하지 않습니다.

## 지원 모델

| 모델 | Mantle 모델 ID | 용도 |
|---|---|---|
| GPT-5.6 Sol | `openai.gpt-5.6-sol` | GPT 최고 성능 |
| GPT-5.6 Terra | `openai.gpt-5.6-terra` | 균형형 기본 모델 |
| GPT-5.6 Luna | `openai.gpt-5.6-luna` | 빠르고 경제적인 fallback |
| GPT-5.6 Cyber | `openai.gpt-5.6-cyber` | 사이버 보안 특화 |
| Claude Opus 5 | `anthropic.claude-opus-5` | Claude 최고 성능 |
| Claude Sonnet 5 | `anthropic.claude-sonnet-5` | 균형형 Claude 모델 |
| Claude Haiku 4.5 | `anthropic.claude-haiku-4-5` | 빠르고 경제적인 Claude 모델 |

GPT 모델은 AWS 자격 증명으로 만든 단기 토큰과 OpenAI Responses API를
사용합니다. Claude 모델은 Anthropic Messages API와 SigV4를 사용합니다.

`Config.LANGUAGE_MAP`에 등록된 모든 언어 코드를 지원합니다. MCP의
`list_supported_languages` 도구로 현재 전체 목록을 확인할 수 있습니다.

## MCP 서버

서버를 시작합니다:

```bash
uv run mcp_server.py
```

### 공통 서버 설정

`/path/to/ppt-translator`를 저장소의 절대 경로로 바꾸세요:

```json
{
  "mcpServers": {
    "ppt-translator": {
      "command": "uv",
      "args": [
        "--project",
        "/path/to/ppt-translator",
        "run",
        "/path/to/ppt-translator/mcp_server.py"
      ],
      "env": {
        "AWS_REGION": "us-east-1",
        "AWS_PROFILE": "default",
        "MANTLE_MODEL_ID": "openai.gpt-5.6-terra"
      }
    }
  }
}
```

MCP 호스트에서 선택한 AWS 자격 증명 또는 IAM role에 접근할 수 있어야
합니다.

### Claude Code

프로젝트 범위 등록:

```bash
claude mcp add ppt-translator \
  --scope project \
  -- uv --project /path/to/ppt-translator \
  run /path/to/ppt-translator/mcp_server.py
```

사용자 범위 등록:

```bash
claude mcp add ppt-translator \
  --scope user \
  -e AWS_REGION=us-east-1 \
  -e AWS_PROFILE=default \
  -- uv --project /path/to/ppt-translator \
  run /path/to/ppt-translator/mcp_server.py
```

Claude Code를 재시작하고 `/mcp`로 연결을 확인합니다. 공통 JSON을
`.mcp.json` 또는 `~/.claude.json`에 직접 추가해도 됩니다.

### Kiro

공통 JSON을 다음 파일에 추가합니다:

- Kiro 데스크톱: `~/.kiro/settings/mcp.json`
- macOS/Linux Kiro CLI: `~/.aws/amazonq/mcp.json`
- Windows Kiro CLI: `%APPDATA%\amazonq\mcp.json`

### 사용 가능한 MCP 도구

| 도구 | 용도 | 주요 매개변수 |
|---|---|---|
| `translate_powerpoint` | 전체 프레젠테이션 번역 | `input_file`, `target_language`, `output_file`, `model_id`, `slide_workers`, `slides_per_worker` |
| `translate_specific_slides` | 지정한 슬라이드 또는 범위 번역 | `input_file`, `slide_numbers`, `target_language`, `slide_workers`, `slides_per_worker` |
| `batch_translate_powerpoint` | 폴더 내 프레젠테이션 일괄 번역 | `input_folder`, `output_folder`, `recursive`, `workers`, `slide_workers`, `slides_per_worker` |
| `export_powerpoint_markdown` | 프레젠테이션 하나를 Markdown으로 정리 | `input_file`, `output_language`, `mode`, `web_verify`, `workers`, `slides_per_chunk` |
| `batch_export_powerpoint_markdown` | 폴더 내 프레젠테이션을 Markdown으로 일괄 정리 | `input_folder`, `output_folder`, `recursive`, `workers`, `chunk_workers`, `slides_per_chunk` |
| `get_slide_info` | 슬라이드 수와 미리보기 반환 | `input_file` |
| `get_slide_preview` | 특정 슬라이드의 상세 텍스트 반환 | `input_file`, `slide_number` |
| `list_supported_languages` | 대상 언어 코드 목록 | 없음 |
| `list_supported_models` | Mantle 모델 ID 목록 | 없음 |
| `get_translation_help` | MCP 사용 도움말 | 없음 |

번역 도구는 공통으로 `enable_polishing`, `glossary_file`,
`cache_backend`, `dry_run`, `translate_charts`, `source_language`,
`auto_detect_source` 옵션도 지원합니다.

요청 예시:

```text
presentation.pptx를 한국어로 번역해줘
2-4번과 8번 슬라이드를 일본어로 번역해줘
presentations/ 폴더를 영어로 번역하되 dry-run부터 실행해줘
presentation.pptx를 구조화된 한국어 Markdown으로 정리해줘
```

## 문제 해결

### 자격 증명

```bash
aws sts get-caller-identity
```

실패하면 `aws configure`를 실행하거나 프로파일을 갱신하고, 또는 IAM
role을 연결하세요. 장기 키는 `AWS_BEARER_TOKEN_BEDROCK`으로 설정할 수
있습니다.

### 모델 액세스

선택한 모델이 `AWS_REGION`에서 제공되고 현재 principal에 Bedrock 접근
권한이 있는지 확인하세요.

### MCP 연결

- MCP JSON의 모든 경로가 올바른 절대 경로인지 확인합니다.
- 호스트에 연결하기 전에 `uv run mcp_server.py`를 직접 실행합니다.
- 시작 실패 시 `claude --debug` 또는 해당 호스트의 MCP 로그를 확인합니다.

### PowerPoint 파일

- 입력 파일은 유효한 `.pptx` 형식이어야 합니다.
- 입력 파일을 읽고 출력 폴더에 쓸 권한이 있어야 합니다.
- 폴더 배치의 출력 폴더는 입력 폴더와 달라야 합니다.

## 개발

주요 모듈:

| 경로 | 역할 |
|---|---|
| `mcp_server.py` | FastMCP 도구 |
| `ppt_translator/cli.py` | CLI 및 폴더 워커 |
| `ppt_translator/ppt_handler.py` | PowerPoint 로드, 번역 및 결과 적용 |
| `ppt_translator/markdown_exporter.py` | Markdown 수집, 구조화 및 웹 검증 |
| `ppt_translator/translation_engine.py` | 프롬프트, 캐시, 메트릭 및 배치 번역 |
| `ppt_translator/bedrock_client.py` | Mantle provider 어댑터 및 모델 fallback |
| `ppt_translator/retry.py` | 재시도 분류와 exponential backoff |
| `ppt_translator/cache.py` | SQLite, 메모리 및 null 캐시 |
| `ppt_translator/progress.py` | Rich 진행률 표시 |

테스트 및 빌드:

```bash
uv run python -m unittest discover -s tests -v
uv build
```
