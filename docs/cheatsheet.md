# PPT Translator Cheatsheet

자주 쓰는 CLI 옵션 모음. 기본 모델은 **GPT-5.6 Terra**
(`openai.gpt-5.6-terra`). 모든 모델 호출은 Amazon Bedrock Mantle을 사용합니다.
캐시는 기본 ON (SQLite,
`~/.ppt-translator/cache.db`). 용어집은 `./glossary.yaml` 자동 탐색.

---

## AWS 인증

```bash
# 기존 AWS CLI 프로파일 사용
aws configure
export AWS_REGION=us-east-1
export AWS_PROFILE=default

# 선택 사항: AWS 자격 증명 대신 Bedrock 장기 API 키 사용
# export AWS_BEARER_TOKEN_BEDROCK=your_bedrock_api_key
```

GPT 모델의 단기 토큰은 자동 생성되며 Claude 모델은 SigV4를 사용합니다.

## 단일 파일 번역

```bash
# 가장 기본
uv run ppt-translate translate 파일.pptx -t ko

# 출력 경로 지정
uv run ppt-translate translate 파일.pptx -t ko -o 번역본.pptx

# 30페이지 초과 시 기본 병렬 처리 (최대 4워커, 워커당 30페이지)
uv run ppt-translate translate 대형자료.pptx -t ko

# 워커당 20페이지로 분할 / 병렬 처리 끄기
uv run ppt-translate translate 대형자료.pptx -t ko --slides-per-worker 20
uv run ppt-translate translate 대형자료.pptx -t ko --slide-workers 1

# 특정 슬라이드만 (쉼표/범위/혼합)
uv run ppt-translate translate-slides 파일.pptx -s "1,3,5"  -t ko
uv run ppt-translate translate-slides 파일.pptx -s "2-4"    -t ko
uv run ppt-translate translate-slides 파일.pptx -s "1,3-5"  -t ko

# 슬라이드 개수·텍스트 미리보기
uv run ppt-translate info 파일.pptx
```

## 폴더 일괄 번역

```bash
# 폴더 안 모든 .pptx 병렬 번역 (하위 폴더 포함, 파일 워커 10개 기본)
uv run ppt-translate batch-translate samples/ -t ko

# 최상위 폴더만 (하위 폴더 제외)
uv run ppt-translate batch-translate samples/ -t ko --no-recursive

# 동시 PPT 파일 수 / PPT당 슬라이드 워커 수 / 출력 폴더 지정
uv run ppt-translate batch-translate samples/ -t ko -w 2 --slide-workers 4
uv run ppt-translate batch-translate samples/ -t ko -w 8 -o translated_ko/
```

파일별 슬라이드 진행바 + 전체 파일 진행바가 함께 표시됩니다. 슬라이드
워커는 각 PPT에 독립적으로 적용됩니다. 기본 30페이지 청크이므로 84페이지
PPT는 30/30/24페이지의 3개 워커로 번역됩니다.

## 비용 미리 확인 (API 호출 없음)

```bash
# 단일
uv run ppt-translate translate 파일.pptx -t ko --dry-run

# 배치 집계
uv run ppt-translate batch-translate samples/ -t ko --dry-run
```

## Markdown Export

```bash
# Structured Markdown notes
uv run ppt-translate export-markdown samples/en.pptx -l ko

# Raw extraction without model calls
uv run ppt-translate export-markdown samples/en.pptx --mode extract -l source

# Optional web verification
uv run ppt-translate export-markdown samples/en.pptx -l ko --web-verify

# Batch export
uv run ppt-translate batch-export-markdown samples/ -l ko

# Batch concurrency and chunk size
uv run ppt-translate batch-export-markdown samples/ -l ko \
  -w 5 --chunk-workers 4 --slides-per-chunk 10
```

출력 예: 원본 언어·슬라이드 수·번역 대상 문자·예상 토큰·예상 비용.

## 모델 바꾸기

```bash
# GPT-5.6 Terra (기본, 균형형)
uv run ppt-translate translate 파일.pptx -t ko -m openai.gpt-5.6-terra

# GPT-5.6 Sol (최고 성능)
uv run ppt-translate translate 파일.pptx -t ko -m openai.gpt-5.6-sol

# GPT-5.6 Luna (저비용 대량 처리)
uv run ppt-translate translate 파일.pptx -t ko -m openai.gpt-5.6-luna

# Claude Sonnet 5 (균형형 Claude)
uv run ppt-translate translate 파일.pptx -t ko -m anthropic.claude-sonnet-5

# Claude Opus 5 / Haiku 4.5
uv run ppt-translate translate 파일.pptx -t ko -m anthropic.claude-opus-5
uv run ppt-translate translate 파일.pptx -t ko -m anthropic.claude-haiku-4-5
```

전체 모델 ID 표는 [구성 및 MCP 참조](reference_ko.md#지원-모델) 참고.

## 원본 언어

```bash
# 기본: 1회 LLM 호출로 자동 감지
uv run ppt-translate translate 파일.pptx -t ko

# 명시 (감지 스킵 → 빠름)
uv run ppt-translate translate 파일.pptx --source-language en -t ko

# 자동 감지 끄기 (예전처럼 모델이 문맥 추론)
uv run ppt-translate translate 파일.pptx -t ko --no-detect-source
```

원본 == 대상이면 Bedrock 호출을 전혀 하지 않습니다.

## 용어집 (YAML)

```bash
# ./glossary.yaml 자동 사용
uv run ppt-translate translate 파일.pptx -t ko

# 명시
uv run ppt-translate translate 파일.pptx -t ko -g my-glossary.yaml
```

`glossary.yaml` 형식:

```yaml
ko:
  "API": "API"                    # src == tgt → 번역 안 함
  "Foundation Model": "파운데이션 모델"
  "Observability": "Observability"
ja:
  "Cloud": "クラウド"
```

## 캐시 제어

```bash
# 기본: SQLite (~/.ppt-translator/cache.db)
uv run ppt-translate translate 파일.pptx -t ko

# 메모리 캐시 (프로세스 내, 디스크 안 씀)
uv run ppt-translate translate 파일.pptx -t ko --cache-backend memory

# 캐시 완전 비활성
uv run ppt-translate translate 파일.pptx -t ko --no-cache

# 커스텀 경로
uv run ppt-translate translate 파일.pptx -t ko --cache-path /tmp/mycache.db

# 캐시 비우기
rm ~/.ppt-translator/cache.db
```

캐시 키: `(원본텍스트, 대상언어, 원본언어, 모델ID, 폴리싱, 용어집해시)` 중 하나라도 바뀌면 miss.

## 기타 자주 쓰는 플래그

```bash
# 폴리싱(자연스러운 문장) 끄기 — 더 literal한 번역
uv run ppt-translate translate 파일.pptx -t ko --no-polishing

# 차트 번역 건너뛰기
uv run ppt-translate translate 파일.pptx -t ko --no-charts
```

## 언어 코드 예시

| 언어 | 코드 | 언어 | 코드 |
|---|---|---|---|
| 한국어 | `ko` | 영어 | `en` |
| 일본어 | `ja` | 중국어(간체) | `zh` |
| 스페인어 | `es` | 프랑스어 | `fr` |
| 독일어 | `de` | 이탈리아어 | `it` |
| 포르투갈어 | `pt` | 러시아어 | `ru` |
| 아랍어 | `ar` | 힌디어 | `hi` |

90+ 언어 지원 (`Config.LANGUAGE_MAP` 참조).

## Help

```bash
uv run ppt-translate --help
uv run ppt-translate translate --help
uv run ppt-translate translate-slides --help
uv run ppt-translate batch-translate --help
uv run ppt-translate info --help
```
