# Configuration and MCP Reference

[Back to README](../README.md) · [한국어](reference_ko.md) ·
[CLI cheatsheet](cheatsheet.md)

## Configuration

The application loads `.env` from the repository root. Creating the file is
optional because normal defaults are built in:

```bash
cp .env.example .env
```

Use the AWS default credential chain whenever possible:

```bash
aws configure
export AWS_REGION=us-east-1
# export AWS_PROFILE=default
```

IAM roles, IAM Identity Center sessions, container credentials, and instance
profiles also work. `AWS_BEARER_TOKEN_BEDROCK` is an optional long-term API key
override.

### Mantle and Model Settings

| Variable | Default | Description |
|---|---|---|
| `AWS_REGION` | `us-east-1` | Bedrock Mantle Region |
| `AWS_PROFILE` | unset | Optional named AWS profile |
| `AWS_BEARER_TOKEN_BEDROCK` | unset | Optional long-term Bedrock API key |
| `MANTLE_MODEL_ID` | `openai.gpt-5.6-terra` | Primary translation model |
| `MANTLE_ENABLE_MODEL_FALLBACK` | `true` | Enable model fallback after retry exhaustion |
| `MANTLE_FALLBACK_MODEL_ID` | `openai.gpt-5.6-luna` | Model used for eligible fallback requests |
| `BEDROCK_MAX_RETRIES` | `5` | Maximum same-model retry attempts |
| `BEDROCK_MANTLE_TIMEOUT_SECONDS` | `300` | Provider request timeout |
| `BEDROCK_MANTLE_OPENAI_BASE_URL` | automatic | Optional OpenAI-compatible endpoint override |
| `OPENAI_REASONING_EFFORT` | `none` | GPT-5.6 reasoning effort |

`BEDROCK_MODEL_ID` remains a backward-compatible alias when
`MANTLE_MODEL_ID` is not set.

After same-model retries are exhausted, HTTP 429 rate-limit and HTTP 503
service-unavailable errors switch to `MANTLE_FALLBACK_MODEL_ID`. Quota or
billing-related 429 errors do not trigger fallback.

### Translation and Parallelism

| Variable | `.env.example` value | Description |
|---|---|---|
| `DEFAULT_TARGET_LANGUAGE` | `ko` | Default target language |
| `MAX_TOKENS` | `2000` | Maximum generated tokens per request |
| `TEMPERATURE` | `0.1` | GPT sampling temperature; omitted for Claude 5 models |
| `ENABLE_POLISHING` | `true` | Prefer natural, polished translations |
| `BATCH_SIZE` | `20` | Text items per model request |
| `CONTEXT_THRESHOLD` | `5` | Context-aware processing threshold |
| `BATCH_WORKERS` | `10` | Concurrent PowerPoint files in folder mode |
| `SLIDE_WORKERS` | `4` | Maximum parallel slide chunks per presentation |
| `SLIDES_PER_WORKER` | `30` | Maximum slides assigned to a chunk |

`BATCH_WORKERS` and `SLIDE_WORKERS` are independent. With default settings,
an 84-slide deck uses three slide workers for chunks `30+30+24`.

CLI flags override environment values:

```bash
uv run ppt-translate batch-translate samples/ -t ko \
  --workers 5 --slide-workers 4 --slides-per-worker 20
```

### Cache, Fonts, and Post-Processing

| Variable | Default | Description |
|---|---|---|
| `CACHE_BACKEND` | `sqlite` | `sqlite`, `memory`, or `none` |
| `CACHE_PATH` | `~/.ppt-translator/cache.db` | SQLite cache file |
| `FONT_KOREAN` | `맑은 고딕` | Korean output font |
| `FONT_JAPANESE` | `Yu Gothic UI` | Japanese output font |
| `FONT_ENGLISH` | `Amazon Ember` | English output font |
| `FONT_CHINESE` | `Microsoft YaHei` | Chinese output font |
| `FONT_DEFAULT` | `Arial` | Fallback font |
| `ENABLE_TEXT_AUTOFIT` | `true` | Enable output text auto-fit |
| `TEXT_LENGTH_THRESHOLD` | `10` | Minimum text length for auto-fit processing |
| `DEBUG` | `false` | Enable debug logging |

Clear the default cache after stopping translation processes:

```bash
rm -f ~/.ppt-translator/cache.db*
```

The cache key includes source text, source and target languages, model,
polishing mode, and glossary hash.

### Glossary

`./glossary.yaml` is detected automatically. A custom file can be supplied
with `--glossary` or the MCP `glossary_file` parameter.

```yaml
ko:
  "API": "API"
  "Foundation Model": "파운데이션 모델"
  "Observability": "Observability"
ja:
  "Cloud": "クラウド"
```

## Markdown Export

`export-markdown` creates one Markdown document from a presentation.
`batch-export-markdown` preserves the input folder structure and creates one
`.md` file per `.pptx`. Generated `translated_*` and `markdown_*` subfolders
are excluded when scanning a parent folder.

```bash
# Structured notes in Korean
uv run ppt-translate export-markdown presentation.pptx -l ko

# Source-preserving extraction with no model calls
uv run ppt-translate export-markdown presentation.pptx \
  --mode extract --language source

# Folder batch
uv run ppt-translate batch-export-markdown presentations/ -l ko \
  --workers 5 --chunk-workers 4 --slides-per-chunk 10
```

Structured mode produces:

- an executive summary and key themes with `[Slide N]` references
- decisions, action items, risks, and open questions when present
- slide-by-slide notes
- Markdown tables, chart labels and values, and speaker notes

Images are counted and marked as content that was not analyzed. OCR and
vision-based image interpretation are not performed.

Web verification is opt-in:

```bash
uv run ppt-translate export-markdown presentation.pptx \
  -l ko --web-verify --max-web-queries 3
```

Amazon Bedrock Mantle does not expose hosted server-side web search on its
OpenAI-compatible endpoint. The application therefore performs client-side
search through `ddgs`, passes bounded result snippets and URLs to the selected
Mantle model, and writes clickable source links into the Markdown. Search
content is treated as untrusted evidence. It is never followed as an
instruction.

Generated search queries are sent to external search backends used by `ddgs`.
Keep `--web-verify` disabled for confidential or sensitive presentations unless
that data handling is acceptable.

| Variable | Default | Description |
|---|---|---|
| `MARKDOWN_WORKERS` | `4` | Concurrent slide-summary chunks per presentation |
| `MARKDOWN_SLIDES_PER_CHUNK` | `10` | Slides per structured-summary request |
| `MARKDOWN_MAX_TOKENS` | `3000` | Maximum output tokens per slide chunk |
| `MARKDOWN_OVERVIEW_MAX_TOKENS` | `2000` | Maximum overview output tokens |
| `MARKDOWN_WEB_MAX_TOKENS` | `1500` | Maximum verification output tokens |
| `MARKDOWN_MAX_WEB_QUERIES` | `3` | Maximum client-side searches per presentation |
| `MARKDOWN_SEARCH_RESULTS_PER_QUERY` | `5` | Search results supplied per query |
| `MARKDOWN_REASONING_EFFORT` | `low` | GPT reasoning effort for Markdown synthesis |

Chunk and overview summaries use the standard cache. Web results and
verification sections are not cached because external information may change.

## Supported Models

| Model | Mantle model ID | Role |
|---|---|---|
| GPT-5.6 Sol | `openai.gpt-5.6-sol` | Highest GPT capability |
| GPT-5.6 Terra | `openai.gpt-5.6-terra` | Balanced default |
| GPT-5.6 Luna | `openai.gpt-5.6-luna` | Fast, economical fallback |
| GPT-5.6 Cyber | `openai.gpt-5.6-cyber` | Cybersecurity-specialized |
| Claude Opus 5 | `anthropic.claude-opus-5` | Highest Claude capability |
| Claude Sonnet 5 | `anthropic.claude-sonnet-5` | Balanced Claude model |
| Claude Haiku 4.5 | `anthropic.claude-haiku-4-5` | Fast, economical Claude model |

GPT models use the OpenAI Responses API with a short-term token generated from
AWS credentials. Claude models use the Anthropic Messages API with SigV4.

The translator accepts every language code in `Config.LANGUAGE_MAP`. Use
`list_supported_languages` through MCP to retrieve the complete current list.

## MCP Server

Start the server:

```bash
uv run mcp_server.py
```

### Shared Server Configuration

Replace `/path/to/ppt-translator` with the absolute repository path:

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

The MCP host must have access to the selected AWS credentials or IAM role.

### Claude Code

Project-scoped registration:

```bash
claude mcp add ppt-translator \
  --scope project \
  -- uv --project /path/to/ppt-translator \
  run /path/to/ppt-translator/mcp_server.py
```

User-scoped registration:

```bash
claude mcp add ppt-translator \
  --scope user \
  -e AWS_REGION=us-east-1 \
  -e AWS_PROFILE=default \
  -- uv --project /path/to/ppt-translator \
  run /path/to/ppt-translator/mcp_server.py
```

Restart Claude Code and run `/mcp` to verify the connection. The shared JSON
can alternatively be added to `.mcp.json` or `~/.claude.json`.

### Kiro

Add the shared JSON to:

- Kiro desktop: `~/.kiro/settings/mcp.json`
- Kiro CLI on macOS/Linux: `~/.aws/amazonq/mcp.json`
- Kiro CLI on Windows: `%APPDATA%\amazonq\mcp.json`

### Available MCP Tools

| Tool | Purpose | Important parameters |
|---|---|---|
| `translate_powerpoint` | Translate an entire presentation | `input_file`, `target_language`, `output_file`, `model_id`, `slide_workers`, `slides_per_worker` |
| `translate_specific_slides` | Translate selected slides or ranges | `input_file`, `slide_numbers`, `target_language`, `slide_workers`, `slides_per_worker` |
| `batch_translate_powerpoint` | Translate all presentations in a folder | `input_folder`, `output_folder`, `recursive`, `workers`, `slide_workers`, `slides_per_worker` |
| `export_powerpoint_markdown` | Export one presentation as Markdown | `input_file`, `output_language`, `mode`, `web_verify`, `workers`, `slides_per_chunk` |
| `batch_export_powerpoint_markdown` | Export a folder of presentations as Markdown | `input_folder`, `output_folder`, `recursive`, `workers`, `chunk_workers`, `slides_per_chunk` |
| `get_slide_info` | Return slide count and previews | `input_file` |
| `get_slide_preview` | Return detailed text for one slide | `input_file`, `slide_number` |
| `list_supported_languages` | List target language codes | none |
| `list_supported_models` | List Mantle model IDs | none |
| `get_translation_help` | Return MCP usage help | none |

Translation tools also accept common options including `enable_polishing`,
`glossary_file`, `cache_backend`, `dry_run`, `translate_charts`,
`source_language`, and `auto_detect_source`.

Example requests:

```text
Translate presentation.pptx to Korean
Translate slides 2-4 and 8 into Japanese
Batch-translate presentations/ into English, dry-run first
Export presentation.pptx as structured Korean Markdown
```

## Troubleshooting

### Credentials

```bash
aws sts get-caller-identity
```

If this fails, run `aws configure`, refresh the configured profile, or attach
an IAM role. A long-term key can be set through
`AWS_BEARER_TOKEN_BEDROCK`.

### Model Access

Confirm that the selected model is available in `AWS_REGION` and that the
active principal has Bedrock access.

### MCP Connection

- Verify every path in the MCP JSON is absolute and correct.
- Run `uv run mcp_server.py` directly before testing the host integration.
- Use `claude --debug` or the host's MCP logs when startup fails.

### PowerPoint Files

- Input files must be valid `.pptx` files.
- Ensure the process can read the input and write the output directory.
- Folder batch output must differ from the input folder.

## Development

Important modules:

| Path | Responsibility |
|---|---|
| `mcp_server.py` | FastMCP tools |
| `ppt_translator/cli.py` | CLI and folder workers |
| `ppt_translator/ppt_handler.py` | PowerPoint loading, translation, and application |
| `ppt_translator/translation_engine.py` | Prompts, cache, metrics, and batch translation |
| `ppt_translator/markdown_exporter.py` | Markdown collection, structuring, and web verification |
| `ppt_translator/bedrock_client.py` | Mantle provider adapter and model fallback |
| `ppt_translator/retry.py` | Retry classification and exponential backoff |
| `ppt_translator/cache.py` | SQLite, memory, and null cache implementations |
| `ppt_translator/progress.py` | Rich progress rendering |

Run tests and build:

```bash
uv run python -m unittest discover -s tests -v
uv build
```
