"""
Source language detection via Bedrock Mantle (1-shot).

We ask the same Mantle model we're about to translate with to classify
a short sample of the document text. One call per presentation, cheap,
and dramatically better than running langdetect on shapes that say "API"
or "2024".

Returns an ISO 639-1 code like 'en', 'ko', 'ja', 'zh'. Falls back to 'en'
on any error (including empty input, throttling, unknown response) so the
pipeline keeps working even when detection goes wrong.
"""
import logging
import re
from typing import List, Optional

from .bedrock_client import BedrockClient
from .config import Config

logger = logging.getLogger(__name__)

_MAX_SAMPLE_CHARS = 1500
_MAX_OUTPUT_TOKENS = 16
_SYSTEM_PROMPT = (
    "You are a language identifier. Given a text sample, respond with ONLY "
    "the ISO 639-1 two-letter code of its dominant language (e.g., en, ko, "
    "ja, zh, es, fr). No explanation, no punctuation, no quotes — just two "
    "lowercase letters."
)


def _sample_text(texts: List[str], max_chars: int = _MAX_SAMPLE_CHARS) -> str:
    """Pack a few representative strings into one short sample for the model.

    We prefer longer strings — single-token items like "API" or "2024" tell
    the detector nothing useful, so we sort by length and stop once we have
    enough chars to disambiguate.
    """
    # Dedup while preserving order, then favor longer items.
    seen = set()
    uniq = []
    for t in texts:
        t = (t or '').strip()
        if not t or t in seen:
            continue
        seen.add(t)
        uniq.append(t)
    uniq.sort(key=len, reverse=True)

    buf: List[str] = []
    total = 0
    for t in uniq:
        if total >= max_chars:
            break
        snippet = t[: max(0, max_chars - total)]
        buf.append(snippet)
        total += len(snippet) + 1  # +1 for newline
    return "\n".join(buf)


def detect_language(texts: List[str],
                    model_id: str = Config.DEFAULT_MODEL_ID,
                    bedrock: Optional[BedrockClient] = None) -> str:
    """Detect the dominant language of `texts`. Returns an ISO 639-1 code.

    `bedrock` may be injected to reuse an existing client (and share its
    retry/backoff config). If omitted, a fresh BedrockClient is created.
    """
    if not texts:
        return 'en'

    sample = _sample_text(texts)
    if not sample:
        return 'en'

    client = bedrock or BedrockClient()

    try:
        response = client.converse(
            modelId=model_id,
            system=[{"text": _SYSTEM_PROMPT}],
            messages=[{
                "role": "user",
                "content": [{"text": f"Sample:\n{sample}"}],
            }],
            inferenceConfig={
                "maxTokens": _MAX_OUTPUT_TOKENS,
                "temperature": 0.0,  # deterministic
            },
        )
        raw = (response.get('output', {}).get('message', {})
                       .get('content', [{}])[0].get('text', '')).strip().lower()
    except Exception as e:
        logger.warning(f"Language detection failed, defaulting to 'en': {e}")
        return 'en'

    # Model occasionally adds punctuation or explanations despite the prompt.
    # Pull out the first two-letter token.
    match = re.search(r'\b([a-z]{2})\b', raw)
    if match:
        code = match.group(1)
    else:
        code = raw[:2]

    if len(code) != 2 or not code.isalpha():
        logger.warning(f"Unrecognized language code '{raw}', defaulting to 'en'")
        return 'en'

    # Accept anything the model returns — Config.LANGUAGE_MAP has 90+ codes
    # and we don't want false negatives for less common ones. Downstream code
    # already tolerates unknown codes.
    logger.info(f"🌐 Detected source language: {code}")
    return code
