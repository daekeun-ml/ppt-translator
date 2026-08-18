"""
Rough token and cost estimation for dry-run reports.

These are approximations — actual Bedrock usage depends on tokenizer specifics,
prompt overhead, and output length. Enough to give the user a ballpark before
they spend money, not a quote.

Pricing table is kept small on purpose: we list the models we see most often
and show "pricing unavailable" for everything else rather than guessing.
Update as AWS pricing changes.
"""
from typing import Dict, Tuple

# (input $/1M tokens, output $/1M tokens)
MODEL_PRICING: Dict[str, Tuple[float, float]] = {
    # Bedrock Mantle in-region, standard tier, USD per 1M tokens.
    # GPT-5.6 values use the short-context rate (up to 272K tokens).
    'openai.gpt-5.6-sol': (5.50, 33.00),
    'openai.gpt-5.6-terra': (2.20, 13.20),
    'openai.gpt-5.6-luna': (0.22, 1.32),
    'openai.gpt-5.6-cyber': (13.75, 82.50),
    'anthropic.claude-opus-5': (5.00, 25.00),
    'anthropic.claude-sonnet-5': (2.00, 10.00),
    'anthropic.claude-haiku-4-5': (1.00, 5.00),
}


# Rough chars-per-token by language family. CJK scripts pack fewer chars per
# token than Latin scripts, so we use a lower divisor to stay conservative.
_CHARS_PER_TOKEN = {
    'en': 4.0, 'fr': 4.0, 'de': 4.0, 'es': 4.0, 'it': 4.0, 'pt': 4.0,
    'ru': 3.5, 'ar': 3.0,
    'ko': 1.5, 'ja': 1.5, 'zh': 1.5,
}


def estimate_tokens(total_chars: int, lang: str = 'en') -> int:
    """Rough char-count → token-count estimate for the given language."""
    if total_chars <= 0:
        return 0
    base = lang.split('-')[0].lower() if lang else 'en'
    divisor = _CHARS_PER_TOKEN.get(base, 4.0)
    return max(1, int(total_chars / divisor))


def estimate_cost(input_tokens: int, output_tokens: int, model_id: str) -> float:
    """Estimated USD cost for a translation run. Returns 0.0 for unknown models."""
    pricing = MODEL_PRICING.get(model_id)
    if pricing is None:
        return 0.0
    p_in, p_out = pricing
    return (input_tokens * p_in + output_tokens * p_out) / 1_000_000
