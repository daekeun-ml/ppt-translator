"""
Configuration constants and settings for PowerPoint Translator
"""
import os
from pathlib import Path
from typing import Dict, List
from dotenv import load_dotenv

# Load environment variables from .env file
env_path = Path(__file__).parent.parent / '.env'
load_dotenv(dotenv_path=env_path)


class Config:
    """Configuration constants"""
    # AWS Configuration
    AWS_REGION = os.getenv('AWS_REGION', 'us-east-1')
    AWS_PROFILE = os.getenv('AWS_PROFILE')
    MANTLE_API_KEY = os.getenv('AWS_BEARER_TOKEN_BEDROCK')
    MANTLE_OPENAI_BASE_URL = os.getenv(
        'BEDROCK_MANTLE_OPENAI_BASE_URL',
        f'https://bedrock-mantle.{AWS_REGION}.api.aws/openai/v1',
    ).rstrip('/')
    MANTLE_TIMEOUT_SECONDS = float(os.getenv('BEDROCK_MANTLE_TIMEOUT_SECONDS', '300'))
    
    # Translation settings from environment
    DEFAULT_TARGET_LANGUAGE = os.getenv('DEFAULT_TARGET_LANGUAGE', 'ko')
    DEFAULT_MODEL_ID = os.getenv(
        'MANTLE_MODEL_ID',
        os.getenv('BEDROCK_MODEL_ID', 'openai.gpt-5.6-terra'),
    )
    ENABLE_MODEL_FALLBACK = (
        os.getenv('MANTLE_ENABLE_MODEL_FALLBACK', 'true').lower() == 'true'
    )
    FALLBACK_MODEL_ID = os.getenv(
        'MANTLE_FALLBACK_MODEL_ID',
        'openai.gpt-5.6-luna',
    ).strip()
    MAX_TOKENS = int(os.getenv('MAX_TOKENS', '4000'))
    TEMPERATURE = float(os.getenv('TEMPERATURE', '0.1'))
    OPENAI_REASONING_EFFORT = os.getenv('OPENAI_REASONING_EFFORT', 'none')
    ENABLE_POLISHING = os.getenv('ENABLE_POLISHING', 'true').lower() == 'true'
    BATCH_SIZE = int(os.getenv('BATCH_SIZE', '20'))
    CONTEXT_THRESHOLD = int(os.getenv('CONTEXT_THRESHOLD', '100'))  # Effectively disable context translation
    BATCH_WORKERS = max(1, int(os.getenv('BATCH_WORKERS', '10')))
    SLIDE_WORKERS = max(1, int(os.getenv('SLIDE_WORKERS', '4')))
    SLIDES_PER_WORKER = max(1, int(os.getenv('SLIDES_PER_WORKER', '30')))

    # PowerPoint-to-Markdown settings
    MARKDOWN_WORKERS = max(1, int(os.getenv('MARKDOWN_WORKERS', '4')))
    MARKDOWN_SLIDES_PER_CHUNK = max(
        1,
        int(os.getenv('MARKDOWN_SLIDES_PER_CHUNK', '10')),
    )
    MARKDOWN_MAX_TOKENS = max(
        256,
        int(os.getenv('MARKDOWN_MAX_TOKENS', '3000')),
    )
    MARKDOWN_OVERVIEW_MAX_TOKENS = max(
        256,
        int(os.getenv('MARKDOWN_OVERVIEW_MAX_TOKENS', '2000')),
    )
    MARKDOWN_WEB_MAX_TOKENS = max(
        256,
        int(os.getenv('MARKDOWN_WEB_MAX_TOKENS', '1500')),
    )
    MARKDOWN_MAX_WEB_QUERIES = max(
        1,
        int(os.getenv('MARKDOWN_MAX_WEB_QUERIES', '3')),
    )
    MARKDOWN_SEARCH_RESULTS_PER_QUERY = max(
        1,
        int(os.getenv('MARKDOWN_SEARCH_RESULTS_PER_QUERY', '5')),
    )
    MARKDOWN_REASONING_EFFORT = os.getenv(
        'MARKDOWN_REASONING_EFFORT',
        'low',
    ).strip()
    
    # Debug settings
    DEBUG = os.getenv('DEBUG', 'false').lower() == 'true'
    
    # Post-processing settings
    ENABLE_TEXT_AUTOFIT = os.getenv('ENABLE_TEXT_AUTOFIT', 'true').lower() == 'true'
    TEXT_LENGTH_THRESHOLD = int(os.getenv('TEXT_LENGTH_THRESHOLD', '10'))
    
    # Font settings by language
    FONT_KOREAN = os.getenv('FONT_KOREAN', '맑은 고딕')
    FONT_JAPANESE = os.getenv('FONT_JAPANESE', 'Yu Gothic UI')
    FONT_ENGLISH = os.getenv('FONT_ENGLISH', 'Amazon Ember')
    FONT_CHINESE = os.getenv('FONT_CHINESE', 'Microsoft YaHei')
    FONT_DEFAULT = os.getenv('FONT_DEFAULT', 'Arial')
    
    # Font mapping by language code
    FONT_MAP = {
        'ko': FONT_KOREAN,
        'ja': FONT_JAPANESE,
        'en': FONT_ENGLISH,
        'en-US': FONT_ENGLISH,
        'en-GB': FONT_ENGLISH,
        'en-AU': FONT_ENGLISH,
        'en-CA': FONT_ENGLISH,
        'zh': FONT_CHINESE,
        'zh-CN': FONT_CHINESE,
        'zh-TW': FONT_CHINESE,
        'zh-HK': FONT_CHINESE,
        'zh-SG': FONT_CHINESE,
        'zh-MY': FONT_CHINESE,
    }
    
    # Models available through the Amazon Bedrock Mantle endpoint.
    SUPPORTED_MODELS = [
        # OpenAI GPT-5.6 family
        "openai.gpt-5.6-sol",
        "openai.gpt-5.6-terra",
        "openai.gpt-5.6-luna",
        "openai.gpt-5.6-cyber",

        # Anthropic Claude
        "anthropic.claude-opus-5",
        "anthropic.claude-sonnet-5",
        "anthropic.claude-haiku-4-5",
    ]
    
    # Language mapping - Comprehensive list of supported languages
    LANGUAGE_MAP = {
        # Major languages
        'en': 'English',
        'ko': 'Korean',
        'ja': 'Japanese',
        'zh': 'Chinese (Simplified)',
        'zh-CN': 'Chinese (Simplified)',
        'zh-TW': 'Chinese (Traditional)',
        'zh-HK': 'Chinese (Hong Kong)',
        
        # European languages
        'fr': 'French',
        'de': 'German',
        'es': 'Spanish',
        'it': 'Italian',
        'pt': 'Portuguese',
        'pt-BR': 'Portuguese (Brazil)',
        'ru': 'Russian',
        'nl': 'Dutch',
        'sv': 'Swedish',
        'no': 'Norwegian',
        'da': 'Danish',
        'fi': 'Finnish',
        'pl': 'Polish',
        'cs': 'Czech',
        'sk': 'Slovak',
        'hu': 'Hungarian',
        'ro': 'Romanian',
        'bg': 'Bulgarian',
        'hr': 'Croatian',
        'sr': 'Serbian',
        'sl': 'Slovenian',
        'et': 'Estonian',
        'lv': 'Latvian',
        'lt': 'Lithuanian',
        'el': 'Greek',
        'tr': 'Turkish',
        'uk': 'Ukrainian',
        'be': 'Belarusian',
        'mk': 'Macedonian',
        'mt': 'Maltese',
        'is': 'Icelandic',
        'ga': 'Irish',
        'cy': 'Welsh',
        'eu': 'Basque',
        'ca': 'Catalan',
        'gl': 'Galician',
        
        # Middle Eastern and African languages
        'ar': 'Arabic',
        'he': 'Hebrew',
        'fa': 'Persian (Farsi)',
        'ur': 'Urdu',
        'sw': 'Swahili',
        'am': 'Amharic',
        'ha': 'Hausa',
        'yo': 'Yoruba',
        'ig': 'Igbo',
        'zu': 'Zulu',
        'af': 'Afrikaans',
        
        # South Asian languages
        'hi': 'Hindi',
        'bn': 'Bengali',
        'te': 'Telugu',
        'mr': 'Marathi',
        'ta': 'Tamil',
        'gu': 'Gujarati',
        'kn': 'Kannada',
        'ml': 'Malayalam',
        'pa': 'Punjabi',
        'or': 'Odia',
        'as': 'Assamese',
        'ne': 'Nepali',
        'si': 'Sinhala',
        'my': 'Burmese',
        
        # Southeast Asian languages
        'th': 'Thai',
        'vi': 'Vietnamese',
        'id': 'Indonesian',
        'ms': 'Malay',
        'tl': 'Filipino (Tagalog)',
        'km': 'Khmer',
        'lo': 'Lao',
        
        # Other languages
        'az': 'Azerbaijani',
        'kk': 'Kazakh',
        'ky': 'Kyrgyz',
        'uz': 'Uzbek',
        'tg': 'Tajik',
        'mn': 'Mongolian',
        'ka': 'Georgian',
        'hy': 'Armenian',
        'sq': 'Albanian',
        'mk': 'Macedonian',
        'lv': 'Latvian',
        'lt': 'Lithuanian',
        'et': 'Estonian',
        
        # Additional variants and regional codes
        'en-US': 'English (US)',
        'en-GB': 'English (UK)',
        'en-AU': 'English (Australia)',
        'en-CA': 'English (Canada)',
        'fr-CA': 'French (Canada)',
        'fr-CH': 'French (Switzerland)',
        'de-AT': 'German (Austria)',
        'de-CH': 'German (Switzerland)',
        'es-MX': 'Spanish (Mexico)',
        'es-AR': 'Spanish (Argentina)',
        'es-CO': 'Spanish (Colombia)',
        'es-CL': 'Spanish (Chile)',
        'es-PE': 'Spanish (Peru)',
        'es-VE': 'Spanish (Venezuela)',
        'pt-PT': 'Portuguese (Portugal)',
        'it-CH': 'Italian (Switzerland)',
        'nl-BE': 'Dutch (Belgium)',
        'sv-FI': 'Swedish (Finland)',
        'ar-SA': 'Arabic (Saudi Arabia)',
        'ar-EG': 'Arabic (Egypt)',
        'ar-AE': 'Arabic (UAE)',
        'ar-MA': 'Arabic (Morocco)',
        'zh-SG': 'Chinese (Singapore)',
        'zh-MY': 'Chinese (Malaysia)',
        'ms-SG': 'Malay (Singapore)',
        'ta-SG': 'Tamil (Singapore)',
        'hi-IN': 'Hindi (India)',
        'bn-BD': 'Bengali (Bangladesh)',
        'ur-PK': 'Urdu (Pakistan)',
        'fa-IR': 'Persian (Iran)',
        'fa-AF': 'Persian (Afghanistan)',
        'ps': 'Pashto',
        'sd': 'Sindhi',
        'ckb': 'Kurdish (Sorani)',
        'ku': 'Kurdish (Kurmanji)',
        'yi': 'Yiddish',
        'la': 'Latin',
        'eo': 'Esperanto',
        'jv': 'Javanese',
        'su': 'Sundanese',
        'ceb': 'Cebuano',
        'haw': 'Hawaiian',
        'mi': 'Maori',
        'sm': 'Samoan',
        'to': 'Tongan',
        'fj': 'Fijian',
        'mg': 'Malagasy',
        'ny': 'Chichewa',
        'sn': 'Shona',
        'st': 'Sesotho',
        'tn': 'Setswana',
        'ts': 'Tsonga',
        've': 'Venda',
        'xh': 'Xhosa',
        'co': 'Corsican',
        'fy': 'Frisian',
        'gd': 'Scottish Gaelic',
        'lb': 'Luxembourgish',
        'rm': 'Romansh'
    }
    
    # Korean-specific terminology rules
    KOREAN_TERMINOLOGY = {
        "Observability": "Observability",
        "AgentCore Observability": "AgentCore Observability",
        "Key concepts": "핵심 개념",
        "Best Practices": "모범 사례",
        "Resources": "리소스",
        "Demos": "데모",
        "Pricing": "가격 책정"
    }
    
    # Text patterns to skip translation
    SKIP_PATTERNS = [
        r'^\d+$',  # Numbers only
        r'^https?://',  # URLs
        r'\S+@\S+\.\S+',  # Email addresses
        r'^```.*```$',  # Code blocks
        r'^\s*[{}\[\]();,.:]+\s*$',  # Code syntax characters only
        r'^\s*import\s+\w+',  # Python imports
        r'^\s*from\s+\w+\s+import',  # Python from imports
        r'^\s*def\s+\w+\(',  # Python function definitions
        r'^\s*class\s+\w+',  # Python class definitions
        r'^\s*if\s+.*:',  # Python if statements
        r'^\s*for\s+.*:',  # Python for loops
        r'^\s*while\s+.*:',  # Python while loops
        r'^\s*try\s*:',  # Python try blocks
        r'^\s*except\s*.*:',  # Python except blocks
        r'^\s*return\s+',  # Python return statements
        r'^\s*print\s*\(',  # Python print statements
        r'^\s*console\.log\s*\(',  # JavaScript console.log
        r'^\s*function\s+\w+\s*\(',  # JavaScript functions
        r'^\s*var\s+\w+\s*=',  # JavaScript var declarations
        r'^\s*let\s+\w+\s*=',  # JavaScript let declarations
        r'^\s*const\s+\w+\s*=',  # JavaScript const declarations
        r'^\s*\$\s*\(',  # jQuery
        r'^\s*<\w+.*>.*</\w+>\s*$',  # HTML tags
        r'^\s*<\w+.*/?>\s*$',  # Self-closing HTML tags
    ]
    
    @classmethod
    def get_language_name(cls, language_code: str) -> str:
        """Get the full language name from language code"""
        return cls.LANGUAGE_MAP.get(language_code, language_code)
    
    @classmethod
    def validate_model_id(cls, model_id: str) -> bool:
        """Validate if the model ID is supported"""
        return model_id in cls.SUPPORTED_MODELS
    
    @classmethod
    def reload_env(cls):
        """Reload environment variables (useful for testing)"""
        load_dotenv(dotenv_path=env_path, override=True)
    
    @classmethod
    def get_font_for_language(cls, language_code: str) -> str:
        """Get the appropriate font for a given language code"""
        return cls.FONT_MAP.get(language_code, cls.FONT_DEFAULT)
    
    @classmethod
    def check_aws_credentials(cls):
        """Check whether a Mantle API key or AWS credentials are usable."""
        if cls.MANTLE_API_KEY:
            return True, "Amazon Bedrock Mantle API key is configured."
        try:
            from aws_bedrock_token_generator import provide_token
            provide_token(region=cls.AWS_REGION)
            return True, "AWS default credentials can generate a Bedrock token."
        except Exception:
            return (
                False,
                "No usable AWS credentials found. Run 'aws configure', use an "
                "IAM role, or set AWS_BEARER_TOKEN_BEDROCK.",
            )
    
    def __init__(self):
        """Initialize configuration with environment variables"""
        self._env_vars = {}
        self._load_env_vars()
    
    def _load_env_vars(self):
        """Load all environment variables"""
        for key, value in os.environ.items():
            self._env_vars[key] = value
    
    def get(self, key: str, default: str = None) -> str:
        """Get configuration value by key"""
        return self._env_vars.get(key, default)
    
    def get_bool(self, key: str, default: bool = False) -> bool:
        """Get boolean configuration value"""
        value = self.get(key, str(default).lower())
        return value.lower() in ('true', '1', 'yes', 'on')
    
    def get_int(self, key: str, default: int = 0) -> int:
        """Get integer configuration value"""
        try:
            return int(self.get(key, str(default)))
        except (ValueError, TypeError):
            return default
    
    def get_float(self, key: str, default: float = 0.0) -> float:
        """Get float configuration value"""
        try:
            return float(self.get(key, str(default)))
        except (ValueError, TypeError):
            return default
    
    def set(self, key: str, value: str):
        """Set configuration value"""
        self._env_vars[key] = value
        os.environ[key] = value
