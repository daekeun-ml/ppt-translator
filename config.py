"""
Configuration constants and settings for PowerPoint Translator
"""
import os
from pathlib import Path
from typing import Dict, List
from dotenv import load_dotenv

# Load environment variables from .env file
env_path = Path(__file__).parent / '.env'
load_dotenv(dotenv_path=env_path)


class Config:
    """Configuration constants"""
    # AWS Configuration
    AWS_REGION = os.getenv('AWS_REGION', 'us-east-1')
    AWS_PROFILE = os.getenv('AWS_PROFILE', 'default')
    
    # Translation settings from environment
    DEFAULT_TARGET_LANGUAGE = os.getenv('DEFAULT_TARGET_LANGUAGE', 'ko')
    DEFAULT_MODEL_ID = os.getenv('BEDROCK_MODEL_ID', 'us.anthropic.claude-3-7-sonnet-20250219-v1:0')
    MAX_TOKENS = int(os.getenv('MAX_TOKENS', '4000'))
    TEMPERATURE = float(os.getenv('TEMPERATURE', '0.1'))
    ENABLE_POLISHING = os.getenv('ENABLE_POLISHING', 'true').lower() == 'true'
    BATCH_SIZE = int(os.getenv('BATCH_SIZE', '20'))
    CONTEXT_THRESHOLD = int(os.getenv('CONTEXT_THRESHOLD', '5'))
    
    # Debug settings
    DEBUG = os.getenv('DEBUG', 'false').lower() == 'true'
    
    # Supported models
    SUPPORTED_MODELS = [
        "amazon.nova-micro-v1:0",
        "amazon.nova-lite-v1:0", 
        "amazon.nova-pro-v1:0",
        "anthropic.claude-3-5-sonnet-20241022-v2:0",
        "anthropic.claude-3-5-haiku-20241022-v1:0",
        "anthropic.claude-3-opus-20240229-v1:0",
        "anthropic.claude-3-sonnet-20240229-v1:0",
        "anthropic.claude-3-haiku-20240307-v1:0",
        "us.anthropic.claude-3-7-sonnet-20250219-v1:0",
        "us.anthropic.claude-sonnet-4-20250514-v1:0",
        "us.anthropic.claude-opus-4-20250514-v1:0"
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
    def check_aws_credentials(cls):
        """Check if AWS credentials are properly configured"""
        import boto3
        from botocore.exceptions import NoCredentialsError, PartialCredentialsError
        
        try:
            # Try to create a session with the specified profile
            if cls.AWS_PROFILE and cls.AWS_PROFILE != 'default':
                session = boto3.Session(profile_name=cls.AWS_PROFILE)
            else:
                session = boto3.Session()
            
            # Try to get credentials
            credentials = session.get_credentials()
            if credentials is None:
                return False, "No AWS credentials found. Please run 'aws configure' to set up your credentials."
            
            # Try to make a simple AWS call to verify credentials work
            sts = session.client('sts', region_name=cls.AWS_REGION)
            sts.get_caller_identity()
            
            return True, "AWS credentials are properly configured."
            
        except NoCredentialsError:
            return False, "No AWS credentials found. Please run 'aws configure' to set up your credentials."
        except PartialCredentialsError:
            return False, "Incomplete AWS credentials. Please run 'aws configure' to complete your credential setup."
        except Exception as e:
            return False, f"AWS credential verification failed: {str(e)}"
