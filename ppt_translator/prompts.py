"""
Translation prompt templates and generators
"""
from typing import List
from .config import Config


class PromptGenerator:
    """Generates translation prompts with consistent rules"""
    
    @classmethod
    def create_single_prompt(cls, target_language: str, enable_polishing: bool = True) -> str:
        """Create prompt for single text translation"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        return f"""Translate to {target_lang_name}. 
CRITICAL: Provide ONLY the translation. No explanations, alternatives, context notes, or additional text."""
    
    @classmethod
    def create_batch_prompt(cls, target_language: str, enable_polishing: bool = True) -> str:
        """Create optimized batch translation prompt"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        return f"""Translate each numbered text to {target_lang_name}. 
CRITICAL RULES:
- Provide ONLY the translation, no explanations
- No alternative translations or context notes
- No markdown formatting (**bold**, *italic*)
- No arrows (→) or additional text
- Keep the same numbered format: [1] translation [2] translation [3] translation
- Do not skip any numbers

Example:
[1] 첫 번째 번역
[2] 두 번째 번역
[3] 세 번째 번역"""
    
    @classmethod
    def create_context_prompt(cls, target_language: str, slide_context: str, enable_polishing: bool = True) -> str:
        """Create context-aware translation prompt"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        return f"Translate numbered texts to {target_lang_name}. Format: [1] translation [2] translation:"
