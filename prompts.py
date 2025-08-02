"""
Translation prompt templates and generators
"""
from typing import List
from config import Config


class PromptGenerator:
    """Generates translation prompts with consistent rules"""
    
    @staticmethod
    def _get_base_rules() -> str:
        """Get base translation rules"""
        return """CRITICAL RULES:
- Keep brand names (AWS, Amazon, Microsoft, Google, Apple, etc.) untranslated
- Keep AWS Services (AWS Lambda, Amazon Bedrock, Amazon S3, etc.) untranslated
- Keep company names, person names, and product names untranslated  
- Keep technical terms that are commonly used in English (API, SDK, CLI, etc.)
- Keep time expressions, currency amounts, and numbers untranslated
- Preserve formatting markers (bullets, numbers, etc.)"""
    
    @staticmethod
    def _get_polishing_instruction(enable_polishing: bool) -> str:
        """Get polishing instruction based on setting"""
        if not enable_polishing:
            return ""
        
        return """
- Focus on natural, fluent translation rather than literal word-for-word translation
- Adapt expressions and idioms to sound natural in the target language
- Maintain the original meaning while making it sound like native content
- Use appropriate tone and style for the target language and context"""
    
    @staticmethod
    def _get_korean_terminology_rules() -> str:
        """Get Korean-specific terminology rules"""
        terminology_list = []
        for en, ko in Config.KOREAN_TERMINOLOGY.items():
            terminology_list.append(f'- "{en}" → "{ko}" (consistently use this term)')
        
        return f"""
TERMINOLOGY CONSISTENCY (Korean):
{chr(10).join(terminology_list)}
- Use the SAME translation for the SAME English term throughout the entire presentation

PERSON NAMES - DO NOT TRANSLATE:
- Keep ALL person names in original English
- Do NOT translate names to Korean phonetic equivalents
- Names should remain exactly as: "Daekeun Kim", "Gil-dong Hong", etc."""
    
    @classmethod
    def create_batch_prompt(cls, target_language: str, enable_polishing: bool = True) -> str:
        """Create optimized batch translation prompt"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        base_rules = cls._get_base_rules()
        polishing_instruction = cls._get_polishing_instruction(enable_polishing)
        
        terminology_rules = ""
        if target_language == 'ko':
            terminology_rules = cls._get_korean_terminology_rules()
        
        return f"""You are a professional translator specializing in business and technical presentations. Translate the following texts to {target_lang_name}.

{base_rules}
- Each text is separated by "---SEPARATOR---"
- Return translations in the SAME ORDER, separated by "---SEPARATOR---"
- Return ONLY the translated texts with NO additional explanations, comments, or metadata
- Do NOT include phrases like "Here is the translation:" or "Translated text:"
- Do NOT add quotation marks around the results{polishing_instruction}{terminology_rules}

Input format:
Text 1
---SEPARATOR---
Text 2
---SEPARATOR---
Text 3

Expected output format:
Translated Text 1
---SEPARATOR---
Translated Text 2
---SEPARATOR---
Translated Text 3

Respond with the translated texts only:"""
    
    @classmethod
    def create_single_prompt(cls, target_language: str, enable_polishing: bool = True) -> str:
        """Create prompt for single text translation"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        base_rules = cls._get_base_rules()
        polishing_instruction = cls._get_polishing_instruction(enable_polishing)
        
        terminology_rules = ""
        if target_language == 'ko':
            terminology_rules = cls._get_korean_terminology_rules()
        
        return f"""You are a professional translator specializing in business and technical presentations. Translate the following text to {target_lang_name}.

{base_rules}
- Return ONLY the translated text with NO additional explanations, comments, or metadata
- Do NOT include phrases like "Here is the translation:" or "Translated text:"
- Do NOT add quotation marks around the result{polishing_instruction}{terminology_rules}

Respond with the translated text only:"""
    
    @classmethod
    def create_context_prompt(cls, target_language: str, slide_context: str, enable_polishing: bool = True) -> str:
        """Create context-aware translation prompt"""
        target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        base_rules = cls._get_base_rules()
        polishing_instruction = cls._get_polishing_instruction(enable_polishing)
        
        if enable_polishing:
            polishing_instruction += """
- Ensure consistent terminology and style throughout the slide
- Consider the relationship between different text elements for better flow"""
        
        terminology_rules = ""
        if target_language == 'ko':
            terminology_rules = cls._get_korean_terminology_rules()
        
        return f"""You are a professional presentation translator. Translate all the numbered texts in the slide content to {target_lang_name}, maintaining context and coherence across the entire slide.

SLIDE CONTEXT:
{slide_context}

{base_rules}
- Maintain consistency in terminology across all texts in the slide
- Consider the context and relationship between different text elements{polishing_instruction}{terminology_rules}

RESPONSE FORMAT:
Return ONLY the translated texts in this exact format:
[1] Translated text 1
[2] Translated text 2
[3] Translated text 3
...

Do NOT include any explanations, comments, or additional text. Just the numbered translations."""
