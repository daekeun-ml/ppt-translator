"""
Core translation engine using AWS Bedrock
"""
import logging
from typing import List, Dict, Any
from .config import Config
from .bedrock_client import BedrockClient
from .prompts import PromptGenerator
from .text_utils import TextProcessor, SlideTextCollector

logger = logging.getLogger(__name__)


class TranslationEngine:
    """Core translation engine using AWS Bedrock"""
    
    def __init__(self, model_id: str = Config.DEFAULT_MODEL_ID, enable_polishing: bool = Config.ENABLE_POLISHING):
        self.model_id = model_id
        self.enable_polishing = enable_polishing
        self.bedrock = BedrockClient()
        self.text_processor = TextProcessor()
        self.prompt_generator = PromptGenerator()
                
        # Log configuration settings
        self._log_configuration()
        logger.info(f"🎨 Translation mode: {'Natural/Polished' if enable_polishing else 'Literal'}")
        
    def _log_configuration(self):
        """Log current configuration settings"""
        logger.info("⚙️ Configuration Settings:")
        logger.info(f"  AWS Region: {Config.AWS_REGION}")
        logger.info(f"  AWS Profile: {Config.AWS_PROFILE}")
        logger.info(f"  Default Language: {Config.DEFAULT_TARGET_LANGUAGE}")
        logger.info(f"  Model ID: {Config.DEFAULT_MODEL_ID}")
        logger.info(f"  Max Tokens: {Config.MAX_TOKENS}")
        logger.info(f"  Temperature: {Config.TEMPERATURE}")
        logger.info(f"  Enable Polishing: {Config.ENABLE_POLISHING}")
        logger.info(f"  Batch Size: {Config.BATCH_SIZE}")
        logger.info(f"  Context Threshold: {Config.CONTEXT_THRESHOLD}")
        logger.info(f"  Debug Mode: {Config.DEBUG}")
        logger.info(f"  Text AutoFit: {Config.ENABLE_TEXT_AUTOFIT}")
        logger.info(f"  Korean Font: {Config.FONT_KOREAN}")
        logger.info(f"  Japanese Font: {Config.FONT_JAPANESE}")
        logger.info(f"  English Font: {Config.FONT_ENGLISH}")
        logger.info(f"  Chinese Font: {Config.FONT_CHINESE}")
        logger.info(f"  Default Font: {Config.FONT_DEFAULT}")
            
    
    def translate_text(self, text: str, target_language: str) -> str:
        """Translate single text"""
        if self.text_processor.should_skip_translation(text):
            return text
        
        try:
            prompt = self.prompt_generator.create_single_prompt(target_language, self.enable_polishing)
            
            target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
            response = self.bedrock.converse(
                modelId=self.model_id,
                system=[{"text": "You are a translator. Provide ONLY the translation. No explanations, alternatives, context notes, arrows, or additional text."}],
                messages=[{
                    "role": "user",
                    "content": [{"text": f"{prompt}\n\nText: {text}"}]
                }],
                inferenceConfig={
                    "maxTokens": Config.MAX_TOKENS,
                    "temperature": Config.TEMPERATURE
                }
            )
            
            translated_text = response['output']['message']['content'][0]['text'].strip()
            translated_text = self.text_processor.clean_translation_response(translated_text)
            
            # If cleaning resulted in empty text, return original
            if not translated_text:
                logger.warning(f"Empty translation response, keeping original: {text[:50]}...")
                return text
            
            # Remove quotes if wrapped
            if (translated_text.startswith('"') and translated_text.endswith('"')) or \
               (translated_text.startswith("'") and translated_text.endswith("'")):
                translated_text = translated_text[1:-1].strip()
            
            logger.debug(f"Translated: '{text[:50]}...' -> '{translated_text[:50]}...'")
            return translated_text
            
        except Exception as e:
            logger.error(f"Translation error: {str(e)}")
            return text
    
    def translate_batch(self, texts: List[str], target_language: str) -> List[str]:
        """Translate multiple texts in a single API call"""
        if not texts:
            return []
        
        logger.info(f"🔄 Starting batch translation of {len(texts)} texts to {target_language}")
        
        # Filter translatable texts
        translatable_texts = []
        skip_indices = []
        
        for i, text in enumerate(texts):
            if self.text_processor.should_skip_translation(text):
                skip_indices.append(i)
                logger.debug(f"⏭️ Skipping text {i}: {text[:30]}...")
            else:
                translatable_texts.append(text)
                logger.debug(f"✅ Will translate text {i}: {text[:30]}...")
        
        if not translatable_texts:
            return texts
        
        try:
            # Create batch input with numbered format for better parsing
            batch_input = ""
            for i, text in enumerate(translatable_texts, 1):
                batch_input += f"[{i}] {text}\n"
            
            prompt = self.prompt_generator.create_batch_prompt(target_language, self.enable_polishing)
            
            logger.info(f"🔄 Batch translating {len(translatable_texts)} texts...")
            
            target_lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
            response = self.bedrock.converse(
                modelId=self.model_id,
                system=[{"text": "You are a translator. Translate each numbered text exactly as provided. Respond ONLY with translations in the same numbered format. Do not add explanations, alternatives, or additional content."}],
                messages=[{
                    "role": "user",
                    "content": [{"text": f"{prompt}\n\n{batch_input}"}]
                }],
                inferenceConfig={
                    "maxTokens": Config.MAX_TOKENS,
                    "temperature": Config.TEMPERATURE
                }
            )
            
            translated_batch = response['output']['message']['content'][0]['text'].strip()
            
            # Try numbered parsing first
            cleaned_parts = self.text_processor.parse_numbered_response(translated_batch, len(translatable_texts))
            
            # If numbered parsing fails, try separator parsing
            if len(cleaned_parts) != len(translatable_texts):
                cleaned_parts = self.text_processor.parse_batch_response(translated_batch, len(translatable_texts))
            
            # Only allow exact match or fallback immediately
            if len(cleaned_parts) != len(translatable_texts):
                logger.warning(f"⚠️ Batch translation count mismatch. Expected {len(translatable_texts)}, got {len(cleaned_parts)}, using fallback")
                return self._fallback_individual_translation(texts, target_language)
            
            # Reconstruct results with skipped texts
            results = texts.copy()
            translatable_idx = 0
            
            for i, text in enumerate(texts):
                if i not in skip_indices:
                    if translatable_idx < len(cleaned_parts):
                        results[i] = cleaned_parts[translatable_idx]
                        translatable_idx += 1
                    # If we run out of translations, keep original text
            
            logger.info(f"✅ Batch translation completed for {min(len(cleaned_parts), len(translatable_texts))} texts")
            return results
            
        except Exception as e:
            logger.error(f"❌ Batch translation error: {str(e)}")
            return self._fallback_individual_translation(texts, target_language)
    
    def translate_with_context(self, text_items: List[Dict], target_language: str, notes_text: str = "") -> List[str]:
        """Translate with full context awareness - simplified to use batch translation"""
        if not text_items:
            return []
        
        logger.info(f"🔄 Context translation requested for {len(text_items)} texts, using batch translation instead")
        
        # Extract texts and use batch translation (more reliable)
        texts = [item['text'] for item in text_items]
        return self.translate_batch(texts, target_language)
    
    def _fallback_individual_translation(self, texts: List[str], target_language: str) -> List[str]:
        """Fallback to individual translation when batch fails"""
        logger.info(f"🔄 Falling back to individual translation for {len(texts)} texts...")
        results = []
        
        for i, text in enumerate(texts):
            try:
                translated = self.translate_text(text, target_language)
                results.append(translated)
                logger.debug(f"✅ Individual translation {i+1}/{len(texts)}")
            except Exception as e:
                logger.error(f"❌ Failed to translate text {i+1}: {str(e)}")
                results.append(text)
        
        logger.info(f"✅ Individual translation fallback completed: {len(results)} results")
        return results
