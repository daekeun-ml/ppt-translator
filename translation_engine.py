"""
Core translation engine using AWS Bedrock
"""
import logging
from typing import List, Dict, Any
from config import Config
from bedrock_client import BedrockClient
from prompts import PromptGenerator
from text_utils import TextProcessor, SlideTextCollector

logger = logging.getLogger(__name__)


class TranslationEngine:
    """Core translation engine using AWS Bedrock"""
    
    def __init__(self, model_id: str = Config.DEFAULT_MODEL_ID, enable_polishing: bool = Config.ENABLE_POLISHING):
        self.model_id = model_id
        self.enable_polishing = enable_polishing
        self.bedrock = BedrockClient()
        self.text_processor = TextProcessor()
        self.prompt_generator = PromptGenerator()
        
        logger.info(f"🎨 Translation mode: {'Natural/Polished' if enable_polishing else 'Literal'}")
    
    def translate_text(self, text: str, target_language: str) -> str:
        """Translate single text"""
        if self.text_processor.should_skip_translation(text):
            return text
        
        try:
            prompt = self.prompt_generator.create_single_prompt(target_language, self.enable_polishing)
            
            response = self.bedrock.converse(
                modelId=self.model_id,
                system=[{"text": prompt}],
                messages=[{
                    "role": "user",
                    "content": [{"text": text}]
                }],
                inferenceConfig={
                    "maxTokens": Config.MAX_TOKENS,
                    "temperature": Config.TEMPERATURE
                }
            )
            
            translated_text = response['output']['message']['content'][0]['text'].strip()
            translated_text = self.text_processor.clean_translation_response(translated_text)
            
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
            # Create batch input
            batch_input = "---SEPARATOR---".join(translatable_texts)
            prompt = self.prompt_generator.create_batch_prompt(target_language, self.enable_polishing)
            
            logger.info(f"🔄 Batch translating {len(translatable_texts)} texts...")
            
            response = self.bedrock.converse(
                modelId=self.model_id,
                system=[{"text": prompt}],
                messages=[{
                    "role": "user",
                    "content": [{"text": batch_input}]
                }],
                inferenceConfig={
                    "maxTokens": Config.MAX_TOKENS,
                    "temperature": Config.TEMPERATURE
                }
            )
            
            translated_batch = response['output']['message']['content'][0]['text'].strip()
            cleaned_parts = self.text_processor.parse_batch_response(translated_batch, len(translatable_texts))
            
            # Handle count mismatch
            if len(cleaned_parts) != len(translatable_texts):
                logger.warning("⚠️ Batch translation count mismatch, using fallback")
                return self._fallback_individual_translation(texts, target_language)
            
            # Reconstruct results with skipped texts
            results = texts.copy()
            translatable_idx = 0
            
            for i, text in enumerate(texts):
                if i not in skip_indices:
                    if translatable_idx < len(cleaned_parts):
                        results[i] = cleaned_parts[translatable_idx]
                        translatable_idx += 1
            
            logger.info(f"✅ Batch translation completed for {len(translatable_texts)} texts")
            return results
            
        except Exception as e:
            logger.error(f"❌ Batch translation error: {str(e)}")
            return self._fallback_individual_translation(texts, target_language)
    
    def translate_with_context(self, text_items: List[Dict], target_language: str, notes_text: str = "") -> List[str]:
        """Translate with full context awareness"""
        if not text_items:
            return []
        
        try:
            slide_context = SlideTextCollector.build_slide_context(text_items, notes_text)
            prompt = self.prompt_generator.create_context_prompt(target_language, slide_context, self.enable_polishing)
            
            logger.info(f"🎯 Context-aware translation for {len(text_items)} texts...")
            
            response = self.bedrock.converse(
                modelId=self.model_id,
                system=[{"text": prompt}],
                messages=[{
                    "role": "user",
                    "content": [{"text": "Please translate all the texts maintaining context and coherence."}]
                }],
                inferenceConfig={
                    "maxTokens": Config.MAX_TOKENS,
                    "temperature": Config.TEMPERATURE
                }
            )
            
            translated_response = response['output']['message']['content'][0]['text'].strip()
            translations = self.text_processor.parse_context_response(translated_response)
            
            if len(translations) == len(text_items):
                logger.info(f"✅ Context-aware translation completed: {len(translations)} texts")
                return translations
            else:
                logger.warning("⚠️ Context translation parsing failed, falling back to batch")
                texts = [item['text'] for item in text_items]
                return self.translate_batch(texts, target_language)
            
        except Exception as e:
            logger.error(f"❌ Context translation error: {str(e)}")
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
