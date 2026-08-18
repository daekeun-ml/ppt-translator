"""Core translation engine using Amazon Bedrock Mantle."""
import logging
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from .config import Config
from .bedrock_client import BedrockAuthenticationError, BedrockClient
from .prompts import PromptGenerator
from .text_utils import TextProcessor, SlideTextCollector
from .cache import TranslationCache, NullCache, make_cache_key
from .glossary import hash_glossary

logger = logging.getLogger(__name__)


@dataclass
class TranslationMetrics:
    """Running counters updated by the engine, read by CLI/UI layers."""
    cache_hits: int = 0
    cache_misses: int = 0
    tokens_in: int = 0
    tokens_out: int = 0
    api_calls: int = 0


class TranslationEngine:
    """Core translation engine using Amazon Bedrock Mantle."""

    _config_logged = False  # class-level flag so parallel workers don't spam logs

    def __init__(self, model_id: str = Config.DEFAULT_MODEL_ID,
                 enable_polishing: bool = Config.ENABLE_POLISHING,
                 cache: Optional[TranslationCache] = None,
                 glossary: Optional[Dict[str, str]] = None,
                 source_language: Optional[str] = None,
                 log_initialization: bool = True):
        if not Config.validate_model_id(model_id):
            raise ValueError(
                f"Unsupported Bedrock Mantle model '{model_id}'. "
                f"Supported models: {', '.join(Config.SUPPORTED_MODELS)}"
            )
        self.model_id = model_id
        self.enable_polishing = enable_polishing
        self.bedrock = BedrockClient()
        self.text_processor = TextProcessor()
        self.prompt_generator = PromptGenerator()
        self.cache: TranslationCache = cache if cache is not None else NullCache()
        self.glossary: Dict[str, str] = glossary or {}
        self._glossary_hash = hash_glossary(self.glossary)
        # None = unknown (let the model guess from context). Explicit value is
        # preferred: it sharpens the prompt, disambiguates cache hits, and
        # enables the "skip if source == target" short-circuit.
        self.source_language: Optional[str] = source_language
        self.metrics = TranslationMetrics()

        # Log configuration once per process, not per engine instance.
        if log_initialization and not TranslationEngine._config_logged:
            self._log_configuration()
            TranslationEngine._config_logged = True
        if log_initialization:
            logger.info(f"🎨 Translation mode: {'Natural/Polished' if enable_polishing else 'Literal'}")
        
    def _log_configuration(self):
        """Log current configuration settings"""
        logger.info("⚙️ Configuration Settings:")
        logger.info(f"  AWS Region: {Config.AWS_REGION}")
        logger.info("  Endpoint: Amazon Bedrock Mantle")
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
            
    
    def _cache_key(self, text: str, target_language: str) -> str:
        return make_cache_key(
            text, target_language, self.model_id,
            self.enable_polishing, self._glossary_hash,
            source_language=self.source_language,
        )

    def _should_skip_translation_entirely(self, target_language: str) -> bool:
        """Skip the whole translation if source language is known and matches target."""
        if not self.source_language:
            return False
        src = self.source_language.lower().split('-')[0]
        tgt = (target_language or '').lower().split('-')[0]
        return src == tgt

    def _record_usage(self, response: Dict[str, Any]) -> None:
        """Accumulate token usage from a Bedrock Converse response (best-effort)."""
        try:
            usage = response.get('usage') or {}
            self.metrics.tokens_in += int(usage.get('inputTokens', 0))
            self.metrics.tokens_out += int(usage.get('outputTokens', 0))
        except (AttributeError, TypeError, ValueError):
            pass

    def translate_text(self, text: str, target_language: str) -> str:
        """Translate a single text, hitting the cache first."""
        if self.text_processor.should_skip_translation(text):
            return text
        if self._should_skip_translation_entirely(target_language):
            return text

        cache_key = self._cache_key(text, target_language)
        cached = self.cache.get(cache_key)
        if cached is not None:
            self.metrics.cache_hits += 1
            return cached

        self.metrics.cache_misses += 1

        try:
            prompt = self.prompt_generator.create_single_prompt(
                target_language, self.enable_polishing, glossary=self.glossary,
                source_language=self.source_language,
            )

            self.metrics.api_calls += 1
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
            self._record_usage(response)

            translated_text = response['output']['message']['content'][0]['text'].strip()
            translated_text = self.text_processor.clean_translation_response(translated_text)

            if not translated_text:
                logger.warning(f"Empty translation response, keeping original: {text[:50]}...")
                return text

            if (translated_text.startswith('"') and translated_text.endswith('"')) or \
               (translated_text.startswith("'") and translated_text.endswith("'")):
                translated_text = translated_text[1:-1].strip()

            self.cache.set(cache_key, translated_text)
            logger.debug(f"Translated: '{text[:50]}...' -> '{translated_text[:50]}...'")
            return translated_text

        except BedrockAuthenticationError as e:
            logger.error(f"Translation authentication error: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Translation error: {str(e)}")
            return text

    def translate_batch(self, texts: List[str], target_language: str) -> List[str]:
        """Translate multiple texts, calling the API only for cache misses."""
        if not texts:
            return []

        if self._should_skip_translation_entirely(target_language):
            logger.info(f"🟰 Source language '{self.source_language}' matches target; skipping batch")
            return list(texts)

        logger.info(f"🔄 Starting batch translation of {len(texts)} texts to {target_language}")

        # Classify each input: skip, cache-hit, or needs API.
        skip_indices: set = set()
        cached_at: Dict[int, str] = {}
        uncached_indices: List[int] = []  # indices into `texts`
        uncached_texts: List[str] = []

        for i, text in enumerate(texts):
            if self.text_processor.should_skip_translation(text):
                skip_indices.add(i)
                continue
            key = self._cache_key(text, target_language)
            hit = self.cache.get(key)
            if hit is not None:
                cached_at[i] = hit
                self.metrics.cache_hits += 1
            else:
                uncached_indices.append(i)
                uncached_texts.append(text)

        if not uncached_texts:
            # Everything skipped or served from cache — no API call needed.
            results = texts.copy()
            for i, v in cached_at.items():
                results[i] = v
            logger.info(f"✅ Batch fully served from cache ({len(cached_at)} hits)")
            return results

        self.metrics.cache_misses += len(uncached_texts)

        try:
            batch_input = ""
            for i, text in enumerate(uncached_texts, 1):
                batch_input += f"[{i}] {text}\n"

            prompt = self.prompt_generator.create_batch_prompt(
                target_language, self.enable_polishing, glossary=self.glossary,
                source_language=self.source_language,
            )

            logger.info(f"🔄 Batch translating {len(uncached_texts)} texts (cache hits: {len(cached_at)})...")

            self.metrics.api_calls += 1
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
            self._record_usage(response)

            translated_batch = response['output']['message']['content'][0]['text'].strip()

            cleaned_parts = self.text_processor.parse_numbered_response(translated_batch, len(uncached_texts))
            if len(cleaned_parts) != len(uncached_texts):
                cleaned_parts = self.text_processor.parse_batch_response(translated_batch, len(uncached_texts))

            if len(cleaned_parts) != len(uncached_texts):
                logger.warning(
                    f"⚠️ Batch translation count mismatch. Expected {len(uncached_texts)}, "
                    f"got {len(cleaned_parts)}, using fallback"
                )
                return self._fallback_individual_translation(texts, target_language)

            # Write API results back into results and cache.
            results = texts.copy()
            for i, v in cached_at.items():
                results[i] = v
            for original_idx, translated in zip(uncached_indices, cleaned_parts):
                results[original_idx] = translated
                self.cache.set(
                    self._cache_key(texts[original_idx], target_language),
                    translated,
                )

            logger.info(f"✅ Batch translation completed: {len(uncached_texts)} translated, {len(cached_at)} from cache")
            return results

        except BedrockAuthenticationError as e:
            logger.error(f"❌ Batch translation authentication error: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"❌ Batch translation error: {str(e)}")
            return self._fallback_individual_translation(texts, target_language)
    
    def _fallback_individual_translation(self, texts: List[str], target_language: str) -> List[str]:
        """Fallback to individual translation when batch fails"""
        logger.info(f"🔄 Falling back to individual translation for {len(texts)} texts...")
        results = []
        
        for i, text in enumerate(texts):
            try:
                translated = self.translate_text(text, target_language)
                results.append(translated)
                logger.debug(f"✅ Individual translation {i+1}/{len(texts)}")
            except BedrockAuthenticationError:
                raise
            except Exception as e:
                logger.error(f"❌ Failed to translate text {i+1}: {str(e)}")
                results.append(text)
        
        logger.info(f"✅ Individual translation fallback completed: {len(results)} results")
        return results
