"""Amazon Bedrock Mantle client with a Bedrock Converse-compatible surface."""
import logging
import threading
import time
from datetime import timedelta
from typing import Any, Callable, Dict, Iterable, Optional

from .config import Config
from .logging_utils import quiet_dependency_logs
from .retry import bedrock_retry, get_status_code, is_model_fallback_error

logger = logging.getLogger(__name__)

_OPENAI_MIN_OUTPUT_TOKENS = 16


class BedrockAuthenticationError(RuntimeError):
    """Raised when neither a Mantle API key nor AWS credentials are usable."""


class BedrockClient:
    """Route OpenAI and Anthropic models through Amazon Bedrock Mantle."""

    def __init__(
        self,
        region: Optional[str] = None,
        api_key: Optional[str] = None,
        openai_client: Optional[Any] = None,
        anthropic_client: Optional[Any] = None,
        token_provider: Optional[Callable[[], str]] = None,
    ):
        self.region = region or Config.AWS_REGION
        self.api_key = api_key if api_key is not None else Config.MANTLE_API_KEY
        self.openai_base_url = Config.MANTLE_OPENAI_BASE_URL
        self.timeout = Config.MANTLE_TIMEOUT_SECONDS
        self._openai_client = openai_client
        self._anthropic_client = anthropic_client
        self._token_provider = token_provider

    def _get_token_provider(self) -> Callable[[], str]:
        if self._token_provider is None:
            try:
                from aws_bedrock_token_generator import provide_token
            except ImportError as exc:
                raise ImportError(
                    "The aws-bedrock-token-generator package is required to "
                    "use existing AWS credentials with GPT models."
                ) from exc

            token_lock = threading.Lock()
            cached_token: Optional[str] = None
            refresh_at = 0.0

            def _provide_token() -> str:
                nonlocal cached_token, refresh_at
                with token_lock:
                    if cached_token is not None and time.monotonic() < refresh_at:
                        return cached_token
                    try:
                        quiet_dependency_logs()
                        cached_token = provide_token(
                            region=self.region,
                            expiry=timedelta(hours=1),
                        )
                        refresh_at = time.monotonic() + (50 * 60)
                        return cached_token
                    except Exception as exc:
                        raise BedrockAuthenticationError(
                            "No usable Amazon Bedrock credentials found. Configure "
                            "the AWS default credential chain (for example, run "
                            "'aws configure' or use an IAM role), or set "
                            "AWS_BEARER_TOKEN_BEDROCK."
                        ) from exc

            self._token_provider = _provide_token
        return self._token_provider

    def _validate_aws_credentials(self) -> Callable[[], str]:
        provider = self._get_token_provider()
        try:
            provider()
        except BedrockAuthenticationError:
            raise
        except Exception as exc:
            raise BedrockAuthenticationError(
                "Unable to generate a short-term Bedrock token from the AWS "
                "default credential chain."
            ) from exc
        return provider

    def _get_openai_client(self) -> Any:
        if self._openai_client is None:
            quiet_dependency_logs()
            try:
                from openai import OpenAI
            except ImportError as exc:
                raise ImportError(
                    "The openai package is required for GPT models. "
                    "Install project dependencies with 'uv sync'."
                ) from exc
            api_key = self.api_key or self._validate_aws_credentials()
            self._openai_client = OpenAI(
                api_key=api_key,
                base_url=self.openai_base_url,
                timeout=self.timeout,
                max_retries=0,
            )
            quiet_dependency_logs()
            logger.info(
                "Initialized Bedrock Mantle OpenAI client for region %s",
                self.region,
            )
        return self._openai_client

    def _get_anthropic_client(self) -> Any:
        if self._anthropic_client is None:
            quiet_dependency_logs()
            try:
                from anthropic import AnthropicBedrockMantle
            except ImportError as exc:
                raise ImportError(
                    "The anthropic[bedrock] package is required for Claude "
                    "models. Install project dependencies with 'uv sync'."
                ) from exc
            client_kwargs: Dict[str, Any] = {
                "aws_region": self.region,
                "timeout": self.timeout,
                "max_retries": 0,
            }
            if Config.AWS_PROFILE:
                client_kwargs["aws_profile"] = Config.AWS_PROFILE
            if self.api_key:
                client_kwargs["api_key"] = self.api_key
            else:
                self._validate_aws_credentials()

            self._anthropic_client = AnthropicBedrockMantle(
                **client_kwargs,
            )
            quiet_dependency_logs()
            logger.info(
                "Initialized Bedrock Mantle Anthropic client for region %s",
                self.region,
            )
        return self._anthropic_client

    @staticmethod
    def _text_from_blocks(blocks: Any) -> str:
        if isinstance(blocks, str):
            return blocks
        if not isinstance(blocks, Iterable):
            return ""

        parts = []
        for block in blocks:
            if isinstance(block, dict):
                text = block.get("text")
            else:
                text = getattr(block, "text", None)
            if text:
                parts.append(str(text))
        return "\n".join(parts)

    @classmethod
    def _system_text(cls, system: Any) -> str:
        return cls._text_from_blocks(system).strip()

    @classmethod
    def _message_list(cls, messages: Any) -> list[Dict[str, str]]:
        normalized = []
        for message in messages or []:
            role = message.get("role", "user")
            content = cls._text_from_blocks(message.get("content", "")).strip()
            normalized.append({"role": role, "content": content})
        return normalized

    @staticmethod
    def _converse_response(
        text: str,
        input_tokens: int = 0,
        output_tokens: int = 0,
    ) -> Dict[str, Any]:
        return {
            "output": {
                "message": {
                    "role": "assistant",
                    "content": [{"text": text}],
                }
            },
            "usage": {
                "inputTokens": int(input_tokens or 0),
                "outputTokens": int(output_tokens or 0),
            },
        }

    def _converse_openai(
        self,
        model_id: str,
        system: Any,
        messages: Any,
        inference_config: Dict[str, Any],
    ) -> Dict[str, Any]:
        max_output_tokens = max(
            _OPENAI_MIN_OUTPUT_TOKENS,
            inference_config.get("maxTokens", Config.MAX_TOKENS),
        )
        request: Dict[str, Any] = {
            "model": model_id,
            "input": self._message_list(messages),
            "max_output_tokens": max_output_tokens,
            "reasoning": {
                "effort": inference_config.get(
                    "reasoningEffort",
                    Config.OPENAI_REASONING_EFFORT,
                )
            },
        }
        instructions = self._system_text(system)
        if instructions:
            request["instructions"] = instructions
        if request["reasoning"]["effort"] == "none":
            request["temperature"] = inference_config.get(
                "temperature",
                Config.TEMPERATURE,
            )

        response = self._get_openai_client().responses.create(**request)
        text = getattr(response, "output_text", "") or ""
        usage = getattr(response, "usage", None)
        return self._converse_response(
            text.strip(),
            getattr(usage, "input_tokens", 0),
            getattr(usage, "output_tokens", 0),
        )

    def _converse_anthropic(
        self,
        model_id: str,
        system: Any,
        messages: Any,
        inference_config: Dict[str, Any],
    ) -> Dict[str, Any]:
        request: Dict[str, Any] = {
            "model": model_id,
            "max_tokens": inference_config.get("maxTokens", Config.MAX_TOKENS),
            "messages": self._message_list(messages),
        }
        system_text = self._system_text(system)
        if system_text:
            request["system"] = system_text

        response = self._get_anthropic_client().messages.create(**request)
        text = self._text_from_blocks(getattr(response, "content", [])).strip()
        usage = getattr(response, "usage", None)
        return self._converse_response(
            text,
            getattr(usage, "input_tokens", 0),
            getattr(usage, "output_tokens", 0),
        )

    def is_ready(self) -> bool:
        """Return whether an API key or AWS default credentials are usable."""
        if self.api_key:
            return True
        try:
            self._validate_aws_credentials()
            return True
        except (BedrockAuthenticationError, ImportError):
            return False

    @staticmethod
    def _is_authentication_error(exc: BaseException) -> bool:
        if isinstance(exc, BedrockAuthenticationError):
            return True
        if getattr(exc, "status_code", None) in {401, 403}:
            return True
        message = str(exc).lower()
        return any(
            marker in message
            for marker in (
                "no aws credentials",
                "unable to locate credentials",
                "invalid security token",
                "expiredtoken",
                "unrecognizedclient",
            )
        )

    def _validate_model(self, model_id: str) -> None:
        if model_id not in Config.SUPPORTED_MODELS:
            raise ValueError(
                f"Unsupported Bedrock Mantle model '{model_id}'. "
                f"Supported models: {', '.join(Config.SUPPORTED_MODELS)}"
            )

    def _converse_once(self, **kwargs) -> Dict[str, Any]:
        """Make one provider call without retry or model fallback."""
        model_id = kwargs.get("modelId", "")
        system = kwargs.get("system", [])
        messages = kwargs.get("messages", [])
        inference_config = kwargs.get("inferenceConfig", {})

        try:
            if model_id.startswith("openai."):
                return self._converse_openai(
                    model_id,
                    system,
                    messages,
                    inference_config,
                )
            if model_id.startswith("anthropic."):
                return self._converse_anthropic(
                    model_id,
                    system,
                    messages,
                    inference_config,
                )
            raise ValueError(f"No Bedrock Mantle provider configured for '{model_id}'.")
        except Exception as exc:
            if self._is_authentication_error(exc):
                if isinstance(exc, BedrockAuthenticationError):
                    raise
                raise BedrockAuthenticationError(
                    "Amazon Bedrock Mantle authentication failed. Refresh your "
                    "AWS credentials or set AWS_BEARER_TOKEN_BEDROCK."
                ) from exc
            raise

    @bedrock_retry
    def _converse_with_retry(self, **kwargs) -> Dict[str, Any]:
        return self._converse_once(**kwargs)

    def _fallback_model_for(self, model_id: str) -> Optional[str]:
        if not Config.ENABLE_MODEL_FALLBACK:
            return None
        fallback_model = Config.FALLBACK_MODEL_ID
        if not fallback_model or fallback_model == model_id:
            return None
        if fallback_model not in Config.SUPPORTED_MODELS:
            logger.error(
                "Ignoring unsupported MANTLE_FALLBACK_MODEL_ID '%s'",
                fallback_model,
            )
            return None
        return fallback_model

    def converse(self, **kwargs) -> Dict[str, Any]:
        """Call Mantle, falling back after retryable model-capacity failures."""
        model_id = kwargs.get("modelId", "")
        self._validate_model(model_id)

        try:
            return self._converse_with_retry(**kwargs)
        except Exception as exc:
            fallback_model = self._fallback_model_for(model_id)
            if fallback_model is None or not is_model_fallback_error(exc):
                raise

            status_code = get_status_code(exc)
            reason = f"HTTP {status_code}" if status_code else type(exc).__name__
            logger.warning(
                "Model %s remained unavailable after retries (%s); "
                "falling back to %s for this request",
                model_id,
                reason,
                fallback_model,
            )
            fallback_kwargs = dict(kwargs)
            fallback_kwargs["modelId"] = fallback_model
            return self._converse_with_retry(**fallback_kwargs)
