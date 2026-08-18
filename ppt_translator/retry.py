"""
Retry policy for Amazon Bedrock API calls.

Centralizes which exceptions are worth retrying (throttling, transient
server errors, network hiccups) across Bedrock Runtime and Mantle SDKs.
Uses tenacity for exponential backoff.

If tenacity is not installed, the `bedrock_retry` decorator degrades
to a no-op so the package still imports.
"""
import logging
import os

logger = logging.getLogger(__name__)

# Bedrock error codes that are worth retrying.
_RETRYABLE_CODES = {
    'ThrottlingException',
    'ServiceUnavailableException',
    'ModelStreamErrorException',
    'InternalServerException',
    'ModelTimeoutException',
    'ModelErrorException',
}

# Explicit non-retryable codes — retrying these just wastes tokens/time.
_NON_RETRYABLE_CODES = {
    'ValidationException',
    'AccessDeniedException',
    'ResourceNotFoundException',
    'ModelNotReadyException',
    'UnauthorizedOperation',
}


def get_status_code(exc: BaseException):
    """Extract an HTTP status code from OpenAI, Anthropic, or botocore errors."""
    status_code = getattr(exc, 'status_code', None)
    if status_code is None:
        response = getattr(exc, 'response', None)
        status_code = getattr(response, 'status_code', None)
        if status_code is None and isinstance(response, dict):
            status_code = (
                response.get('ResponseMetadata', {})
                .get('HTTPStatusCode')
            )
    if status_code is not None:
        try:
            return int(status_code)
        except (TypeError, ValueError):
            return None
    return None


def _bedrock_error_code(exc: BaseException) -> str:
    response = getattr(exc, 'response', None)
    if not isinstance(response, dict):
        return ""
    return str(response.get('Error', {}).get('Code', ''))


def is_model_fallback_error(exc: BaseException) -> bool:
    """Return True when retry exhaustion may be helped by another model."""
    status_code = get_status_code(exc)
    message = str(exc).lower()

    # A different model does not resolve account quota or billing limits.
    if status_code == 429 and any(
        marker in message
        for marker in ('insufficient_quota', 'billing', 'usage limit', 'quota')
    ):
        return False

    if status_code in {429, 503}:
        return True

    error_code = _bedrock_error_code(exc)
    if error_code in {'ThrottlingException', 'ServiceUnavailableException'}:
        return True

    exception_name = type(exc).__name__.lower()
    return (
        'ratelimit' in exception_name
        or 'serviceunavailable' in exception_name
    )


def is_retryable(exc: BaseException) -> bool:
    """Return True if the given exception should trigger a retry."""
    status_code = get_status_code(exc)
    if status_code in {408, 409, 429, 500, 502, 503, 504}:
        return True
    if status_code in {400, 401, 403, 404, 422}:
        return False

    exception_name = type(exc).__name__.lower()
    if any(token in exception_name for token in ('timeout', 'connection', 'ratelimit')):
        return True

    try:
        from botocore.exceptions import ClientError, ReadTimeoutError, EndpointConnectionError, ConnectTimeoutError
    except ImportError:
        return False

    if isinstance(exc, (ReadTimeoutError, EndpointConnectionError, ConnectTimeoutError)):
        return True

    if isinstance(exc, ClientError):
        code = _bedrock_error_code(exc)
        if code in _NON_RETRYABLE_CODES:
            return False
        return code in _RETRYABLE_CODES

    return False


def _build_retry_decorator():
    """Build the tenacity retry decorator, or a no-op if tenacity isn't available."""
    try:
        from tenacity import (
            retry, stop_after_attempt, wait_exponential,
            retry_if_exception, before_sleep_log,
        )
    except ImportError:
        logger.warning(
            "tenacity not installed; Bedrock retries disabled. "
            "Install with `pip install tenacity` to enable automatic retry on throttling."
        )

        def _noop(func):
            return func
        return _noop

    try:
        max_attempts = int(os.getenv('BEDROCK_MAX_RETRIES', '5'))
    except ValueError:
        max_attempts = 5

    return retry(
        stop=stop_after_attempt(max_attempts),
        wait=wait_exponential(multiplier=2, min=1, max=30),
        retry=retry_if_exception(is_retryable),
        before_sleep=before_sleep_log(logger, logging.WARNING),
        reraise=True,
    )


bedrock_retry = _build_retry_decorator()
