"""Logging helpers for keeping third-party SDK noise out of normal output."""
import logging


_NOISY_DEPENDENCY_LOGGERS = (
    "botocore.credentials",
    "httpx",
    "httpx2",
    "httpcore",
    "httpcore2",
)


class _DependencyNoiseFilter(logging.Filter):
    """Drop routine dependency logs while preserving warnings and errors."""

    _ppt_translator_dependency_filter = True

    def filter(self, record: logging.LogRecord) -> bool:
        if record.levelno >= logging.WARNING:
            return True
        return not any(
            record.name == logger_name
            or record.name.startswith(f"{logger_name}.")
            for logger_name in _NOISY_DEPENDENCY_LOGGERS
        )


def _install_filter(target) -> None:
    if any(
        getattr(existing, "_ppt_translator_dependency_filter", False)
        for existing in target.filters
    ):
        return
    target.addFilter(_DependencyNoiseFilter())


def quiet_dependency_logs() -> None:
    """Suppress routine dependency INFO logs even if SDKs reset logger levels."""
    for logger_name in _NOISY_DEPENDENCY_LOGGERS:
        dependency_logger = logging.getLogger(logger_name)
        dependency_logger.setLevel(logging.WARNING)
        _install_filter(dependency_logger)

    for handler in logging.getLogger().handlers:
        _install_filter(handler)


def quiet_batch_worker_logs() -> None:
    """Keep child-process INFO logs from corrupting the parent Rich display."""
    quiet_dependency_logs()
    logging.getLogger("ppt_translator").setLevel(logging.WARNING)
