"""
Translation cache — pluggable backends.

Backends share a minimal interface (`get` / `set` / `close`). Choose the
backend at the CLI boundary with `--cache-backend sqlite|memory|none`.

SQLite (default) serializes access when shared by slide translation threads.
Across ProcessPoolExecutor workers, each process must still open its own
connection. Never share a sqlite3.Connection across a fork boundary.
"""
import hashlib
import logging
import sqlite3
import threading
import time
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


def make_cache_key(source_text: str, target_language: str, model_id: str,
                   enable_polishing: bool, glossary_hash: str,
                   source_language: Optional[str] = None) -> str:
    """Build a deterministic cache key for one translation request.

    All inputs that could affect the translation output must participate,
    otherwise a stale hit returns the wrong language / style / terminology.
    """
    parts = [
        source_text,
        target_language,
        source_language or 'auto',
        model_id,
        'polish' if enable_polishing else 'literal',
        glossary_hash or 'none',
    ]
    payload = ''.join(parts).encode('utf-8')
    return hashlib.sha256(payload).hexdigest()


class TranslationCache(ABC):
    """Abstract cache backend."""

    @abstractmethod
    def get(self, key: str) -> Optional[str]:
        ...

    @abstractmethod
    def set(self, key: str, value: str) -> None:
        ...

    def close(self) -> None:  # pragma: no cover - optional hook
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False


class NullCache(TranslationCache):
    """No-op cache. Useful for `--no-cache` or dry-run paths."""

    def get(self, key: str) -> Optional[str]:
        return None

    def set(self, key: str, value: str) -> None:
        return None


class InMemoryCache(TranslationCache):
    """Process-local dict cache. Lost when the process exits."""

    def __init__(self):
        self._store: dict = {}
        self._lock = threading.RLock()

    def get(self, key: str) -> Optional[str]:
        with self._lock:
            return self._store.get(key)

    def set(self, key: str, value: str) -> None:
        with self._lock:
            self._store[key] = value


class SQLiteCache(TranslationCache):
    """Thread-safe file cache; each process must open its own instance."""

    def __init__(self, path: str):
        self.path = str(Path(path).expanduser())
        self._lock = threading.RLock()
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        # isolation_level=None → autocommit; simpler than manual transactions.
        self._conn = sqlite3.connect(self.path, isolation_level=None, check_same_thread=False)
        # WAL lets multiple processes read concurrently while one writes.
        self._conn.execute("PRAGMA journal_mode=WAL")
        self._conn.execute("PRAGMA synchronous=NORMAL")
        self._conn.execute("PRAGMA busy_timeout=5000")
        self._conn.execute(
            """
            CREATE TABLE IF NOT EXISTS cache (
              k TEXT PRIMARY KEY,
              v TEXT NOT NULL,
              created_at REAL NOT NULL
            )
            """
        )

    def get(self, key: str) -> Optional[str]:
        try:
            with self._lock:
                row = self._conn.execute(
                    "SELECT v FROM cache WHERE k = ?",
                    (key,),
                ).fetchone()
            return row[0] if row else None
        except sqlite3.Error as e:
            logger.debug(f"SQLite get failed: {e}")
            return None

    def set(self, key: str, value: str) -> None:
        try:
            with self._lock:
                self._conn.execute(
                    "INSERT OR REPLACE INTO cache (k, v, created_at) VALUES (?, ?, ?)",
                    (key, value, time.time()),
                )
        except sqlite3.Error as e:
            logger.debug(f"SQLite set failed: {e}")

    def close(self) -> None:
        try:
            with self._lock:
                self._conn.close()
        except sqlite3.Error:
            pass


def build_cache(backend: str, path: Optional[str] = None) -> TranslationCache:
    """Factory. Called from CLI and (importantly) inside worker processes.

    Never hand a pre-built cache across a ProcessPool boundary; call
    build_cache() from inside the worker so each gets its own connection.
    """
    backend = (backend or 'none').lower()
    if backend in ('none', 'off', 'disabled', ''):
        return NullCache()
    if backend == 'memory':
        return InMemoryCache()
    if backend == 'sqlite':
        default_path = str(Path.home() / '.ppt-translator' / 'cache.db')
        return SQLiteCache(path or default_path)
    logger.warning(f"Unknown cache backend '{backend}', falling back to none")
    return NullCache()
