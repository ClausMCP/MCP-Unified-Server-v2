#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCP Rate Limiter + Circuit Breaker v1.5
- Исправлена ошибка AttributeError: 'sqlite3.Connection' object has no attribute 'rowcount'
- Удалена неиспользуемая и неработающая функция cleanup_old_records (таблица rate_log не существовала)
- Персистентность в SQLite, Retry-After, авто health-check
"""

import time
import threading
import sqlite3
import json
from pathlib import Path
from typing import Tuple, Optional, Dict, Any

from mcp_shared import _log

DB_PATH = Path(__file__).parent / "mcp_rate_limiter.db"

def _init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS rate_history (
            service TEXT PRIMARY KEY,
            timestamps TEXT,
            failures INTEGER DEFAULT 0,
            last_failure REAL DEFAULT 0,
            last_success REAL DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()

_init_db()


class PersistentRateLimiter:
    def __init__(self, max_calls: int = 25, window_sec: float = 60.0):
        self.max_calls = max_calls
        self.window = window_sec
        self._lock = threading.Lock()

    def allow(self, service: str) -> Tuple[bool, Optional[float]]:
        now = time.monotonic()
        with self._lock:
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute("BEGIN IMMEDIATE")
                row = conn.execute("SELECT timestamps FROM rate_history WHERE service = ?", (service,)).fetchone()
                timestamps = json.loads(row[0]) if row and row[0] else []
                timestamps = [ts for ts in timestamps if ts > now - self.window]
                if len(timestamps) >= self.max_calls:
                    retry_after = (timestamps[0] + self.window - now) if timestamps else self.window
                    conn.rollback()
                    return False, round(max(0.0, retry_after), 1)
                timestamps.append(now)
                conn.execute("REPLACE INTO rate_history (service, timestamps) VALUES (?, ?)",
                             (service, json.dumps(timestamps)))
                conn.commit()
                return True, None
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

    def reset(self, service: str):
        """Удаляет всю историю вызовов для указанного сервиса."""
        with self._lock:
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute("DELETE FROM rate_history WHERE service = ?", (service,))
                conn.commit()
            finally:
                conn.close()


class EnhancedCircuitBreaker:
    def __init__(self, failure_threshold: int = 5, recovery_timeout: float = 45.0):
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self._lock = threading.Lock()

    def can_execute(self, service: str) -> bool:
        with self._lock:
            conn = sqlite3.connect(DB_PATH)
            try:
                row = conn.execute("SELECT failures, last_failure FROM rate_history WHERE service = ?", (service,)).fetchone()
                failures = row[0] if row else 0
                last_failure = row[1] if row else 0
                now = time.monotonic()
                if now - last_failure > self.recovery_timeout:
                    return True
                return failures < self.failure_threshold
            finally:
                conn.close()

    def record_success(self, service: str):
        with self._lock:
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute("UPDATE rate_history SET failures = 0, last_success = ? WHERE service = ?",
                             (time.monotonic(), service))
                conn.commit()
            finally:
                conn.close()

    def record_failure(self, service: str):
        with self._lock:
            conn = sqlite3.connect(DB_PATH)
            try:
                conn.execute("BEGIN IMMEDIATE")
                row = conn.execute("SELECT failures FROM rate_history WHERE service = ?", (service,)).fetchone()
                failures = (row[0] if row else 0) + 1
                conn.execute("REPLACE INTO rate_history (service, failures, last_failure) VALUES (?, ?, ?)",
                             (service, failures, time.monotonic()))
                conn.commit()
                if failures >= self.failure_threshold:
                    self._send_alert(service, failures)
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

    def _send_alert(self, service: str, failures: int):
        alert_msg = f"⚠️ CIRCUIT BREAKER OPENED → {service} | Failures: {failures}/{self.failure_threshold}"
        _log(alert_msg)


rate_limiter = PersistentRateLimiter(max_calls=25, window_sec=60)
circuit_breaker = EnhancedCircuitBreaker(failure_threshold=5, recovery_timeout=45)


def safe_call(service: str, func, *args, **kwargs):
    """Универсальная безопасная обёртка для вызовов внешних сервисов."""
    allowed, retry_after = rate_limiter.allow(service)
    if not allowed:
        return {"error": f"Rate limit exceeded for {service}", "retry_after": retry_after}
    if not circuit_breaker.can_execute(service):
        return {"error": f"Сервис {service} временно отключён (Circuit Breaker)", "retry_after": 30}
    try:
        result = func(*args, **kwargs)
        circuit_breaker.record_success(service)
        return result
    except Exception as e:
        circuit_breaker.record_failure(service)
        return {"error": str(e), "service": service}

# Функция очистки устаревших записей удалена, так как таблица rate_log не используется.
# В текущей реализации rate_limiter и circuit breaker хранят только актуальные данные,
# и периодическая очистка не требуется. Если понадобится — можно добавить позже.