"""Cross-cutting decorators: retry with backoff, perf measurement."""
from __future__ import annotations

import functools
import time
from typing import Any, Callable, TypeVar

import structlog

F = TypeVar("F", bound=Callable[..., Any])

log = structlog.get_logger().bind(component="decorators")


def retry(*, max_attempts: int = 3, backoff: float = 2.0, exceptions: tuple = (Exception,)) -> Callable[[F], F]:
    """Retry with exponential backoff. Re-raises the final exception if all attempts fail."""

    def deco(fn: F) -> F:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            delay = 1.0
            last: Exception | None = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return fn(*args, **kwargs)
                except exceptions as exc:
                    last = exc
                    if attempt == max_attempts:
                        log.warning(
                            "retry exhausted",
                            fn=fn.__name__,
                            attempts=attempt,
                            error=repr(exc),
                        )
                        raise
                    log.info(
                        "retry attempt failed",
                        fn=fn.__name__,
                        attempt=attempt,
                        delay=delay,
                        error=repr(exc),
                    )
                    time.sleep(delay)
                    delay *= backoff
            raise RuntimeError(f"unreachable; last={last!r}")

        return wrapper  # type: ignore[return-value]

    return deco


def measure_time(label: str | None = None) -> Callable[[F], F]:
    """Log wall-clock duration of the wrapped function as `duration_ms`."""

    def deco(fn: F) -> F:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            start = time.perf_counter()
            try:
                return fn(*args, **kwargs)
            finally:
                ms = (time.perf_counter() - start) * 1000.0
                log.info(
                    "timing",
                    fn=label or fn.__name__,
                    duration_ms=round(ms, 1),
                )

        return wrapper  # type: ignore[return-value]

    return deco
