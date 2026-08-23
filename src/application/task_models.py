"""Unified models for task status, progress, cancellation, and execution outcomes."""

from __future__ import annotations

import threading
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Generic, TypeVar

T = TypeVar("T")


class TaskState(Enum):
    IDLE = "idle"
    RUNNING = "running"
    SUCCESS = "success"
    WARNING = "warning"
    FAILED = "failed"
    CANCELLED = "cancelled"


class CancellationToken:
    """Thread-safe cancellation token for long-running batch tasks."""

    def __init__(self):
        self._cancelled = threading.Event()

    def cancel(self) -> None:
        self._cancelled.set()

    @property
    def is_cancelled(self) -> bool:
        return self._cancelled.is_set()

    def reset(self) -> None:
        self._cancelled.clear()


@dataclass
class TaskProgress:
    current: int = 0
    total: int = 0
    message: str = ""
    percentage: int = 0
    data: Any = None

    @classmethod
    def calculate(cls, current: int, total: int, message: str = "", data: Any = None) -> TaskProgress:
        pct = int((current / total * 100)) if total > 0 else 0
        return cls(current=current, total=total, message=message, percentage=pct, data=data)


@dataclass
class ServiceResult(Generic[T]):
    success: bool
    data: T | None = None
    error: str = ""
    state: TaskState = TaskState.SUCCESS
    metrics: dict[str, Any] = field(default_factory=dict)
