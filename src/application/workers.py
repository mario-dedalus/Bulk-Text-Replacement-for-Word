"""Asynchronous background worker execution using QRunnable / QThreadPool with signature pre-inspection."""

from __future__ import annotations

import inspect
import sys
from typing import Any, Callable

from application.task_models import CancellationToken, TaskProgress

try:
    from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class WorkerSignals(QObject):
        """Defines signals available from a running worker thread."""

        started = Signal()
        progress = Signal(object)  # TaskProgress
        data = Signal(object)      # Intermediate data chunk
        result = Signal(object)    # Final ServiceResult or return value
        error = Signal(str)        # Error message string
        finished = Signal()        # Triggered on completion

    class TaskWorker(QRunnable):
        """Qt Runnable worker for executing callables asynchronously on QThreadPool."""

        def __init__(self, fn: Callable[..., Any], *args: Any, **kwargs: Any):
            super().__init__()
            self.fn = fn
            self.args = args
            self.kwargs = kwargs
            self.signals = WorkerSignals()
            self.cancel_token = CancellationToken()
            self.setAutoDelete(True)

            # Pre-inspect signature to avoid broad TypeError catch-and-retry
            self._inject_progress = False
            self._inject_cancel = False
            self._inspect_target_signature()

        def _inspect_target_signature(self) -> None:
            try:
                sig = inspect.signature(self.fn)
                params = sig.parameters
                has_var_keyword = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in params.values())

                if "progress_cb" in params or has_var_keyword:
                    self._inject_progress = True
                if "cancel_token" in params or has_var_keyword:
                    self._inject_cancel = True
            except (ValueError, TypeError):
                # Builtins or wrapped C-extensions
                self._inject_progress = False
                self._inject_cancel = False

        def cancel(self) -> None:
            self.cancel_token.cancel()

        def _handle_progress(self, *args: Any) -> None:
            if args:
                self.signals.progress.emit(args[0])

        def run(self) -> None:
            self.signals.started.emit()
            call_kwargs = dict(self.kwargs)
            if self._inject_progress and "progress_cb" not in call_kwargs:
                call_kwargs["progress_cb"] = self._handle_progress
            if self._inject_cancel and "cancel_token" not in call_kwargs:
                call_kwargs["cancel_token"] = self.cancel_token

            try:
                res = self.fn(*self.args, **call_kwargs)
                self.signals.result.emit(res)
            except Exception as exc:
                self.signals.error.emit(str(exc))
            finally:
                self.signals.finished.emit()

else:

    class TaskWorker:  # type: ignore
        """Fallback thread worker when PySide6 is not installed."""

        def __init__(self, fn: Callable[..., Any], *args: Any, **kwargs: Any):
            self.fn = fn
            self.args = args
            self.kwargs = kwargs
            self.cancel_token = CancellationToken()

        def cancel(self) -> None:
            self.cancel_token.cancel()

        def run(self) -> Any:
            return self.fn(*self.args, cancel_token=self.cancel_token, **self.kwargs)
