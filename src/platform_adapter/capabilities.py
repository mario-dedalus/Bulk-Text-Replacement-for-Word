"""Platform capabilities detection and runtime environment query."""

from __future__ import annotations

import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class PlatformCapabilities:
    is_windows: bool
    is_macos: bool
    is_linux: bool
    has_word_com: bool
    platform_name: str

    @classmethod
    def detect(cls) -> PlatformCapabilities:
        system = sys.platform
        is_windows = system.startswith("win")
        is_macos = system == "darwin"
        is_linux = system.startswith("linux")

        has_word_com = False
        if is_windows:
            try:
                import win32com.client  # type: ignore # noqa: F401
                has_word_com = True
            except ImportError:
                has_word_com = False

        if is_windows:
            name = "Windows"
        elif is_macos:
            name = "macOS"
        elif is_linux:
            name = "Linux"
        else:
            name = system

        return cls(
            is_windows=is_windows,
            is_macos=is_macos,
            is_linux=is_linux,
            has_word_com=has_word_com,
            platform_name=name,
        )


CAPABILITIES = PlatformCapabilities.detect()
