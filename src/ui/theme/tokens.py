"""Windows 11 Fluent Design & macOS design tokens."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ColorTokens:
    window_bg: str
    card_bg: str
    card_secondary: str
    text_primary: str
    text_secondary: str
    text_placeholder: str
    border: str
    border_subtle: str
    border_strong: str
    accent: str
    accent_hover: str
    accent_pressed: str
    accent_text: str
    button_bg: str
    button_hover: str
    button_pressed: str
    button_border: str
    entry_bg: str
    entry_border: str
    badge_bg: str
    error: str
    error_bg: str
    success: str
    success_bg: str
    warning: str
    warning_bg: str
    table_header_bg: str
    table_row_alt: str
    table_selected_bg: str
    log_bg: str
    log_fg: str


LIGHT_TOKENS = ColorTokens(
    window_bg="#F3F3F3",
    card_bg="#FFFFFF",
    card_secondary="#F9F9F9",
    text_primary="#1A1A1A",
    text_secondary="#5D5D5D",
    text_placeholder="#8A8A8A",
    border="#DADADA",
    border_subtle="#E5E5E5",
    border_strong="#BFBFBF",
    accent="#005FB8",
    accent_hover="#004E98",
    accent_pressed="#003D78",
    accent_text="#FFFFFF",
    button_bg="#FAFAFA",
    button_hover="#EAEAEA",
    button_pressed="#DFDFDF",
    button_border="#D1D1D1",
    entry_bg="#FFFFFF",
    entry_border="#CECECE",
    badge_bg="#E8EDF2",
    error="#C42B1C",
    error_bg="#FDF3F2",
    success="#0F7B0F",
    success_bg="#F1F9F1",
    warning="#9D5D00",
    warning_bg="#FFF9E6",
    table_header_bg="#F3F3F3",
    table_row_alt="#FAFAFA",
    table_selected_bg="#E5F1FB",
    log_bg="#FFFFFF",
    log_fg="#1E1E1E",
)

DARK_TOKENS = ColorTokens(
    window_bg="#202020",
    card_bg="#2B2B2B",
    card_secondary="#323232",
    text_primary="#FFFFFF",
    text_secondary="#C7C7C7",
    text_placeholder="#808080",
    border="#454545",
    border_subtle="#383838",
    border_strong="#555555",
    accent="#60CDFF",
    accent_hover="#4EB8EB",
    accent_pressed="#3DA3D6",
    accent_text="#000000",
    button_bg="#383838",
    button_hover="#454545",
    button_pressed="#2E2E2E",
    button_border="#4E4E4E",
    entry_bg="#202020",
    entry_border="#4E4E4E",
    badge_bg="#383838",
    error="#FF99A4",
    error_bg="#3D2022",
    success="#6CCB5F",
    success_bg="#1F3620",
    warning="#FCE100",
    warning_bg="#3B3615",
    table_header_bg="#303030",
    table_row_alt="#272727",
    table_selected_bg="#004A7C",
    log_bg="#181818",
    log_fg="#D4D4D4",
)
