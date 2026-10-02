# -*- coding: utf-8 -*-
"""라이트/다크 공통 디자인 토큰.

두 테마는 같은 QSS 템플릿(`stylesheet.build_stylesheet`)을 쓰고 색 토큰만 다르다.
변환기 앱 성격에 맞게 중립 회색 바탕 + 단일 파란 강조색을 사용한다.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemePalette:
    name: str
    window_bg: str
    surface: str
    surface_alt: str
    input_bg: str
    border: str
    border_strong: str
    text: str
    text_muted: str
    text_disabled: str
    accent: str
    accent_hover: str
    accent_pressed: str
    accent_soft: str
    on_accent: str
    danger: str
    danger_hover: str
    danger_soft: str
    warning_text: str
    warning_bg: str
    warning_border: str
    success: str
    # 토스트는 앱 배경과 확실히 구분되는 불투명 패널을 쓴다.
    toast_bg: str
    toast_border: str
    toast_text: str
    shadow_alpha: int


LIGHT_PALETTE = ThemePalette(
    name="light",
    window_bg="#f3f5f9",
    surface="#ffffff",
    surface_alt="#f6f7fa",
    input_bg="#ffffff",
    border="#dce1e9",
    border_strong="#c2c9d6",
    text="#1f2533",
    text_muted="#5f6b7c",
    text_disabled="#a4acb9",
    accent="#2f6fed",
    accent_hover="#2560d8",
    accent_pressed="#1e50b8",
    accent_soft="#e7effe",
    on_accent="#ffffff",
    danger="#d63c4a",
    danger_hover="#bf2f3d",
    danger_soft="#fdecee",
    warning_text="#8a5a00",
    warning_bg="#fff6e0",
    warning_border="#f1d48b",
    success="#1f9d55",
    toast_bg="#1f2533",
    toast_border="#1f2533",
    toast_text="#ffffff",
    shadow_alpha=70,
)

DARK_PALETTE = ThemePalette(
    name="dark",
    window_bg="#14161b",
    surface="#1c1f26",
    surface_alt="#232731",
    input_bg="#171a20",
    border="#2e333e",
    border_strong="#414858",
    text="#e7e9ee",
    text_muted="#9aa3b2",
    text_disabled="#5c6474",
    accent="#4c8dff",
    accent_hover="#6a9fff",
    accent_pressed="#3a78e6",
    accent_soft="#22324f",
    on_accent="#ffffff",
    danger="#ef5d6a",
    danger_hover="#f47984",
    danger_soft="#3a2228",
    warning_text="#f3c565",
    warning_bg="#2e2716",
    warning_border="#5a4920",
    success="#3ecf7a",
    toast_bg="#2b303b",
    toast_border="#454d5e",
    toast_text="#f5f6f8",
    shadow_alpha=150,
)


def palette_for(theme_name: str) -> ThemePalette:
    return DARK_PALETTE if theme_name == "dark" else LIGHT_PALETTE
