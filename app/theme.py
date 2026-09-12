# -*- coding: utf-8 -*-
"""Цветовая схема и константы интерфейса."""

NAVY = "#1A4C8C"
NAVY_HOVER = "#2563A8"
TEAL = "#0F766E"
TEAL_HOVER = "#0D9488"
BG = "#F4F7FB"
CARD = "#FFFFFF"
BORDER = "#D8E0EA"
TEXT = "#0F172A"
MUTED = "#64748B"
SUCCESS = "#059669"
SUCCESS_BG = "#D1FAE5"
WARNING = "#D97706"
WARNING_BG = "#FEF3C7"
DANGER = "#DC2626"
DANGER_BG = "#FEE2E2"
ACCENT_SOFT = "#E8F0FA"

FONT_FAMILY = "Segoe UI"

WINDOW_SIZE = "1000x720"
WINDOW_MIN = (900, 640)

DEFAULT_OUTPUT_NAME = "Сводный_отчёт_мотивирующий_мониторинг.xlsx"

REQUIRED_COLUMNS = {
    "timeliness": ["Регион", "Район", "Пищеблок", "Процент Своевременности"],
    "typical_menu": [
        "Пищеблок Наименование",
        "Район",
        "Дата",
        "Число Несоблюдений",
        "Процент Соблюдения",
        "Замечания",
    ],
    "sanpin": ["Район", "Пищеблок", "Дата", "Прием Пищи", "Ошибка"],
}

FILE_LABELS = {
    "timeliness": "Своевременность размещения меню",
    "typical_menu": "Соответствие типовому меню",
    "sanpin": "Ошибки по СанПиН",
}
