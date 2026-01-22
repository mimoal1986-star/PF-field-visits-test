import os
from pathlib import Path

# Базовые пути
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"

# Пути к исходным файлам Excel
EXCEL_PATHS = {
    'массив': RAW_DATA_DIR / "Массив.xlsx",           # Портал + Автокодификация
    'гугл_таблица': RAW_DATA_DIR / "Гугл таблица.xlsx", # Проекты Сервизория
    'зод_асс': RAW_DATA_DIR / "ЗОД+АСС.xlsx",         # Иерархия
    'основной_отчет': RAW_DATA_DIR / "ИУ Аудиты ПланФакт v.1.7_5 .xlsx"  # Готовый отчет
}

# Создаем директории если их нет
for path in [DATA_DIR, RAW_DATA_DIR, PROCESSED_DATA_DIR]:
    path.mkdir(parents=True, exist_ok=True)

# Настройки кэширования
CACHE_SETTINGS = {
    'ttl': 3600,  # Время жизни кэша в секундах (1 час)
    'max_entries': 100,
    'show_spinner': True
}