"""
Парсер Excel файла с коэффициентами регионов
"""

import pandas as pd
import streamlit as st

# Импортируем REGION_MAPPING из data_cleaner
from data_cleaner import REGION_MAPPING


def parse_percent(value: str) -> float:
    """Преобразует значение из Excel в коэффициент"""
    if pd.isna(value):
        return 1.0
    
    value_str = str(value).strip().replace(',', '.')
    
    try:
        return float(value_str)
    except:
        return 1.0


def parse_region_coefficients_excel(excel_file) -> dict:
    """
    Парсит Excel с коэффициентами регионов
    
    Формат Excel:
    - Строка 1: коды регионов (AA, AD, AL...)
    - Строка 2: коэффициенты (100%, 95%, 90%...)
    
    Параметры:
        excel_file: загруженный файл (BytesIO)
    
    Возвращает:
        dict: {код_региона: коэффициент (float)}
        Пример: {'AA': 1.0, 'BL': 0.95, 'BR': 0.9}
    """
    if excel_file is None:
        st.warning("Файл не загружен")
        return {}
    
    try:
        # Читаем Excel
        df = pd.read_excel(excel_file, header=None, dtype=str)
        
        if df.empty:
            st.warning("Файл пуст")
            return {}
        
        if len(df) < 2:
            st.warning("Файл должен содержать минимум 2 строки")
            return {}
        
        region_codes = df.iloc[0].dropna().astype(str).str.strip().tolist()
        coefficients_raw = df.iloc[1].dropna().astype(str).str.strip().tolist()
        
        if len(region_codes) != len(coefficients_raw):
            st.warning(f"Несоответствие количества регионов ({len(region_codes)}) и коэффициентов ({len(coefficients_raw)})")
            return {}
        
        result = {}
        for code, coeff_str in zip(region_codes, coefficients_raw):
            if not code or code in ['nan', 'None', '']:
                continue
            
            coeff_value = parse_percent(coeff_str)
            result[code] = coeff_value
        
        if not result:
            st.warning("Не удалось извлечь ни одного коэффициента")
            return {}
        
        st.success(f"✅ Загружено {len(result)} коэффициентов регионов")
        return result
        
    except Exception as e:
        st.error(f"Ошибка при парсинге Excel: {e}")
        return {}


def preview_region_coefficients(coefficients_dict: dict) -> pd.DataFrame:
    """
    Возвращает DataFrame для предпросмотра в UI
    
    Параметры:
        coefficients_dict: словарь {код_региона: коэффициент}
    
    Возвращает:
        DataFrame с колонками: Регион (код), Регион (название), Коэффициент, Регион определен
    """
    if not coefficients_dict:
        return pd.DataFrame()
    
    records = []
    for code, coeff in coefficients_dict.items():
        region_name = REGION_MAPPING.get(code, code)
        records.append({
            'Регион (код)': code,
            'Регион (название)': region_name,
            'Коэффициент': coeff,
            'Регион определен': 'Да' if code in REGION_MAPPING else 'Нет'
        })
    
    return pd.DataFrame(records)
