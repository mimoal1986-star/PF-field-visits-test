"""
Парсер Excel файла с распределением RS для Optima
"""

import pandas as pd
import streamlit as st
import re


def parse_optima_rs_excel(excel_file):
    """
    Парсит Excel с распределением RS для Optima
    
    Формат Excel:
    - Колонка 'регион': код региона (AR, AS, KG...)
    - Колонка 'ФИО ЭМ': RS для этого региона
    - Колонка 'Москва': название клиента для Москвы
    - Колонка 'ЭМ' (под Москвой): RS для этого клиента
    - Колонка 'распределение по Питеру': название клиента для СПб
    - Колонка 'эм' (под Питером): RS для этого клиента
    
    Параметры:
        excel_file: загруженный файл (BytesIO)
    
    Возвращает:
        tuple: (region_mapping, moscow_mapping, spb_mapping)
    """
    if excel_file is None:
        st.warning("Файл не загружен")
        return {}, {}, {}
    
    try:
        # Читаем Excel
        df = pd.read_excel(excel_file)
        
        if df.empty:
            st.warning("Файл пуст")
            return {}, {}, {}
        
        # Определяем колонки
        region_col = None
        for col in df.columns:
            if 'регион' in str(col).lower():
                region_col = col
                break
        
        rs_col = None
        for col in df.columns:
            if col == 'ФИО ЭМ':
                rs_col = col
                break
        
        if region_col is None or rs_col is None:
            st.warning("Не найдены колонки 'регион' и/или 'ФИО ЭМ'")
            return {}, {}, {}
        
        # 1. Маппинг по регионам
        region_mapping = {}
        for _, row in df.iterrows():
            region_value = str(row[region_col]).strip()
            rs_value = str(row[rs_col]).strip()
            
            if not region_value or region_value in ['nan', 'None', '']:
                continue
            if not rs_value or rs_value in ['nan', 'None', '']:
                continue
            
            # Извлекаем код региона (первые 2 символа)
            region_match = re.match(r'^([A-Z]{2})', region_value)
            if region_match:
                region_code = region_match.group(1)
                region_mapping[region_code] = rs_value
        
        # 2. Маппинг по Москве
        moscow_mapping = {}
        moscow_client_col = None
        moscow_rs_col = None
        
        for col in df.columns:
            if 'Москва' in str(col):
                moscow_client_col = col
            if col == 'ЭМ' and moscow_client_col is not None:
                moscow_rs_col = col
        
        if moscow_client_col and moscow_rs_col:
            for _, row in df.iterrows():
                client = str(row[moscow_client_col]).strip()
                rs = str(row[moscow_rs_col]).strip()
                
                if client and client not in ['nan', 'None', '']:
                    if rs and rs not in ['nan', 'None', '']:
                        moscow_mapping[client] = rs
        
        # 3. Маппинг по Санкт-Петербургу
        spb_mapping = {}
        spb_client_col = None
        spb_rs_col = None
        
        for col in df.columns:
            if 'Питер' in str(col) or 'распределение по Питеру' in str(col):
                spb_client_col = col
            if col == 'эм' and spb_client_col is not None:
                spb_rs_col = col
        
        if spb_client_col and spb_rs_col:
            for _, row in df.iterrows():
                client = str(row[spb_client_col]).strip()
                rs = str(row[spb_rs_col]).strip()
                
                if client and client not in ['nan', 'None', '']:
                    if rs and rs not in ['nan', 'None', '']:
                        spb_mapping[client] = rs
        
        # Статистика
        st.success(f"✅ Загружено: {len(region_mapping)} регионов, {len(moscow_mapping)} клиентов (Москва), {len(spb_mapping)} клиентов (СПб)")
        
        return region_mapping, moscow_mapping, spb_mapping
        
    except Exception as e:
        st.error(f"Ошибка при парсинге Excel: {e}")
        return {}, {}, {}


def preview_optima_rs_mapping(region_mapping: dict, moscow_mapping: dict, spb_mapping: dict):
    """Возвращает DataFrame для предпросмотра в UI"""
    records = []
    
    for region, rs in region_mapping.items():
        records.append({'Тип': 'Регион', 'Ключ': region, 'RS': rs})
    
    for client, rs in moscow_mapping.items():
        records.append({'Тип': 'Москва (клиент)', 'Ключ': client, 'RS': rs})
    
    for client, rs in spb_mapping.items():
        records.append({'Тип': 'Санкт-Петербург (клиент)', 'Ключ': client, 'RS': rs})
    
    return pd.DataFrame(records)