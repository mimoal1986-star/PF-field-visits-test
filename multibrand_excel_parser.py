"""
Парсер для файла распределения плана Мультибренд 2024
Ожидаемая структура файла:
- Вкладка "Дилеры_май" (название месяца динамическое)
- Вкладка "Пронто_май"
"""

import pandas as pd
import streamlit as st
import re


def extract_month_from_sheet_name(sheet_name):
    """
    Извлекает месяц из названия вкладки
    Пример: "Дилеры_май" -> "май"
    """
    match = re.search(r'_(.+)$', sheet_name)
    if match:
        return match.group(1)
    return None


def parse_multibrand_excel(file_obj):
    """
    Парсит Excel файл с распределением плана Мультибренд 2024
    
    Args:
        file_obj: загруженный Excel файл
    
    Returns:
        tuple: (dilers_df, pronto_df, month) 
               - dilers_df: DataFrame для Дилеры
               - pronto_df: DataFrame для Пронто
               - month: месяц (например, "май")
    """
    if file_obj is None:
        return pd.DataFrame(), pd.DataFrame(), None
    
    try:
        # Получаем все имена листов
        xl = pd.ExcelFile(file_obj)
        sheet_names = xl.sheet_names
        
        dilers_df = pd.DataFrame()
        pronto_df = pd.DataFrame()
        month = None
        
        # Ищем вкладки
        for sheet_name in sheet_names:
            sheet_lower = sheet_name.lower()
            
            # Вкладка Дилеры
            if 'дилер' in sheet_lower:
                month = extract_month_from_sheet_name(sheet_name)
                dilers_df = pd.read_excel(file_obj, sheet_name=sheet_name, dtype=str)
                st.info(f"✅ Найдена вкладка Дилеры: {sheet_name}, месяц: {month}")
            
            # Вкладка Пронто
            elif 'пронто' in sheet_lower:
                if month is None:
                    month = extract_month_from_sheet_name(sheet_name)
                pronto_df = pd.read_excel(file_obj, sheet_name=sheet_name, dtype=str)
                st.info(f"✅ Найдена вкладка Пронто: {sheet_name}, месяц: {month}")
        
        # Очистка данных
        if not dilers_df.empty:
            dilers_df = clean_dilers_table(dilers_df)
        
        if not pronto_df.empty:
            pronto_df = clean_pronto_table(pronto_df)
        
        return dilers_df, pronto_df, month
        
    except Exception as e:
        st.error(f"Ошибка при парсинге файла: {e}")
        return pd.DataFrame(), pd.DataFrame(), None


def clean_dilers_table(df):
    """
    Очистка и приведение таблицы Дилеры к стандартному формату
    """
    if df.empty:
        return df
    
    df_clean = df.copy()
    
    # Переименовываем колонки
    column_mapping = {
        'Обозначение': 'region_code',
        'Регион полный': 'region_full',
        'АСС': 'asm',
        'ЭМ': 'rs',
        'Дилеры': 'plan'
    }
    
    for old_name, new_name in column_mapping.items():
        if old_name in df_clean.columns:
            df_clean = df_clean.rename(columns={old_name: new_name})
    
    # Оставляем нужные колонки
    required_cols = ['region_code', 'region_full', 'asm', 'rs', 'plan']
    existing_cols = [col for col in required_cols if col in df_clean.columns]
    df_clean = df_clean[existing_cols]
    
    # Очищаем данные
    df_clean['region_code'] = df_clean['region_code'].astype(str).str.strip().str.upper()
    df_clean['plan'] = pd.to_numeric(df_clean['plan'], errors='coerce').fillna(0)
    
    # Для Дилеры всегда wave_type = 'Дилеры'
    df_clean['wave_type'] = 'Дилеры'
    
    # Удаляем пустые строки
    df_clean = df_clean[df_clean['region_code'].notna() & (df_clean['region_code'] != '')]
    df_clean = df_clean[df_clean['region_code'] != 'nan']
    
    return df_clean.reset_index(drop=True)

def clean_pronto_table(df):
    """
    Очистка и приведение таблицы Пронто к стандартному формату
    """
    if df.empty:
        return df
    
    df_clean = df.copy()
    
    # Переименовываем колонки
    column_mapping = {
        'Обозначение': 'region_code',
        'Регион полный': 'region_full',
        'Регион': 'wave_origin',
        'АСС': 'asm',
        'ЭМ': 'rs',
        'Пронто': 'plan'
    }
    
    for old_name, new_name in column_mapping.items():
        if old_name in df_clean.columns:
            df_clean = df_clean.rename(columns={old_name: new_name})
    
    # Оставляем нужные колонки
    required_cols = ['region_code', 'region_full', 'asm', 'rs', 'plan']
    existing_cols = [col for col in required_cols if col in df_clean.columns]
    df_clean = df_clean[existing_cols]
    
    # Очищаем данные
    df_clean['region_code'] = df_clean['region_code'].astype(str).str.strip().str.upper()
    df_clean['plan'] = pd.to_numeric(df_clean['plan'], errors='coerce').fillna(0)
    
    # Определяем wave_type на основе региона
    def get_wave_type(region_full):
        if 'МСК дистр.' in str(region_full):
            return 'Пронто М'
        else:
            return 'Пронто'
    
    df_clean['wave_type'] = df_clean['region_full'].apply(get_wave_type)
    
    # Удаляем пустые строки
    df_clean = df_clean[df_clean['region_code'].notna() & (df_clean['region_code'] != '')]
    df_clean = df_clean[df_clean['region_code'] != 'nan']
    
    return df_clean.reset_index(drop=True)


def preview_multibrand_plan(dilers_df, pronto_df):
    """
    Создает предпросмотр для отображения в интерфейсе
    
    Args:
        dilers_df: DataFrame таблицы Дилеры
        pronto_df: DataFrame таблицы Пронто
    
    Returns:
        DataFrame: объединенный предпросмотр
    """
    preview_data = []
    
    # Добавляем данные Дилеры
    if not dilers_df.empty:
        for _, row in dilers_df.iterrows():
            preview_data.append({
                'Тип': 'Дилеры',
                'Регион': row.get('region_code', ''),
                'Регион полный': row.get('region_full', ''),
                'АСС': row.get('asm', ''),
                'ЭМ': row.get('rs', ''),
                'План': row.get('plan', 0)
            })
    
    # Добавляем данные Пронто
    if not pronto_df.empty:
        for _, row in pronto_df.iterrows():
            preview_data.append({
                'Тип': 'Пронто',
                'Регион': row.get('region_code', ''),
                'Регион полный': row.get('region_full', ''),
                'АСС': row.get('asm', ''),
                'ЭМ': row.get('rs', ''),
                'План': row.get('plan', 0)
            })
    
    if not preview_data:
        return pd.DataFrame()
    
    return pd.DataFrame(preview_data)
