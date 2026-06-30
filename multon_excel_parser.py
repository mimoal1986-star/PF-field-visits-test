"""
Парсер Excel файла с распределением плана Мултон
"""

import pandas as pd
import streamlit as st
import re


def parse_multon_excel_to_df(excel_file) -> pd.DataFrame:
    """
    Парсит Excel с распределением плана Мултон
    
    Формат Excel:
    - Первая колонка: "Названия строк" (город)
    - Вторая колонка: "регион" (например "VG - Волгоградская область")
    - Третья колонка: "RS" (ФИО сотрудника)
    - Остальные колонки: коды проектов (начинаются с RU00.)
    
    Параметры:
        excel_file: загруженный файл (BytesIO)
    
    Возвращает:
        DataFrame с колонками:
            - project_code: код проекта
            - region: код региона (первые 2 символа, например "VG")
            - rs: ФИО сотрудника
            - plan: план в штуках (int)
    """
    if excel_file is None:
        st.warning("Файл не загружен")
        return pd.DataFrame()
    
    try:
        # Читаем Excel
        df = pd.read_excel(excel_file)
        
        if df.empty:
            st.warning("Файл пуст")
            return pd.DataFrame()
        
        # Определяем колонки с кодами проектов (начинаются с RU00.)
        project_cols = []
        for col in df.columns:
            col_str = str(col).strip()
            if col_str.startswith('RU00.'):
                project_cols.append(col_str)
        
        if not project_cols:
            st.warning("В файле не найдены колонки с кодами проектов (должны начинаться с RU00.)")
            return pd.DataFrame()
        
        # Находим колонку с регионом
        region_col = None
        for col in df.columns:
            if 'регион' in str(col).lower():
                region_col = col
                break
        
        if region_col is None:
            st.warning("Не найдена колонка 'регион'")
            return pd.DataFrame()
        
        # Находим колонку с RS
        rs_col = None
        for col in df.columns:
            if str(col).strip().upper() == 'RS':
                rs_col = col
                break
        
        if rs_col is None:
            st.warning("Не найдена колонка 'RS'")
            return pd.DataFrame()
        
        # Собираем данные
        records = []
        skipped_rows = 0
        
        for idx, row in df.iterrows():
            # Извлекаем регион (первые 2 символа)
            region_value = str(row[region_col]).strip()
            if not region_value or region_value in ['nan', 'None', '']:
                skipped_rows += 1
                continue
            
            # Извлекаем код региона (первые 2 символа)
            # Пример: "VG - Волгоградская область" → "VG"
            region_match = re.match(r'^([A-Z]{2})', region_value)
            if region_match:
                region_code = region_match.group(1)
            else:
                region_code = region_value[:2].upper()
            
            # Извлекаем RS
            rs_value = str(row[rs_col]).strip()
            if not rs_value or rs_value in ['nan', 'None', '']:
                skipped_rows += 1
                continue
            
            # Для каждого проекта берем план
            for project_col in project_cols:
                plan_value = row[project_col]
                
                # Пропускаем пустые значения
                if pd.isna(plan_value):
                    continue
                
                # Преобразуем в число
                try:
                    plan = float(plan_value)
                    if plan <= 0:
                        continue
                except (ValueError, TypeError):
                    continue
                
                records.append({
                    'project_code': str(project_col).strip(),
                    'region': region_code,
                    'rs': rs_value,
                    'plan': plan
                })
        
        if not records:
            st.warning(f"Не удалось извлечь данные. Пропущено строк: {skipped_rows}")
            return pd.DataFrame()
        
        result_df = pd.DataFrame(records)
        
        # Убираем дубликаты (если одинаковый проект+регион+RS)
        result_df = result_df.drop_duplicates(
            subset=['project_code', 'region', 'rs'],
            keep='first'
        ).reset_index(drop=True)
        
        st.success(f"✅ Загружено {len(result_df)} записей (проект + регион + RS)")
        
        return result_df
        
    except Exception as e:
        st.error(f"Ошибка при парсинге Excel: {e}")
        return pd.DataFrame()


def preview_multon_plan(df: pd.DataFrame) -> pd.DataFrame:
    """
    Возвращает DataFrame для предпросмотра в UI
    
    Параметры:
        df: DataFrame от parse_multon_excel_to_df
    
    Возвращает:
        DataFrame с колонками: Проект, Регион, RS, План
    """
    if df.empty:
        return pd.DataFrame()
    
    preview_df = df.copy()
    preview_df = preview_df.rename(columns={
        'project_code': 'Проект',
        'region': 'Регион',
        'rs': 'RS',
        'plan': 'План'
    })
    
    return preview_df
