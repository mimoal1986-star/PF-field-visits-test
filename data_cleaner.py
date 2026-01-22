# utils/data_cleaner.py
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st
import io


class DataCleaner:
    """
    Очистка данных по инструкции для ИУ Аудиты
    """
    
    def clean_google(self, df):
        """
        Шаги 1-7: Очистка Гугл таблицы (Проекты Сервизория)
        """
        if df is None or df.empty:
            st.warning("⚠️ Гугл таблица пустая или не загружена")
            return None
        
        df_clean = df.copy()
        original_rows = len(df_clean)
        original_cols = len(df_clean.columns)
        
        st.info(f"🧹 Начинаю очистку Гугл таблицы: {original_rows} строк × {original_cols} колонок")
        
        # === ШАГ 1: Удалить дубликаты записей ===
        st.write("**1️⃣ Удаляю дубликаты записей...**")
        
        # Ищем поля с учетом реальных названий
        code_field = self._find_column(df_clean, [
            'Код проекта RU00.000.00.01SVZ24',  # Основное название
        ])
        
        start_date_field = self._find_column(df_clean, [
            'Дата старта', # Основное название
        ])
        
        end_date_field = self._find_column(df_clean, [
            'Дата финиша с продлением',  # Основное название
        ])
        
        # Собираем найденные поля
        existing_fields = []
        field_display_names = []
        
        if code_field:
            existing_fields.append(code_field)
            field_display_names.append('Код проекта')
            
        if start_date_field:
            existing_fields.append(start_date_field)
            field_display_names.append('Дата старта')
            
        if end_date_field:
            existing_fields.append(end_date_field)
            field_display_names.append('Дата финиша')
        
        # Проверяем сколько полей нашлось
        if len(existing_fields) == 3:
            before = len(df_clean)
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first')
            after = len(df_clean)
            removed = before - after
            
            if removed > 0:
                st.success(f"   ✅ Удалено {removed} дубликатов")
                st.info(f"   По полям: {', '.join(field_display_names)}")
                st.info(f"   Фактические имена: {', '.join(existing_fields)}")
            else:
                st.info("   ℹ️ Дубликатов не найдено")
                
        elif len(existing_fields) >= 1:
            st.warning(f"   ⚠️ Найдено только {len(existing_fields)} из 3 полей: {', '.join(field_display_names)}")
            
            # Все равно пытаемся удалить дубли по найденным полям
            before = len(df_clean)
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first')
            after = len(df_clean)
            removed = before - after
            
            if removed > 0:
                st.success(f"   ✅ Удалено {removed} дубликатов (по найденным полям)")
            else:
                st.info("   ℹ️ Дубликатов не найдено по найденным полям")
                
        else:
            st.warning("   ⚠️ Не найдено ни одного ключевого поля для проверки дубликатов")
            st.info("   Проверьте названия колонок в файле")
            
            # Показываем какие колонки есть
            st.info("   **Найденные колонки:**")
            for i, col in enumerate(df_clean.columns[:10]):  # Первые 10 колонок
                st.info(f"   {i+1}. {col}")
        
        # === ШАГ 2: Сжать пробелы в кодах проектов ===
        st.write("**2️⃣ Чищу пробелы в кодах проектов...**")
        
        code_col = self._find_column(df_clean, ['Код проекта', 'Код', 'Project Code', 'КодПроекта'])
        
        if code_col:
            # Сохраняем оригинальные значения
            original_codes = df_clean[code_col].copy()
            
            # Приводим к строке
            df_clean[code_col] = df_clean[code_col].astype(str)
            
            # ТОЛЬКО удаляем пробелы в начале и конце (по инструкции)
            df_clean[code_col] = df_clean[code_col].str.strip()
            
            # НЕ меняем внутренние пробелы!
            # df_clean[code_col] = df_clean[code_col].str.replace(r'\s+', ' ', regex=True)  # УБРАТЬ!
            
            # Считаем изменения
            changed = (original_codes.fillna('') != df_clean[code_col].fillna('')).sum()
            if changed > 0:
                st.success(f"   ✅ Исправлено {changed} кодов проектов (удалены пробелы в начале/конце)")
            else:
                st.info("   ℹ️ Пробелы в кодах не найдены")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")
        
        # === ШАГ 3: Заполнить пустые коды проектов ===
        st.write("**3️⃣ Заполняю пустые коды проектов...**")
        
        if code_col:
            name_col = self._find_column(df_clean, ['Имя проекта', 'Название проекта', 'Проект', 'Project Name'])
            
            if name_col:
                # Определяем пустые коды
                empty_mask = (
                    df_clean[code_col].isna() | 
                    (df_clean[code_col].astype(str).str.strip() == '') |
                    (df_clean[code_col].astype(str).str.strip() == 'nan') |
                    (df_clean[code_col].astype(str).str.strip() == 'None')
                )
                
                empty_count = empty_mask.sum()
                
                if empty_count > 0:
                    # Базовая логика: Код проекта = Имя проекта
                    df_clean.loc[empty_mask, code_col] = df_clean.loc[empty_mask, name_col]
                    st.success(f"   ✅ Заполнено {empty_count} пустых кодов (временное решение)")
                    st.info("   ⚠️ Полная логика требует объединения с массивом")
                else:
                    st.info("   ℹ️ Пустых кодов не найдено")
            else:
                st.warning("   ⚠️ Колонка с именем проекта не найдена")
        
        # === ШАГ 4: Форматировать Пилоты/Семплы/Мультикоды ===
        st.write("**4️⃣ Форматирую Пилоты/Семплы/Мультикоды...**")
        
        # 1. Найти колонку с кодом проекта
        code_col = self._find_column(df_clean, [
            'Код проекта RU00.000.00.01SVZ24',
            'Код проекта',
            'Код'
        ])
        
        if code_col:
            changes_count = 0
            
            # Значения которые ищем (в нижнем регистре)
            target_values = ['пилот', 'семпл', 'мультикод']
            
            # 2. Проверить каждое значение в колонке
            for idx, value in df_clean[code_col].items():
                if pd.isna(value):
                    continue
                    
                str_value = str(value).strip()
                
                # Приводим к нижнему регистру для сравнения
                lower_value = str_value.lower()
                
                # ШАГ 1: Найти если значение содержит target
                found_match = False
                for target in target_values:
                    if lower_value == target:  # Точное совпадение
                        found_match = True
                        break
                
                if found_match:
                    # ШАГ 2: Форматировать - первая заглавная, остальные строчные
                    formatted_value = str_value.capitalize() if str_value else str_value
                    
                    if formatted_value != str_value:
                        df_clean.at[idx, code_col] = formatted_value
                        changes_count += 1
            
            if changes_count > 0:
                st.success(f"   ✅ Отформатировано {changes_count} значений")
                st.info("   Пример: 'пиЛот' → 'Пилот', 'СЕМПЛ' → 'Семпл'")
            else:
                st.info("   ℹ️ Значения уже отформатированы")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")

        return df_clean
        
# Глобальный экземпляр
data_cleaner = DataCleaner()


