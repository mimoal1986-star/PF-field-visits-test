# utils/data_cleaner.py
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st
import io
from io import BytesIO


class DataCleaner:
    """
    Очистка данных по инструкции для ИУ Аудиты
    """
    
    def _find_column(self, df, possible_names):
        """
        Найти колонку по возможным названиям
        """
        if df is None or df.empty:
            return None
            
        # Приводим все названия колонок к нижнему регистру для сравнения
        actual_columns = [str(col).lower().strip() for col in df.columns]
        
        for possible in possible_names:
            possible_lower = str(possible).lower().strip()
            # Ищем точное совпадение
            if possible_lower in actual_columns:
                # Возвращаем оригинальное название
                for col in df.columns:
                    if str(col).lower().strip() == possible_lower:
                        return col
            # Ищем частичное совпадение
            for col in df.columns:
                if possible_lower in str(col).lower().strip():
                    return col
                    
        return None
    
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
        
        # 1. Найти колонку с кодом проекта (используем новое имя переменной)
        code_col_step4 = self._find_column(df_clean, [
            'Код проекта RU00.000.00.01SVZ24',
            'Код проекта',
            'Код'
        ])
        
        if code_col_step4:
            changes_count = 0
            examples = []
            
            # Список целевых слов и их правильное написание
            target_words = {
                'пилот': 'Пилот',
                'семпл': 'Семпл', 
                'мультикод': 'Мультикод'
            }
            
            # 2. Пройти по всем строкам
            for i in range(len(df_clean)):
                try:
                    original_value = df_clean.at[i, code_col_step4]
                    
                    # Пропускаем пустые значения
                    if pd.isna(original_value):
                        continue
                        
                    # Приводим к строке
                    str_value = str(original_value).strip()
                    if not str_value:
                        continue
                        
                    # Приводим к нижнему регистру для поиска
                    lower_value = str_value.lower()
                    
                    # Проверяем, содержит ли строка любое из целевых слов
                    word_found = None
                    for target_lower, target_proper in target_words.items():
                        # Ищем полное совпадение (значение == целевому слову)
                        if lower_value == target_lower:
                            word_found = target_proper
                            break
                        # Или начинается с целевого слова (например "пилот_123")
                        elif lower_value.startswith(target_lower + '_') or lower_value.startswith(target_lower + '-'):
                            word_found = target_proper
                            break
                    
                    # Если нашли целевое слово
                    if word_found:
                        # Заменяем в строке
                        if str_value.lower() == word_found.lower():  # Полное совпадение
                            formatted_value = word_found
                        elif str_value.lower().startswith(word_found.lower() + '_'):
                            # Сохраняем суффикс (например "_123")
                            suffix = str_value[len(word_found):]
                            formatted_value = word_found + suffix
                        elif str_value.lower().startswith(word_found.lower() + '-'):
                            # Сохраняем суффикс (например "-250")
                            suffix = str_value[len(word_found):]
                            formatted_value = word_found + suffix
                        else:
                            # Если не распознан формат, просто капитализируем
                            formatted_value = str_value.title()
                        
                        if formatted_value != str_value:
                            df_clean.at[i, code_col_step4] = formatted_value
                            changes_count += 1
                            
                            # Сохраняем примеры (максимум 3)
                            if len(examples) < 3:
                                examples.append(f"'{str_value}' → '{formatted_value}'")
                except Exception as e:
                    # Пропускаем ошибки в отдельных строках
                    continue
            
            if changes_count > 0:
                st.success(f"   ✅ Отформатировано {changes_count} значений")
                if examples:
                    st.info(f"   Примеры: {', '.join(examples)}")
                st.info("   📝 Изменения: первая заглавная, остальные строчные")
            else:
                st.info("   ℹ️ Целевые слова не найдены или уже отформатированы")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")


# Глобальный экземпляр
data_cleaner = DataCleaner()
