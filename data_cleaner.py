# utils/data_cleaner.py
# draft 1.1
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st
import io


class DataCleaner:

    def _find_column(self, df, possible_names):
        """Находит колонку по возможным названиям"""
        for name in possible_names:
            if name in df.columns:
                return name
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
        
        code_col = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24', 'Код', 'Project Code', 'КодПроекта'])
        
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
            name_col = self._find_column(df_clean, ['Проекты в  https://ru.checker-soft.com', 'Название проекта', 'Проект', 'Project Name'])
            
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
        
        # Используем code_col из шага 2 если есть
        if 'code_col' in locals() and code_col:
            target_col = code_col
        else:
            target_col = self._find_column(df_clean, [
                'Код проекта RU00.000.00.01SVZ24',
                'Код проекта',
                'Код'
            ])
        
        if target_col:
            changes_count = 0
            target_values = ['пилот', 'семпл', 'мультикод']
            
            for idx, value in df_clean[target_col].items():
                if pd.isna(value):
                    continue
                    
                str_value = str(value).strip()
                lower_value = str_value.lower()
                
                for target in target_values:
                    if target in lower_value:
                        formatted_value = str_value.capitalize() if str_value else str_value
                        
                        if formatted_value != str_value:
                            df_clean.at[idx, target_col] = formatted_value
                            changes_count += 1
                            break
            
            if changes_count > 0:
                st.success(f"   ✅ Отформатировано {changes_count} значений")
                st.info("   Пример: 'пиЛот' → 'Пилот', 'СЕМПЛ' → 'Семпл'")
            else:
                st.info("   ℹ️ Значения уже отформатированы")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")
        
        # === ШАГ 5: Заполнить пустые даты ===
        st.write("**5️⃣ Заполняю пустые даты...**")
        
        date_patterns = ['дата', 'date', 'срок', 'time', 'начал', 'старт', 'финиш', 'конец', 'заверш']
        date_cols = []
        
        for col in df_clean.columns:
            col_lower = str(col).lower()
            if any(pattern in col_lower for pattern in date_patterns):
                date_cols.append(col)
        
        if date_cols:
            date_fixes = 0
            
            for col in date_cols:
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    empty_dates = df_clean[col].isna().sum()
                    
                    if empty_dates > 0:
                        col_lower = str(col).lower()
                        is_start_date = any(word in col_lower for word in ['старт', 'начал', 'start'])
                        is_end_date = any(word in col_lower for word in ['финиш', 'конец', 'end', 'заверш'])
                        current_date = pd.Timestamp.now()
                        
                        for idx in df_clean[df_clean[col].isna()].index:
                            if is_start_date:
                                df_clean.at[idx, col] = current_date.replace(day=1)
                            elif is_end_date:
                                next_month = current_date.replace(day=28) + timedelta(days=4)
                                df_clean.at[idx, col] = next_month - timedelta(days=next_month.day)
                            else:
                                df_clean.at[idx, col] = current_date
                        
                        date_fixes += empty_dates
                        st.info(f"   Заполнено {empty_dates} пустых дат в '{col}'")
                except Exception as e:
                    st.warning(f"   Ошибка в колонке '{col}': {str(e)[:100]}")
            
            if date_fixes > 0:
                st.success(f"   ✅ Заполнено {date_fixes} пустых дат")
            else:
                st.info("   ℹ️ Пустых дат не найдено")
        else:
            st.warning("   ⚠️ Колонки с датами не найдены")
        
        # === ШАГ 6: Исправить даты по бизнес-правилам ===
        st.write("**6️⃣ Применяю бизнес-правила для дат...**")
        
        date_rules_applied = 0
        today = pd.Timestamp.now()
        first_day_current_month = today.replace(day=1, hour=0, minute=0, second=0)
        next_month = today.replace(day=28) + timedelta(days=4)
        last_day_current_month = next_month - timedelta(days=next_month.day)
        
        for col in date_cols:
            if col not in df_clean.columns:
                continue
                
            col_lower = str(col).lower()
            
            if any(word in col_lower for word in ['старт', 'начал', 'start']):
                try:
                    if df_clean[col].dtype != 'datetime64[ns]':
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    mask = df_clean[col] < first_day_current_month
                    
                    if mask.any():
                        df_clean.loc[mask, col] = first_day_current_month
                        date_rules_applied += mask.sum()
                        st.info(f"   Исправлено {mask.sum()} дат старта")
                except Exception as e:
                    st.warning(f"   Не удалось обработать даты старта в '{col}': {str(e)[:100]}")
            
            elif any(word in col_lower for word in ['финиш', 'конец', 'end']):
                try:
                    if df_clean[col].dtype != 'datetime64[ns]':
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    mask = df_clean[col] > last_day_current_month
                    
                    if mask.any():
                        df_clean.loc[mask, col] = last_day_current_month
                        date_rules_applied += mask.sum()
                        st.info(f"   Исправлено {mask.sum()} дат финиша")
                except Exception as e:
                    st.warning(f"   Не удалось обработать даты финиша в '{col}': {str(e)[:100]}")

        # === Проверка ошибок в годе ===
        st.info("   🔍 Проверяю ошибки в годе дат...")
        
        # Ищем колонки старта и финиша
        start_date_cols = []
        end_date_cols = []
        
        for col in date_cols:
            col_lower = str(col).lower()
            if any(word in col_lower for word in ['старт', 'начал', 'start']):
                start_date_cols.append(col)
            elif any(word in col_lower for word in ['финиш', 'конец', 'end']):
                end_date_cols.append(col)
        
        # Если нашли обе колонки
        if start_date_cols and end_date_cols:
            for start_col in start_date_cols:
                for end_col in end_date_cols:
                    try:
                        # Убедимся что обе колонки - datetime
                        if (df_clean[start_col].dtype == 'datetime64[ns]' and 
                            df_clean[end_col].dtype == 'datetime64[ns]'):
                            
                            # Находим строки где финиш раньше старта
                            mask = df_clean[end_col] < df_clean[start_col]
                            
                            if mask.any():
                                corrected_count = 0
                                
                                for idx in df_clean[mask].index:
                                    start_date = df_clean.at[idx, start_col]
                                    end_date = df_clean.at[idx, end_col]
                                    
                                    # Проверяем разницу (в днях)
                                    diff_days = (start_date - end_date).days
                                    
                                    # Если разница от 1 до 365 дней
                                    # → считаем что ошибка в годе
                                    if 1 <= diff_days <= 365:
                                        # Исправляем год финиша = год старта
                                        corrected_date = end_date.replace(year=start_date.year)
                                        df_clean.at[idx, end_col] = corrected_date
                                        corrected_count += 1
                                        st.info(f"      Строка {idx+1}: {end_date.date()} → {corrected_date.date()}")
                                
                                if corrected_count > 0:
                                    st.success(f"   ✅ Исправлено {corrected_count} ошибок в годе")
                                    date_rules_applied += corrected_count
                                    
                    except Exception as e:
                        st.warning(f"   Ошибка проверки '{start_col}' и '{end_col}': {str(e)[:50]}")
        
        # === ИТОГИ ===
        if date_rules_applied > 0:
            st.success(f"   ✅ Применено {date_rules_applied} бизнес-правил для дат")
        else:
            st.info("   ℹ️ Бизнес-правила для дат не потребовались")
        
        # === ШАГ 7: Добавить признак 'Полевой' ===
        st.write("**7️⃣ Добавляю признак 'Полевой'...**")
        
        if 'Полевой' not in df_clean.columns:
            df_clean['Полевой'] = 1
            st.success("   ✅ Добавлен признак 'Полевой' = 1 для всех записей")
        else:
            empty_field = df_clean['Полевой'].isna().sum()
            if empty_field > 0:
                df_clean['Полевой'] = df_clean['Полевой'].fillna(1)
                st.success(f"   ✅ Заполнено {empty_field} пустых значений")
            else:
                st.info("   ℹ️ Признак 'Полевой' уже заполнен")
        
        # === ИТОГИ ОЧИСТКИ ===
        st.markdown("---")
        final_rows = len(df_clean)
        final_cols = len(df_clean.columns)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Строк до очистки", original_rows, 
                     delta=f"{final_rows - original_rows}")
        
        with col2:
            st.metric("Строк после", final_rows)
        
        with col3:
            removed_pct = ((original_rows - final_rows) / original_rows * 100) if original_rows > 0 else 0
            st.metric("Удалено", f"{removed_pct:.1f}%")
        
        st.success(f"✅ Гугл таблица успешно очищена!")
        
        return df_clean

    def clean_array(self, df):
        """Очистка файла Массив  """
        if df is None or df.empty:
            st.warning("⚠️ Массив пустой или не загружен")
            return None
        
        df_clean = df.copy()
        original_rows = len(df_clean)
        original_cols = len(df_clean.columns)
        
        st.info(f"🧹 Начинаю очистку Массива: {original_rows} строк × {original_cols} колонок")
        
        # === Удалить нули в датах ===
        st.write("**1️⃣ Заменяю нули в датах на суррогатную дату (1900-01-01)...**")
        
        # Конкретные колонки с датами из Массива
        DATE_COLUMNS = [
            'Дата визита',
            'Дата создания проверки', 
            'Дата назначения опроса за тайным покупателем',
            'Дата подтверждения опроса тайным покупателем',
            'Время окончания',
            'Время завершения ожидания статуса утверждения (Дата проведения опроса?)',
            'Время утверждения'
        ]
        
        # Находим только те колонки, которые реально есть в данных
        existing_date_cols = [col for col in DATE_COLUMNS if col in df_clean.columns]
        
        if existing_date_cols:
            # Суррогатная дата для "событие еще не наступило"
            SURROGATE_DATE = pd.Timestamp('1900-01-01')
            
            total_replacements = 0
            
            for col in existing_date_cols:
                try:
                    # 🔴 УПРОЩЕННАЯ ЛОГИКА:
                    # 1. Конвертируем ВСЕ значения в datetime
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    # 2. Находим NaT (невалидные даты)
                    nat_mask = df_clean[col].isna()
                    
                    # 3. Заменяем все NaT на суррогатную дату
                    if nat_mask.any():
                        df_clean.loc[nat_mask, col] = SURROGATE_DATE
                        col_replacements = nat_mask.sum()
                        total_replacements += col_replacements
                        
                        # Показываем примеры изменений
                        example_indices = nat_mask[nat_mask].index[:3]
                        if len(example_indices) > 0:
                            st.info(f"   '{col}': заменено {col_replacements} значений")
                            for idx in example_indices:
                                if idx < len(df):  
                                    orig_val = df.at[idx, col]  
                                    st.info(f"     Строка {idx}: '{orig_val}' → '{SURROGATE_DATE.date()}'")
                                
                except Exception as e:
                    st.warning(f"   Ошибка в колонке '{col}': {str(e)[:100]}")
            
            if total_replacements > 0:
                st.success(f"   ✅ Заменено {total_replacements} невалидных дат на {SURROGATE_DATE.date()}")
                st.info("   **Обозначает:** 'Событие еще не наступило'")
            else:
                st.info("   ℹ️ Невалидных дат не найдено")
            
        else:
            st.warning(f"   ⚠️ Не найдено ни одной колонки с датами")
            st.info(f"   Искал: {', '.join(DATE_COLUMNS[:3])}...")
        
        # === Проверить массив на Н/Д ===
        st.write("**2️⃣ Заменяю Н/Д на пустые значения...**")
        
        # Значения Н/Д которые нужно заменить
        na_values = ['Н/Д', 'н/д', 'N/A', 'n/a', '#Н/Д', '#н/д', 'NA', 'na', '-', '—', '–']
        
        na_replacements = 0
        
        # Проверяем ВСЕ колонки (не только даты)
        for col in df_clean.columns:
            try:
                # Заменяем каждое значение Н/Д
                for na_val in na_values:
                    mask = df_clean[col].astype(str).str.strip() == na_val
                    if mask.any():
                        df_clean.loc[mask, col] = ''
                        na_replacements += mask.sum()
                
                # Дополнительно: заменяем текстовые 'nan', 'NaN'
                nan_mask = df_clean[col].astype(str).str.strip().str.lower().isin(['nan', 'none', 'null'])
                if nan_mask.any():
                    df_clean.loc[nan_mask, col] = ''
                    na_replacements += nan_mask.sum()
                    
            except Exception as e:
                st.warning(f"   Ошибка в колонке '{col}': {str(e)[:50]}")
        
        if na_replacements > 0:
            st.success(f"   ✅ Заменено {na_replacements} значений Н/Д")
        else:
            st.info("   ℹ️ Значений Н/Д не найдено")
        
        # === ИТОГИ ОЧИСТКИ ===
        st.markdown("---")
        st.success(f"✅ Массив успешно очищен!")

        # === Сохраняем информацию о строках с Н/Д для отчета ===
        st.write("**3️⃣ Сохраняю информацию о строках с Н/Д для отчета...**")
        
        # Создаем маску для строк, которые имели Н/Д
        had_na_mask = pd.Series(False, index=df_clean.index)
        
        for col in df_clean.columns:
            try:
                # Ищем оригинальные значения Н/Д
                for na_val in na_values:
                    mask = original_df[col].astype(str).str.strip() == na_val
                    had_na_mask = had_na_mask | mask
            except:
                continue
        
        # Сохраняем маску как атрибут DataFrame
        df_clean.attrs['had_na_rows'] = had_na_mask
        df_clean.attrs['na_rows_count'] = had_na_mask.sum()
        
        st.success(f"   ✅ Сохранено {had_na_mask.sum()} строк с Н/Д для отчета")
        
        return df_clean

    def export_array_to_excel(self, cleaned_array_df, filename="очищенный_массив"):
        """
        Создает Excel файл для очищенного массива:
        - Вкладка 1: Очищенные данные
        - Вкладка 2: Строки с Н/Д (до замены)
        - Вкладка 3: Строки с нулями в датах (до замены)
        """
        try:
            if cleaned_array_df is None or cleaned_array_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                # === ВКЛАДКА 1: Очищенные данные ===
                cleaned_array_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ МАССИВ', index=False)
                
                # === ВКЛАДКА 2: Строки где были Н/Д ===
                # Используем сохраненную информацию
                if 'had_na_rows' in cleaned_array_df.attrs:
                    had_na_mask = cleaned_array_df.attrs['had_na_rows']
                    
                    if had_na_mask.any():
                        na_rows_df = cleaned_array_df[had_na_mask].copy()
                        
                        # Добавляем информацию о пустых колонках
                        reasons = []
                        for idx in na_rows_df.index:
                            empty_cols = []
                            for col in cleaned_array_df.columns:
                                # Просто проверяем пустые значения
                                if str(cleaned_array_df.at[idx, col]).strip() == '':
                                    empty_cols.append(col)
                            
                            if empty_cols:
                                reasons.append(', '.join(empty_cols[:3]) + ('...' if len(empty_cols) > 3 else ''))
                            else:
                                reasons.append('не определено')
                        
                        na_rows_df.insert(0, 'ПУСТЫЕ_КОЛОНКИ', reasons)
                        
                        na_rows_df.insert(0, 'БЫЛИ_Н/Д_В_КОЛОНКАХ', reasons)
                        na_rows_df.to_excel(writer, sheet_name='СТРОКИ С Н Д', index=False)
                    else:
                        pd.DataFrame({'Сообщение': ['Строк с Н/Д не найдено']}).to_excel(
                            writer, sheet_name='СТРОКИ С Н Д', index=False
                        )
                else:
                    pd.DataFrame({'Сообщение': ['Информация о Н/Д не сохранена']}).to_excel(
                        writer, sheet_name='СТРОКИ С Н Д', index=False
                    )
                
                # === ВКЛАДКА 3: Строки с суррогатными датами ===
                DATE_COLUMNS = [
                    'Дата визита', 'Дата создания проверки',
                    'Дата назначения опроса за тайным покупателем',
                    'Дата подтверждения опроса тайным покупателем',
                    'Время окончания',
                    'Время завершения ожидания статуса утверждения (Дата проведения опроса?)',
                    'Время утверждения'
                ]
                
                existing_date_cols = [col for col in DATE_COLUMNS if col in cleaned_array_df.columns]
                
                if existing_date_cols:
                    SURROGATE_DATE = pd.Timestamp('1900-01-01')
                    surrogate_mask = pd.Series(False, index=cleaned_array_df.index)
                    
                    for col in existing_date_cols:
                        if cleaned_array_df[col].dtype == 'datetime64[ns]':
                            mask = cleaned_array_df[col] == SURROGATE_DATE
                            surrogate_mask = surrogate_mask | mask
                    
                    if surrogate_mask.any():
                        surrogate_rows_df = cleaned_array_df[surrogate_mask].copy()
                        
                        # Добавляем информацию о каких датах идет речь
                        date_reasons = []
                        for idx in surrogate_rows_df.index:
                            surrogate_dates = []
                            for col in existing_date_cols:
                                if (col in cleaned_array_df.columns and 
                                    cleaned_array_df[col].dtype == 'datetime64[ns]' and
                                    cleaned_array_df.at[idx, col] == SURROGATE_DATE):
                                    surrogate_dates.append(col)
                            
                            if surrogate_dates:
                                date_reasons.append(', '.join(surrogate_dates))
                            else:
                                date_reasons.append('дата не наступила')
                        
                        surrogate_rows_df.insert(0, 'НУЛИ_В_ДАТАХ', date_reasons)
                        surrogate_rows_df.to_excel(writer, sheet_name='НУЛИ В ДАТАХ', index=False)
                    else:
                        pd.DataFrame({'Сообщение': ['Строк с нулями в датах не найдено']}).to_excel(
                            writer, sheet_name='НУЛИ В ДАТАХ', index=False
                        )
                else:
                    pd.DataFrame({'Сообщение': ['Колонки с датами не найдены']}).to_excel(
                        writer, sheet_name='НУЛИ В ДАТАХ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"Ошибка при создании Excel: {e}")
            return None

        
    def enrich_array_with_project_codes(self, cleaned_array_df, projects_df):
        """
        Ищет и заполняет пустые 'Код анкеты' в очищенном Массиве,
        используя данные из таблицы Проектов Сервизория.
    
        Логика сопоставления:
        - 'Имя клиента' (Массив) -> 'Проекты в  https://ru.checker-soft.com' (Проекты)
        - 'Название проекта' (Массив) -> 'Название волны на Чекере/ином ПО' (Проекты)
    
        Возвращает:
        tuple: (enriched_array, discrepancy_df, stats_dict)
        """
        array_df = cleaned_array_df.copy()

        
        # ============================================
        # ПОДГОТОВКА ДАННЫХ
        # ============================================
        st.write("\n**4. ПОДГОТОВКА ДАННЫХ:**")
        
        # Копируем данные
        projects_df = projects_df.copy()
        
        # Находим строки с пустым 'Код анкеты'
        empty_code_mask = (
            array_df['Код анкеты'].isna() |
            (array_df['Код анкеты'].astype(str).str.strip() == '')
        )
        rows_to_process = array_df[empty_code_mask]
        total_empty = len(rows_to_process)
        
        st.write(f"- Найдено строк с пустым 'Код анкеты': {total_empty}/{len(array_df)}")
        
        if total_empty == 0:
            st.success("✅ Нечего заполнять. Все коды анкеты уже заполнены.")
            return array_df, pd.DataFrame(), {'processed': 0, 'filled': 0, 'discrepancies': 0}
        
        # ============================================
        # ОСНОВНОЙ ЦИКЛ ПОИСКА
        # ============================================
        st.write("\n**5. ПОИСК СОВПАДЕНИЙ:**")
        st.write(f"- Обрабатываю {total_empty} строк...")
        
        # Подготовка проектов для быстрого поиска
        projects_df['_match_client'] = projects_df['Проекты в  https://ru.checker-soft.com'].astype(str).str.strip()
        projects_df['_match_wave'] = projects_df['Название волны на Чекере/ином ПО'].astype(str).str.strip()
        
        # Счетчики
        filled_count = 0
        discrepancy_rows = []
        match_stats = {
            'client_match': 0,  # совпадение по клиенту
            'wave_match': 0,    # совпадение по волне
            'both_match': 0,    # совпадение по обоим полям
            'code_empty': 0,    # код проекта пустой
            'no_match': 0       # нет совпадений
        }
        
        # Примеры для отладки
        examples = []
        
        for idx, row in rows_to_process.iterrows():
            client_name = str(row['Имя клиента']).strip() if pd.notna(row['Имя клиента']) else ''
            project_name = str(row['Название проекта']).strip() if pd.notna(row['Название проекта']) else ''
            
            # Ищем точное совпадение
            match_mask = (
                (projects_df['_match_client'] == client_name) &
                (projects_df['_match_wave'] == project_name)
            )
            
            matched_rows = projects_df[match_mask]
            
            if not matched_rows.empty:
                match_stats['both_match'] += 1
                project_code = matched_rows.iloc[0]['Код проекта RU00.000.00.01SVZ24']
                
                if pd.notna(project_code) and str(project_code).strip() != '':
                    # Заполняем код
                    array_df.at[idx, 'Код анкеты'] = str(project_code).strip()
                    filled_count += 1
                    
                    # Сохраняем пример для отладки (первые 3)
                    if len(examples) < 3:
                        examples.append({
                            'клиент': client_name[:30] + '...' if len(client_name) > 30 else client_name,
                            'проект': project_name[:30] + '...' if len(project_name) > 30 else project_name,
                            'найденный код': str(project_code).strip()[:20] + '...' if len(str(project_code)) > 20 else str(project_code)
                        })
                else:
                    match_stats['code_empty'] += 1
                    discrepancy_rows.append(row.to_dict())
            else:
                match_stats['no_match'] += 1
                discrepancy_rows.append(row.to_dict())
        
        # ============================================
        # РЕЗУЛЬТАТЫ
        # ============================================
        st.write("\n**6. РЕЗУЛЬТАТЫ ПОИСКА:**")
        st.write(f"- Совпадений по обоим полям (клиент+волна): {match_stats['both_match']}/{total_empty}")
        st.write(f"- Из них с заполненным кодом проекта: {filled_count}/{match_stats['both_match']}")
        st.write(f"- Из них с пустым кодом проекта: {match_stats['code_empty']}/{match_stats['both_match']}")
        st.write(f"- Без совпадений: {match_stats['no_match']}/{total_empty}")
        
        if examples:
            st.write("\n**Примеры найденных совпадений:**")
            for i, example in enumerate(examples, 1):
                st.write(f"  {i}. Клиент: '{example['клиент']}'")
                st.write(f"     Проект: '{example['проект']}'")
                st.write(f"     Код: '{example['найденный код']}'")
        
        # Формируем результат
        discrepancy_df = pd.DataFrame(discrepancy_rows) if discrepancy_rows else pd.DataFrame()
        
        st.write("\n**7. ИТОГИ:**")
        st.write(f"- Всего обработано: {total_empty} строк")
        st.write(f"- Успешно заполнено: {filled_count} кодов")
        st.write(f"- Осталось расхождений: {len(discrepancy_df)} строк")
        
        # Удаляем временные колонки
        projects_df.drop(['_match_client', '_match_wave'], axis=1, inplace=True, errors='ignore')
        
        st.success(f"✅ Обогащение завершено!")
        st.write("=" * 50)
        
        stats = {
            'processed': total_empty,
            'filled': filled_count,
            'discrepancies': len(discrepancy_df),
            'match_stats': match_stats
        }
        
        return array_df, discrepancy_df, stats


    def export_discrepancies_to_excel(self, discrepancy_df, filename="Расхождение_Массив"):
        """Создает Excel файл для расхождений"""
        try:
            if discrepancy_df is None or discrepancy_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Добавляем пояснительную вкладку
                info_df = pd.DataFrame({
                    'Информация': [
                        'Файл создан автоматически',
                        f'Дата создания: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}',
                        f'Количество строк: {len(discrepancy_df)}',
                        'Эти строки не удалось обогатить кодами проектов'
                    ]
                })
                info_df.to_excel(writer, sheet_name='ИНФО', index=False)
                
                # Основные данные
                discrepancy_df.to_excel(writer, sheet_name='РАСХОЖДЕНИЯ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"Ошибка при создании Excel с расхождениями: {e}")
            return None

    
    
    def export_to_excel(self, original_df, cleaned_df, filename="очищенные_данные"):
        """
        Создает Excel файл с тремя вкладками для сравнения
        """
        try:
            if original_df is None or cleaned_df is None:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Вкладка 1: Оригинальные данные
                original_df.to_excel(writer, sheet_name='ОРИГИНАЛ', index=False)
                
                # Вкладка 2: Очищенные данные
                cleaned_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ', index=False)
                
                # Вкладка 3: Сравнение изменений
                comparison_data = []
                
                comparison_data.append({
                    'Параметр': 'Количество строк',
                    'Оригинал': len(original_df),
                    'Очищено': len(cleaned_df),
                    'Изменение': len(cleaned_df) - len(original_df)
                })
                
                comparison_data.append({
                    'Параметр': 'Количество колонок',
                    'Оригинал': len(original_df.columns),
                    'Очищено': len(cleaned_df.columns),
                    'Изменение': len(cleaned_df.columns) - len(original_df.columns)
                })
                
                added_cols = set(cleaned_df.columns) - set(original_df.columns)
                removed_cols = set(original_df.columns) - set(cleaned_df.columns)
                
                if added_cols:
                    comparison_data.append({
                        'Параметр': 'Добавленные колонки',
                        'Оригинал': '-',
                        'Очищено': ', '.join(added_cols),
                        'Изменение': f'+{len(added_cols)}'
                    })
                
                if removed_cols:
                    comparison_data.append({
                        'Параметр': 'Удаленные колонки',
                        'Оригинал': ', '.join(removed_cols),
                        'Очищено': '-',
                        'Изменение': f'-{len(removed_cols)}'
                    })
                
                comparison_df = pd.DataFrame(comparison_data)
                comparison_df.to_excel(writer, sheet_name='СРАВНЕНИЕ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"Ошибка при создании Excel: {e}")
            return None


# Глобальный экземпляр
data_cleaner = DataCleaner()






















