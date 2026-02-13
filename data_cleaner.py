# utils/data_cleaner.py
# draft 2.0 
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
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first') # удаляем дубликаты записей
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
                
        else:
            st.warning("   ⚠️ Не найдено ни одного ключевого поля для проверки дубликатов")
            st.info("   Проверьте названия колонок в файле")
            
        
        # === ШАГ 2: Сжать пробелы в кодах проектов ===
        st.write("**2️⃣ Чищу пробелы в кодах проектов...**")
        
        code_col = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24'])
        
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
        st.write("**3️⃣ Проверяю пустые коды проектов...**")
        
        if code_col:
            empty_mask = (
                df_clean[code_col].isna() | 
                (df_clean[code_col].astype(str).str.strip() == '') |
                (df_clean[code_col].astype(str).str.strip() == 'nan') |
                (df_clean[code_col].astype(str).str.strip() == 'None')
            )
            
            empty_count = empty_mask.sum()
            if empty_count > 0:
                st.warning(f"   ⚠️ Найдено {empty_count} проектов без кода")
            else:
                st.info("   ℹ️ Пустых кодов не найдено")
                
        
        # === ШАГ 4: Форматировать Пилоты/Семплы/Мультикоды ===
        st.write("**4️⃣ Форматирую Пилоты/Семплы/Мультикоды...**")
        
        # Используем код проекта из шага 2 если есть
        if 'code_col' in locals() and code_col:
            target_col = code_col
        else:
            target_col = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24'])
        
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
        
        # === ШАГ 5: Оставляем пустые даты как есть (ТОЛЬКО конвертация) ===
        st.write("**5️⃣ Конвертирую даты...**")

        # ТОЛЬКО эти 2 колонки обрабатываем как даты
        date_cols_to_process = []
        for col_name in ['Дата старта', 'Дата финиша с продлением']:
            if col_name in df_clean.columns:
                date_cols_to_process.append(col_name)
            else:
                # Ищем возможные варианты названий
                found = self._find_column(df_clean, [col_name])
                if found:
                    date_cols_to_process.append(found)
        
        date_cols = date_cols_to_process

        
        if date_cols:
            date_fixes = 0
            
            for col in date_cols:
                try:
                    # ТОЛЬКО конвертация, БЕЗ заполнения пустых!
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    empty_dates = df_clean[col].isna().sum()
                    
                    if empty_dates > 0:
                        # ВАЖНО: НЕ заполняем пустые даты!
                        st.warning(f"   ⚠️ В колонке '{col}': {empty_dates} ПУСТЫХ дат")
                        st.warning(f"      Проекты с пустыми датами БУДУТ ПРОПУЩЕНЫ в расчетах")
                except Exception as e:
                    st.warning(f"   Ошибка в колонке '{col}': {str(e)[:100]}")
            
            st.success(f"   ✅ Даты сконвертированы")
        else:
            st.warning("   ⚠️ Колонки с датами не найдены")
        

        # === ШАГ 6: Исправить даты по бизнес-правилам ===

        st.write("**6️⃣ Применяю бизнес-правила для дат...**")
        
        date_rules_applied = 0
        
        # 1. НАХОДИМ МАКСИМАЛЬНЫЙ МЕСЯЦ И НАИБОЛЕЕ ЧАСТЫЙ ГОД В ДАННЫХ
        st.info("   🔍 Анализирую даты в данных...")
        
        # Собираем все даты и года
        all_dates = []
        all_years = []
        
        for col in date_cols:
            if col in df_clean.columns:
                try:
                    if df_clean[col].dtype != 'datetime64[ns]':
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    # Добавляем валидные даты
                    valid_dates = df_clean[col].dropna()
                    if not valid_dates.empty:
                        all_dates.extend(valid_dates.tolist())
                        all_years.extend(valid_dates.dt.year.tolist())
                except:
                    continue
        
        if all_dates:
            # Находим максимальную дату (для бизнес-правил)
            max_date = max(all_dates)
            max_year = max_date.year
            max_month = max_date.month
            
            st.success(f"   ✅ Максимальная дата в данных: {max_date.strftime('%d.%m.%Y')}")
            
            # Находим наиболее частый год (для исправления ошибок)
            from collections import Counter
            if all_years:
                year_counts = Counter(all_years)
                target_year, target_count = year_counts.most_common(1)[0]
                st.success(f"   🎯 Наиболее частый год: {target_year} ({target_count} дат)")
            else:
                target_year = max_year
                st.info(f"   🎯 Использую максимальный год как целевой: {target_year}")
            
            # 2. ВЫЧИСЛЯЕМ ГРАНИЦЫ МАКСИМАЛЬНОГО МЕСЯЦА
            # Первый день максимального месяца
            first_day = pd.Timestamp(year=max_year, month=max_month, day=1)
            
            # Последний день максимального месяца
            if max_month == 12:
                next_month = pd.Timestamp(year=max_year+1, month=1, day=1)
            else:
                next_month = pd.Timestamp(year=max_year, month=max_month+1, day=1)
            
            last_day = next_month - pd.Timedelta(days=1)
            
            st.info(f"   📅 Период для бизнес-правил: {first_day.strftime('%d.%m.%Y')} - {last_day.strftime('%d.%m.%Y')}")
            
            # 3. ИСПРАВЛЯЕМ ОШИБКИ В ГОДЕ НА ОСНОВЕ СТАТИСТИКИ
            year_errors_corrected = 0
            
            # Ищем колонки старта и финиша
            start_date_cols = []
            end_date_cols = []
            
            for col in date_cols:
                col_lower = str(col).lower()
                if any(word in col_lower for word in ['старт', 'начал', 'start']):
                    start_date_cols.append(col)
                elif any(word in col_lower for word in ['финиш', 'конец', 'end']):
                    end_date_cols.append(col)
            
            # Исправляем ошибки если есть обе колонки
            if start_date_cols and end_date_cols:
                for start_col in start_date_cols:
                    for end_col in end_date_cols:
                        try:
                            # Находим строки где год сильно отличается от целевого
                            mask_year_error = (
                                (df_clean[start_col].dt.year != target_year) | 
                                (df_clean[end_col].dt.year != target_year)
                            )
                            
                            if mask_year_error.any():
                                for idx in df_clean[mask_year_error].index:
                                    start_date = df_clean.at[idx, start_col]
                                    end_date = df_clean.at[idx, end_col]
                                    
                                    # Исправляем год старта если сильно отличается
                                    if abs(start_date.year - target_year) > 1:
                                        df_clean.at[idx, start_col] = start_date.replace(year=target_year)
                                        year_errors_corrected += 1
                                    
                                    # Исправляем год финиша если сильно отличается
                                    if abs(end_date.year - target_year) > 1:
                                        df_clean.at[idx, end_col] = end_date.replace(year=target_year)
                                        year_errors_corrected += 1
                            
                            # Также исправляем ситуации "финиш раньше старта" в пределах года
                            mask_finish_before_start = df_clean[end_col] < df_clean[start_col]
                            
                            if mask_finish_before_start.any():
                                for idx in df_clean[mask_finish_before_start].index:
                                    start_date = df_clean.at[idx, start_col]
                                    end_date = df_clean.at[idx, end_col]
                                    
                                    # Если разница небольшая (< 30 дней) - возможно ошибка в месяце
                                    if (start_date - end_date).days <= 30:
                                        # Делаем финиш тем же числом, но месяцем старта
                                        corrected_date = end_date.replace(year=start_date.year, month=start_date.month)
                                        df_clean.at[idx, end_col] = corrected_date
                                        year_errors_corrected += 1
                                        
                        except Exception as e:
                            st.warning(f"   Ошибка при проверке '{start_col}' и '{end_col}': {str(e)[:50]}")
            
            if year_errors_corrected > 0:
                st.success(f"   ✅ Исправлено {year_errors_corrected} ошибок в годе")
            
            # 4. ПРИМЕНЯЕМ БИЗНЕС-ПРАВИЛА
            
            # ПРАВИЛО 1: Если дата СТАРТА раньше первого дня максимального месяца → ставим первый день
            for col in date_cols:
                col_lower = str(col).lower()
                
                if any(word in col_lower for word in ['старт', 'начал', 'start']):
                    try:
                        mask = df_clean[col] < first_day
                        if mask.any():
                            df_clean.loc[mask, col] = first_day
                            date_rules_applied += mask.sum()
                            st.info(f"   ⚙️ Исправлено {mask.sum()} дат старта (были раньше {first_day.strftime('%d.%m.%Y')})")
                    except Exception as e:
                        st.warning(f"   Ошибка с датами старта: {str(e)[:100]}")
            
            # ПРАВИЛО 2: Если дата ФИНИША раньше первого дня ИЛИ позже последнего дня → ставим последний день
            for col in date_cols:
                col_lower = str(col).lower()
                
                if any(word in col_lower for word in ['финиш', 'конец', 'end']):
                    try:
                        # ПРАВИЛО 2.1: Если финиш РАНЬШЕ первого дня → last_day
                        mask_too_early = df_clean[col] < first_day
                        if mask_too_early.any():
                            df_clean.loc[mask_too_early, col] = last_day
                            date_rules_applied += mask_too_early.sum()
                            st.info(f"   ⚙️ Исправлено {mask_too_early.sum()} дат финиша (были раньше {first_day.strftime('%d.%m.%Y')})")
                        
                        # ПРАВИЛО 2.2: Если финиш ПОЗЖЕ последнего дня → last_day
                        mask_too_late = df_clean[col] > last_day
                        if mask_too_late.any():
                            df_clean.loc[mask_too_late, col] = last_day
                            date_rules_applied += mask_too_late.sum()
                            st.info(f"   ⚙️ Исправлено {mask_too_late.sum()} дат финиша (были позже {last_day.strftime('%d.%m.%Y')})")
                            
                    except Exception as e:
                        st.warning(f"   Ошибка с датами финиша: {str(e)[:100]}")
        
        else:
            st.warning("   ⚠️ Не найдено валидных дат для анализа")
        
        # 5. ИТОГИ
        total_corrections = date_rules_applied + year_errors_corrected
        if total_corrections > 0:
            st.success(f"   ✅ Применено {total_corrections} исправлений ({year_errors_corrected} ошибок в годе + {date_rules_applied} бизнес-правил)")
        else:
            st.info("   ℹ️ Исправления дат не потребовались")
        
        # === ИТОГИ ===
        if date_rules_applied > 0:
            st.success(f"   ✅ Применено {date_rules_applied} бизнес-правил для дат")
        else:
            st.info("   ℹ️ Бизнес-правила для дат не потребовались")
        
        # # === ШАГ 7: Добавить признак 'Полевой' ===
        # st.write("**7️⃣ Добавляю признак 'Полевой'...**")
        
        # if 'Полевой' not in df_clean.columns:
        #     df_clean['Полевой'] = 0
        #     st.success("   ✅ Добавлен признак 'Полевой' = 0 для всех записей")
        # else:
        #     empty_field = df_clean['Полевой'].isna().sum()
        #     if empty_field > 0:
        #         df_clean['Полевой'] = df_clean['Полевой'].fillna(1)
        #         st.success(f"   ✅ Заполнено {empty_field} пустых значений")
        #     else:
        #         st.info("   ℹ️ Признак 'Полевой' уже заполнен")
        
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
            
        # === ШАГ 8: Добавить колонку ЗОД ===
        st.write("**3️⃣ Добавляю колонку ЗОД (пока пустую)...**")
        
        if 'ЗОД' not in df_clean.columns:
            df_clean['ЗОД'] = ''
            st.success("   ✅ Добавлена колонка 'ЗОД' (будет заполнена позже из справочника)")
        else:
            st.info("   ℹ️ Колонка 'ЗОД' уже существует")
        
        # === ИТОГИ ОЧИСТКИ ===
        st.markdown("---")
        st.success(f"✅ Массив успешно очищен!")

        # === Сохраняем информацию о строках с Н/Д для отчета ===
        st.write("**4️⃣ Сохраняю информацию о строках с Н/Д для отчета...**")
        
        # Создаем маску для строк, которые имели Н/Д
        had_na_mask = pd.Series(False, index=df_clean.index)
        
        for col in df_clean.columns:
            try:
                # Ищем оригинальные значения Н/Д
                for na_val in na_values:
                    mask = df[col].astype(str).str.strip() == na_val
                    had_na_mask = had_na_mask | mask
            except:
                continue
        
        # Сохраняем маску как атрибут DataFrame
        df_clean.attrs['had_na_rows'] = had_na_mask
        df_clean.attrs['na_rows_count'] = had_na_mask.sum()
        
        st.success(f"   ✅ Сохранено {had_na_mask.sum()} строк с Н/Д для отчета")
        
        return df_clean
    
    def add_zod_from_hierarchy(self, array_df, hierarchy_df):
        """
        Добавляет колонку ЗОД в массив на основе справочника ЗОД+АСС
        Логика: АСС (массив) -> ЗОД (справочник)
        """
        try:
            if array_df is None or array_df.empty:
                return array_df
                
            if hierarchy_df is None or hierarchy_df.empty:
                st.warning("⚠️ Справочник ЗОД+АСС пустой")
                return array_df
            
            array_clean = array_df.copy()
            hierarchy_clean = hierarchy_df.copy()
            
            # Находим колонки в справочнике
            zodiac_col = self._find_column(hierarchy_clean, ['ЗОД', 'zod', 'ZOD'])
            acc_col = self._find_column(hierarchy_clean, ['АСС', 'acc', 'ACC'])
            
            if not zodiac_col or not acc_col:
                st.error("❌ В справочнике не найдены колонки ЗОД и/или АСС")
                return array_df
            
            # Находим колонку АСС в массиве
            array_acc_col = self._find_column(array_clean, ['АСС', 'acc', 'ACC'])
            
            if not array_acc_col:
                st.error("❌ В массиве не найдена колонка АСС")
                return array_df
            
            # Создаем словарь сопоставления {АСС: ЗОД}
            zod_mapping = {}
            for _, row in hierarchy_clean.iterrows():
                acc_val = str(row[acc_col]).strip()
                zod_val = str(row[zodiac_col]).strip()
                
                if acc_val and acc_val.lower() not in ['nan', 'none', 'null', '']:
                    zod_mapping[acc_val] = zod_val
            
            st.info(f"🔍 Загружено {len(zod_mapping)} сопоставлений АСС → ЗОД")
            
            # Добавляем или обновляем колонку ЗОД
            if 'ЗОД' in array_clean.columns:
                array_clean['ЗОД'] = ''
            else:
                array_clean['ЗОД'] = ''
            
            # Заполняем ЗОД на основе АСС
            def get_zod_by_acc(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip().lower() in ['nan', 'none', 'null', '']:
                    return ''
                clean_acc = str(acc_value).strip()
                return zod_mapping.get(clean_acc, '')
            
            array_clean['ЗОД'] = array_clean[array_acc_col].apply(get_zod_by_acc)
            
            filled_count = (array_clean['ЗОД'] != '').sum()
            st.success(f"✅ Добавлен ЗОД: заполнено {filled_count} значений")
            
            return array_clean
            
        except Exception as e:
            st.error(f"❌ Ошибка в add_zod_from_hierarchy: {str(e)[:100]}")
            return array_df

    def export_array_to_excel(self, cleaned_array_df, filename="очищенный_массив"):
        """
        Создает Excel файл для очищенного массива:
        - Вкладка 1: Очищенные данные
        """
        try:
            if cleaned_array_df is None or cleaned_array_df.empty:
                return None
            
            output = io.BytesIO()
            
            # Используем контекстный менеджер pd.ExcelWriter
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # === ВКЛАДКА 1: Очищенные данные ===
                cleaned_array_df.to_excel(
                    writer, 
                    sheet_name='ОЧИЩЕННЫЙ МАССИВ', 
                    index=False
                )
            
            # Важно: перемещаем указатель в начало
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

    
    
    def export_to_excel(self, df, cleaned_df, filename="очищенные_данные"):
        try:
            if df is None or cleaned_df is None:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Вкладка 1: Оригинальные данные
                df.to_excel(writer, sheet_name='ОРИГИНАЛ', index=False)
                
                # Вкладка 2: Очищенные данные
                cleaned_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"Ошибка при создании Excel: {e}")
            return None
    
    # ============================================
    # Подготовка Массива для финального датасета
    # ============================================
    
    def update_field_projects_flag(self, google_df):
        """
        Обновляет поле 'Полевой' в гугл таблице
        Полевой = 1 если:
        1. Страна = RU00 И Направление = .01/.02
        2. Или код содержит: мультикод/пилот/семпл (любой регистр)
        """
        try:
            google_df = google_df.copy()
            
            # Колонка с кодом проекта
            code_col = 'Код проекта RU00.000.00.01SVZ24'
            
            if code_col not in google_df.columns:
                return google_df
            
            def is_field_project(code):
                try:
                    if pd.isna(code):
                        return 0
                        
                    code_str = str(code).strip()
                    lower_code = code_str.lower()
                    
                    # Проверка на мультикод/пилот/семпл
                    if any(word in lower_code for word in ['мультикод', 'пилот', 'семпл']):
                        return 1
                    
                    # Проверка формата RU00.001.06.01SVZ24
                    parts = code_str.split('.')
                    if len(parts) >= 4:
                        country = parts[0]  # RU00
                        # Получаем направление как .01 или .02
                        if len(parts[2]) >= 2:
                            direction = '.' + parts[2][:2]  # .01 или .02
                        else:
                            direction = ''
                        
                        if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
                            return 1
                            
                    return 0
                except:
                    return 0
                            
            
            google_df['Полевой'] = google_df[code_col].apply(is_field_project)
            
            field_count = google_df['Полевой'].sum()
            st.info(f"🔍 Определено {field_count} полевых проектов")
            
            return google_df
            
        except Exception as e:
            st.error(f"Ошибка в update_field_projects_flag: {e}")
            return google_df


    def add_field_flag_to_array(self, array_df):
        """
        Добавляет 'Полевой' в массив
        Полевой = 1 если:
        1. Страна = RU00 И Направление = .01/.02
        2. Или код содержит: мультикод/пилот/семпл (любой регистр)
        """
        try:
            array_df = array_df.copy()
            
            # Ищем колонку с кодом анкеты
            code_col = None
            for col in array_df.columns:
                if 'код' in str(col).lower() and 'анкет' in str(col).lower():
                    code_col = col
                    break
            
            if not code_col:
                st.warning("⚠️ Не найдена колонка 'Код анкеты'")
                return array_df
            
            def is_field_project(code):
                try:
                    if pd.isna(code):
                        return 0
                        
                    code_str = str(code).strip()
                    lower_code = code_str.lower()
                    
                    # Проверка на мультикод/пилот/семпл
                    if any(word in lower_code for word in ['мультикод', 'пилот', 'семпл']):
                        return 1
                    
                    # Проверка формата RU00.001.06.01SVZ24
                    parts = code_str.split('.')
                    if len(parts) >= 4:
                        country = parts[0]  # RU00
                        if len(parts[2]) >= 2:
                            direction = '.' + parts[2][:2]  # '.' + '02' = '.02'
                        else:
                            direction = ''
                        
                        if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
                            return 1
                            
                    return 0
                except:
                    return 0
            
            array_df['Полевой'] = array_df[code_col].apply(is_field_project)
            
            field_count = array_df['Полевой'].sum()
            st.success(f"✅ Добавлен 'Полевой': {field_count} полевых записей")
            
            return array_df
            
        except Exception as e:
            st.error(f"Ошибка в add_field_flag_to_array: {e}")
            return array_df
    
    def split_array_by_field_flag(self, array_df):
        """
        Разделяет массив на Полевые и Неполевые проекты
        Сохраняет колонку 'Полевой' в обеих частях
        Возвращает 9 колонок (8 основных + Полевой)
        """
        try:
            if array_df is None or array_df.empty:
                st.warning("⚠️ Массив пустой")
                return None, None
            
            array_df_clean = array_df.copy()
            
            if 'Полевой' not in array_df_clean.columns:
                st.error("❌ В массиве нет колонки 'Полевой'")
                return None, None
            
            # Маппинг стандартных колонок (9 колонок)
            column_mapping = {
                'Код анкеты': ['Код анкеты'],
                'Имя клиента': ['Имя клиента'],
                'Название проекта': ['Название проекта'],
                'ЗОД': ['ЗОД', 'ZOD', 'Зод', 'zod'],
                'АСС': ['АСС', 'ASS', 'Асс', 'ass'],
                'ЭМ': ['ЭМ рег'],
                'Регион short': ['Регион'],
                'Регион': ['Регион '],
                'Полевой': ['Полевой'],
                ' Статус': [' Статус', 'Статус'],
                'Дата визита': ['Дата визита'],
                'ПО': ['ПО']
            }
            
            # Находим фактические названия колонок
            actual_columns = {}
            
            for std_col, possible_names in column_mapping.items():
                found_col = self._find_column(array_df_clean, possible_names)
                if found_col:
                    actual_columns[std_col] = found_col
                else:
                    # Создаем пустую колонку (кроме Полевой)
                    if std_col != 'Полевой':
                        array_df_clean[std_col] = ''
                        actual_columns[std_col] = std_col
            
            # Отбираем нужные колонки
            selected_cols = list(actual_columns.values())
            
            # Фильтруем данные
            field_mask = array_df_clean['Полевой'] == 1
            field_projects = array_df_clean.loc[field_mask, selected_cols].copy()
            non_field_projects = array_df_clean.loc[~field_mask, selected_cols].copy()
            
            # Переименовываем колонки
            reverse_mapping = {v: k for k, v in actual_columns.items()}
            
            if not field_projects.empty:
                field_projects = field_projects.rename(columns=reverse_mapping)
            
            if not non_field_projects.empty:
                non_field_projects = non_field_projects.rename(columns=reverse_mapping)
            
            # Правильный порядок колонок
            final_columns = ['Код анкеты', 'Имя клиента', 'Название проекта', 
                           'ЗОД', 'АСС', 'ЭМ', 'Регион short', 'Регион', 'ПО', 'Полевой']
            
            # Реорганизуем колонки
            if not field_projects.empty:
                for col in final_columns:
                    if col not in field_projects.columns:
                        field_projects[col] = '' if col != 'Полевой' else 0
                field_projects = field_projects.reindex(columns=final_columns)
            
            if not non_field_projects.empty:
                for col in final_columns:
                    if col not in non_field_projects.columns:
                        non_field_projects[col] = '' if col != 'Полевой' else 0
                non_field_projects = non_field_projects.reindex(columns=final_columns)
            
            st.success(f"✅ Разделение завершено:")
            st.info(f"   📊 Полевые: {len(field_projects)} записей")
            st.info(f"   📊 Неполевые: {len(non_field_projects)} записей")
            
            return field_projects, non_field_projects
            
        except Exception as e:
            st.error(f"❌ Ошибка в split_array_by_field_flag: {str(e)[:100]}")
            return None, None
    
    
    def export_split_array_to_excel(self, field_df, non_field_df, filename="разделенный_массив"):
        """ Создает Excel с вкладками Полевые/Неполевые проекты """
        try:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Полевые проекты
                if field_df is not None and not field_df.empty:
                    field_df.to_excel(writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет полевых проектов']}).to_excel(
                        writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
                
                # Неполевые проекты
                if non_field_df is not None and not non_field_df.empty:
                    non_field_df.to_excel(writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет неполевых проектов']}).to_excel(
                        writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
                
                # # Статистика
                # stats_data = {
                #     'Метрика': ['Всего записей', 'Полевые', 'Неполевые', 'Дата обработки'],
                #     'Значение': [
                #         (len(field_df) if field_df is not None else 0) + 
                #         (len(non_field_df) if non_field_df is not None else 0),
                #         len(field_df) if field_df is not None else 0,
                #         len(non_field_df) if non_field_df is not None else 0,
                #         pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                #     ]
                # }
                # pd.DataFrame(stats_data).to_excel(writer, sheet_name='СТАТИСТИКА', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"❌ Ошибка при создании Excel: {str(e)[:100]}")
            return None
            
    def export_field_projects_only(self, field_df, filename="полевые_проекты"):
        """
        Создает Excel файл ТОЛЬКО с полевыми проектами
        """
        try:
            if field_df is None or field_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Полевые проекты
                if not field_df.empty:
                    field_df.to_excel(writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет полевых проектов']}).to_excel(
                        writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"❌ Ошибка при создании Excel (полевые): {str(e)[:100]}")
            return None

    def export_non_field_projects_only(self, non_field_df, filename="неполевые_проекты"):
        """
        Создает Excel файл ТОЛЬКО с неполевыми проектами
        УДАЛЯЕТ ДУБЛИКАТЫ по: Код проекта, Имя клиента, Название проекта
        """
        try:
            if non_field_df is None or non_field_df.empty:
                return None
            
            # Удаляем дубликаты по 3 полям
            non_field_clean = non_field_df.copy()
            
            # Проверяем наличие нужных колонок
            required_cols = ['Код анкеты', 'Имя клиента', 'Название проекта']
            missing_cols = [col for col in required_cols if col not in non_field_clean.columns]
            before_rows = len(non_field_clean)
            after_rows = before_rows
            duplicates_removed = 0
            
            if missing_cols:
                st.warning(f"⚠️ В неполевых проектах нет колонок: {missing_cols}")
                # Если нет нужных колонок - не удаляем дубликаты
                non_field_unique = non_field_clean
                after_rows = before_rows
                duplicates_removed = 0
            else:
                # Удаляем дубликаты
                before_rows = len(non_field_clean)
                non_field_unique = non_field_clean.drop_duplicates(
                    subset=required_cols, 
                    keep='first'
                )
                after_rows = len(non_field_unique)
                duplicates_removed = before_rows - after_rows
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Неполевые проекты (без дубликатов)
                if not non_field_unique.empty:
                    # УДАЛЯЕМ ЛИШНИЕ КОЛОНКИ
                    columns_to_drop = ['ЗОД', 'АСС', 'ЭМ', 'Регион short', 'Регион', 'Полевой']
                    columns_exist = [col for col in columns_to_drop if col in non_field_unique.columns]
                    
                    if columns_exist:
                        non_field_export = non_field_unique.drop(columns=columns_exist, errors='ignore')
                    else:
                        non_field_export = non_field_unique
                    
                    non_field_export.to_excel(writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False)

                    
                    # Добавляем информацию об удаленных дубликатах на отдельную вкладку
                    info_data = {
                        'Информация': [
                            f'Файл создан: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}',
                            f'Всего записей до удаления дубликатов: {before_rows}',
                            f'Всего записей после удаления: {after_rows}',
                            f'Удалено дубликатов: {duplicates_removed}',
                            f'Поля для удаления дубликатов: Код проекта, Имя клиента, Название проекта'
                        ]
                    }
                    pd.DataFrame(info_data).to_excel(writer, sheet_name='ИНФОРМАЦИЯ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет неполевых проектов']}).to_excel(
                        writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"❌ Ошибка при создании Excel (неполевые): {str(e)[:100]}")
            return None

    def check_problematic_projects(self, google_df, autocoding_df, array_df):
        """Проверка проблемных проектов - оптимизированная версия"""
        result = google_df.copy()
        
        # Базовые проверки
        code_col = 'Код проекта RU00.000.00.01SVZ24'
        project_col = 'Проекты в  https://ru.checker-soft.com'
        wave_col = 'Название волны на Чекере/ином ПО'
        portal_col = 'Портал на котором идет проект (для работы полевой команды)'

        # 1. Код проекта пусто
        result['Код проекта пусто'] = (
            result[code_col].isna() |                     # стандартный pandas NaN
            (result[code_col].astype(str).str.strip() == 'nan') |   # текстовый 'nan'
            (result[code_col].astype(str).str.strip() == 'NaN') |   # текстовый 'NaN' (заглавные)
            (result[code_col].astype(str).str.strip() == 'None') |  # текстовый 'None'
            (result[code_col].astype(str).str.strip() == 'null') |  # текстовый 'null'
            (result[code_col].astype(str).str.strip() == '')        # пустая строка
        )
        
        # 2. Проекта нет в автокодификации
        if autocoding_df is not None:
            ak_code_col = 'ИТОГО КОД'
            
            # Создаем множество уникальных кодов из АК
            ak_codes = set()
            for code in autocoding_df[ak_code_col].astype(str).str.strip():
                if code and code.lower() not in ['nan', 'none', 'null', '']:
                    if code.startswith('RU'):
                        ak_codes.add(code)
            
            # Проверяем коды проектов
            result['Проекта НЕТ в АК'] = [
                (str(code).strip() not in ak_codes) 
                if pd.notna(code) and str(code).strip().startswith('RU')
                else False
                for code in result[code_col]
            ]
       
        else:
            result['Проекта НЕТ в АК'] = True
        
        # 3. Проект есть в массиве, но не полевой (есть в массиве и помечен как неполевой)
        if array_df is not None:
            array_code_col = 'Код анкеты'
            array_project_col = 'Название проекта'
            
            # Все проекты из массива
            all_array_keys = set(zip(
                array_df[array_code_col].astype(str).str.strip().fillna(''),
                array_df[array_project_col].astype(str).str.strip().fillna('')
            ))
            
            # Неполевые проекты из массива
            non_field_mask = array_df['Полевой'] == 0
            non_field_keys = set(zip(
                array_df.loc[non_field_mask, array_code_col].astype(str).str.strip().fillna(''),
                array_df.loc[non_field_mask, array_project_col].astype(str).str.strip().fillna('')
            ))
            
            # Ключи из гугл таблицы
            google_project_keys = list(zip( 
                result[code_col].astype(str).str.strip().fillna(''),
                result[project_col].astype(str).str.strip().fillna('')
            ))
            
            # Проверяем: проект есть в массиве И он неполевой
            result['Проект есть в массиве, но не полевой'] = [
                key in all_array_keys and key in non_field_keys 
                for key in google_project_keys
            ]
        else:
            result['Проект есть в массиве, но не полевой'] = False
        
        # 4. Проекта нет в массиве (только Чеккер)
        if array_df is not None:
            array_code_col = 'Код анкеты'
            array_project_col = 'Название проекта'
            
            # Все проекты из массива
            all_array_keys = set(zip(
                array_df[array_code_col].astype(str).str.strip().fillna(''),
                array_df[array_project_col].astype(str).str.strip().fillna('')
            ))
            
            # Только чеккер проекты из гугл
            checker_mask = result[portal_col] == 'Чеккер'
            google_checker_keys = list(zip(
                result.loc[checker_mask, code_col].astype(str).str.strip().fillna(''),
                result.loc[checker_mask, wave_col].astype(str).str.strip().fillna('')
            ))
            
            # Создаем серию для результата
            result['Проекта нет в массиве'] = False
            result.loc[checker_mask, 'Проекта нет в массиве'] = [
                key not in all_array_keys for key in google_checker_keys
            ]
        else:
            result['Проекта нет в массиве'] = False
        
        # Выбираем нужные колонки
        columns = [
            'ФИО ОМ', project_col, 
            'Сценарий, если разные квоты и сроки', code_col,
            wave_col,
            'Дата старта', 'Дата финиша с продлением',
            'вводные запрошены / вводные получены, готовится старт / стартовал',
            portal_col,
            'Код проекта пусто', 'Проекта НЕТ в АК', 'Проект есть в массиве, но не полевой', 'Проекта нет в массиве'
        ]

        # Только существующие колонки
        existing_cols = [col for col in columns if col in result.columns]
        result = result[existing_cols]

        # ============================================
        # 🆕 ФИЛЬТРАЦИЯ: ОСТАВЛЯЕМ ТОЛЬКО ПРОБЛЕМНЫЕ ПРОЕКТЫ
        # ============================================
        
        # Колонки с проверками
        check_columns = ['Код проекта пусто', 'Проекта НЕТ в АК', 'Проект есть в массиве, но не полевой', 'Проекта нет в массиве']
        
        # Оставляем только те, что есть в результате
        existing_checks = [col for col in check_columns if col in result.columns]
        
        if existing_checks:
            # Находим строки где ХОТЯ БЫ ОДНА проверка = True
            problem_mask = result[existing_checks].any(axis=1)
            
            # Фильтруем - оставляем только проблемные
            problematic_only = result[problem_mask].copy()
            
            return problematic_only
        else:
            # Если нет колонок с проверками - возвращаем пустой DataFrame
            return pd.DataFrame()

    
    def clean_cxway(self, df, hierarchy_df, google_df):
        """Очистка файла CXWAY и приведение к структуре полевых проектов"""
        if df is None or df.empty:
            return pd.DataFrame()
        
        df_clean = df.copy()
        
        # 1. Базовые колонки
        column_mapping = {
            'Код анкеты': ['Код проекта', 'Project Code', 'Код', 'Code'],
            'Имя клиента': ['Клиент', 'Client', 'Имя клиента', 'Клиент имя'],
            'Название проекта': ['Wave Name'],
            'АСС': ['АСС', 'ACC', 'Аккаунт'],
            'ЭМ': ['ЭМ', 'EM', 'ЭМ рег', 'Эксперт', 'Эксперт менеджер'],
            'Регион short': ['Регион short', 'Регион кратко', 'Region short', 'Регион']
        }
        
        # Создаем стандартный DataFrame
        result = pd.DataFrame()
        
        for std_col, possible_names in column_mapping.items():
            source_col = self._find_column(df_clean, possible_names)
            if source_col:
                result[std_col] = df_clean[source_col].astype(str).fillna('')
            else:
                result[std_col] = ''
        
        # 2. Определение "Полевой" (логика как в массиве)
        if 'Код анкеты' in result.columns and not result.empty:
            result['Полевой'] = result['Код анкеты'].apply(self._is_field_project)
            # Оставляем только полевые проекты
            result = result[result['Полевой'] == 1]
        else:
            result['Полевой'] = 1
        
        # 3. Добавление ЗОД из справочника (по АСС)
        if hierarchy_df is not None and 'АСС' in result.columns:
            zodiac_col = self._find_column(hierarchy_df, ['ЗОД', 'zod', 'ZOD'])
            acc_col = self._find_column(hierarchy_df, ['АСС', 'acc', 'ACC'])
            
            if zodiac_col and acc_col:
                zod_mapping = {}
                for _, row in hierarchy_df.iterrows():
                    acc_val = str(row[acc_col]).strip()
                    zod_val = str(row[zodiac_col]).strip()
                    if acc_val:
                        zod_mapping[acc_val] = zod_val
                
                def get_zod_by_acc(acc_value):
                    if pd.isna(acc_value) or str(acc_value).strip() == '':
                        return ''
                    clean_acc = str(acc_value).strip()
                    return zod_mapping.get(clean_acc, '')
                
                result['ЗОД'] = result['АСС'].apply(get_zod_by_acc)
            else:
                result['ЗОД'] = ''
        else:
            result['ЗОД'] = ''
        
        # 4. Добавление ПО из гугла (по коду проекта)
        if google_df is not None and 'Код анкеты' in result.columns:
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            
            if google_code_col and google_portal_col:
                portal_mapping = {}
                for _, row in google_df.iterrows():
                    code = str(row[google_code_col]).strip()
                    portal = str(row[google_portal_col]).strip()
                    if code:
                        portal_mapping[code] = portal
                
                def get_portal(code_value):
                    if pd.isna(code_value) or str(code_value).strip() == '':
                        return 'не определено'
                    clean_code = str(code_value).strip()
                    return portal_mapping.get(clean_code, 'не определено')
                
                result['ПО'] = result['Код анкеты'].apply(get_portal)
            else:
                result['ПО'] = 'не определено'
        else:
            result['ПО'] = 'не определено'
        
        # 5. Добавление полного региона из справочника
        region_mapping = {
            'AD': 'Республика Адыгея',
            'AL': 'Алтайский край',
            'AM': 'Амурская область',
            'AR': 'Архангельская область',
            'AS': 'Астраханская область',
            'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область',
            'BR': 'Брянская область',
            'BU': 'Республика Бурятия',
            'CK': 'Чукотский автономный округ',
            'CL': 'Челябинская область',
            'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика',
            'DA': 'Республика Дагестан',
            'DN': 'Донецкая Народная Республика',
            'GA': 'Республика Алтай',
            'IN': 'Республика Ингушетия',
            'IR': 'Иркутская область',
            'IV': 'Ивановская область',
            'KA': 'Камчатский край',
            'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика',
            'KD': 'Краснодарский край',
            'KE': 'Кемеровская область',
            'KG': 'Калужская область',
            'KH': 'Хабаровский край',
            'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия',
            'KL': 'Республика Калмыкия',
            'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область',
            'KO': 'Республика Коми',
            'KS': 'Курская область',
            'KT': 'Костромская область',
            'KU': 'Курганская область',
            'KV': 'Кировская область',
            'KY': 'Красноярский край',
            'LG': 'Луганская Народная Республика',
            'LN': 'Ленинградская область',
            'LP': 'Липецкая область',
            'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область',
            'MM': 'Мурманская область',
            'MR': 'Республика Мордовия',
            'MS': 'Московская область',
            'NG': 'Новгородская область',
            'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия',
            'NS': 'Новосибирская область',
            'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область',
            'OL': 'Орловская область',
            'OM': 'Омская область',
            'PE': 'Пермский край',
            'PR': 'Приморский край',
            'PS': 'Псковская область',
            'PZ': 'Пензенская область',
            'RK': 'Республика Крым',
            'RO': 'Ростовская область',
            'RZ': 'Рязанская область',
            'SA': 'Самарская область',
            'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область',
            'SM': 'Смоленская область',
            'SR': 'Саратовская область',
            'ST': 'Ставропольский край',
            'SV': 'Свердловская область',
            'TB': 'Тамбовская область',
            'TL': 'Тульская область',
            'TO': 'Томская область',
            'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва',
            'TV': 'Тверская область',
            'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика',
            'UL': 'Ульяновская область',
            'VG': 'Волгоградская область',
            'VL': 'Владимирская область',
            'VO': 'Вологодская область',
            'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ',
            'YS': 'Ярославская область',
            'YV': 'Еврейская автономная область',
            'ZO': 'Запорожская область',
            'ZK': 'Забайкальский край'
        }
        
        if 'Регион short' in result.columns:
            def get_full_region(short):
                if pd.isna(short) or str(short).strip() == '':
                    return 'не определен'
                short_clean = str(short).strip().upper()
                return region_mapping.get(short_clean, 'не определен')
            
            result['Регион'] = result['Регион short'].apply(get_full_region)
        else:
            result['Регион'] = 'не определен'
        
        # 6. Метаданные
        result['Источник'] = 'CXWAY'
        
        return result
    
    def _is_field_project(self, code):
        """Логика определения полевого проекта (как в массиве)"""
        try:
            if pd.isna(code):
                return 0
                
            code_str = str(code).strip()
            lower_code = code_str.lower()
            
            # Проверка на мультикод/пилот/семпл
            if any(word in lower_code for word in ['мультикод', 'пилот', 'семпл']):
                return 1
            
            # Проверка формата RU00.001.06.01SVZ24
            parts = code_str.split('.')
            if len(parts) >= 4:
                country = parts[0]  # RU00
                if len(parts[2]) >= 2:
                    direction = '.' + parts[2][:2]  # .01 или .02
                else:
                    direction = ''
                
                if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
                    return 1
                    
            return 0
        except:
            return 0
    
    def merge_field_projects(self, array_field_df, cxway_field_df):
        try:
            # 1. Фильтруем массив
            if array_field_df is not None and not array_field_df.empty:
                mask_array = (array_field_df['ПО'] == 'Чеккер') | (array_field_df['ПО'] == 'не определено')
                array_filtered = array_field_df[mask_array].copy()
            else:
                array_filtered = pd.DataFrame()
            
            # 2. Фильтруем CXWAY
            if cxway_field_df is not None and not cxway_field_df.empty:
                mask_cxway = (cxway_field_df['ПО'] == 'CXWAY') | (cxway_field_df['ПО'] == 'не определено')
                cxway_filtered = cxway_field_df[mask_cxway].copy()
            else:
                cxway_filtered = pd.DataFrame()
            
            # 3. Объединяем
            if not array_filtered.empty and not cxway_filtered.empty:
                combined = pd.concat([array_filtered, cxway_filtered], ignore_index=True)
            elif not array_filtered.empty:
                combined = array_filtered
            elif not cxway_filtered.empty:
                combined = cxway_filtered
            else:
                combined = pd.DataFrame()
            
            return combined
            
        except Exception as e:
            st.error(f"Ошибка объединения: {e}")
            return array_field_df if array_field_df is not None else pd.DataFrame()
        
    def add_portal_to_array(self, array_df, google_df):
        """
        Добавляет колонку 'ПО' в массив из гугл таблицы
        Сопоставление: Код анкеты (массив) → Код проекта (гугл) → ПО
        """
        try:
            array_df = array_df.copy()
            
            # Ищем колонку Код анкеты
            array_code_col = None
            for col in array_df.columns:
                if 'код' in str(col).lower() and 'анкет' in str(col).lower():
                    array_code_col = col
                    break
            
            if not array_code_col:
                array_df['ПО'] = 'не определено'
                return array_df
                
            # Колонки в гугл
            google_code_col = 'Код проекта RU00.000.00.01SVZ24'
            google_portal_col = 'Портал на котором идет проект (для работы полевой команды)'
            
            # Создаем словарь {код_проекта: ПО}
            portal_mapping = {}
            for _, row in google_df.iterrows():
                code = str(row.get(google_code_col, '')).strip()
                portal = str(row.get(google_portal_col, '')).strip()
                
                if code and code.lower() not in ['nan', 'none', 'null', '']:
                    portal_mapping[code] = portal
            
            # Добавляем колонку ПО
            def get_portal(code):
                if pd.isna(code):
                    return 'не определено'
                clean_code = str(code).strip()
                return portal_mapping.get(clean_code, 'не определено')
            
            array_df['ПО'] = array_df[array_code_col].apply(get_portal)
            
            defined_count = (array_df['ПО'] != 'не определено').sum()
            st.success(f"✅ Добавлено ПО: определено для {defined_count} из {len(array_df)} записей")
            
            return array_df
            
        except Exception as e:
            st.error(f"Ошибка в add_portal_to_array: {e}")
            return array_df

# Глобальный экземпляр
data_cleaner = DataCleaner()









