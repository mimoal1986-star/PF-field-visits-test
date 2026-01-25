# utils/data_cleaner.py
# draft 1.3

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
        # ... существующий код остается без изменений ...
        # (пропускаю для краткости, он правильный)
        
        return df_clean

    def clean_array(self, df):
        """Очистка файла Массив"""
        if df is None or df.empty:
            st.warning("⚠️ Массив пустой или не загружен")
            return None
        
        df_clean = df.copy()
        original_rows = len(df_clean)
        original_cols = len(df_clean.columns)
        
        st.info(f"🧹 Начинаю очистку Массива: {original_rows} строк × {original_cols} колонок")
        
        # === Удалить нули в датах ===
        st.write("**1️⃣ Заменяю нули в датах на суррогатную дату (1900-01-01)...**")
        
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
            SURROGATE_DATE = pd.Timestamp('1900-01-01')
            total_replacements = 0
            
            for col in existing_date_cols:
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    nat_mask = df_clean[col].isna()
                    
                    if nat_mask.any():
                        df_clean.loc[nat_mask, col] = SURROGATE_DATE
                        col_replacements = nat_mask.sum()
                        total_replacements += col_replacements
                        
                        example_indices = nat_mask[nat_mask].index[:3]
                        if len(example_indices) > 0:
                            st.info(f"   '{col}': заменено {col_replacements} значений")
                            
                except Exception as e:
                    st.warning(f"   Ошибка в колонке '{col}': {str(e)[:100]}")
            
            if total_replacements > 0:
                st.success(f"   ✅ Заменено {total_replacements} невалидных дат на {SURROGATE_DATE.date()}")
                st.info("   **Обозначает:** 'Событие еще не наступило'")
            else:
                st.info("   ℹ️ Невалидных дат не найдено")
        else:
            st.warning(f"   ⚠️ Не найдено ни одной колонки с датами")
        
        # === Проверить массив на Н/Д ===
        st.write("**2️⃣ Заменяю Н/Д на пустые значения...**")
        
        na_values = ['Н/Д', 'н/д', 'N/A', 'n/a', '#Н/Д', '#н/д', 'NA', 'na', '-', '—', '–']
        na_replacements = 0
        
        for col in df_clean.columns:
            try:
                for na_val in na_values:
                    mask = df_clean[col].astype(str).str.strip() == na_val
                    if mask.any():
                        df_clean.loc[mask, col] = ''
                        na_replacements += mask.sum()
                
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
        
        # === Сохраняем информацию о строках с Н/Д для отчета ===
        st.write("**3️⃣ Сохраняю информацию о строках с Н/Д для отчета...**")
        
        # ФИКС: Используем df вместо original_df которого нет
        original_df = df.copy()  # Сохраняем оригинальные данные
        
        had_na_mask = pd.Series(False, index=df_clean.index)
        
        for col in df_clean.columns:
            try:
                for na_val in na_values:
                    mask = original_df[col].astype(str).str.strip() == na_val
                    had_na_mask = had_na_mask | mask
            except:
                continue
        
        df_clean.attrs['had_na_rows'] = had_na_mask
        df_clean.attrs['na_rows_count'] = had_na_mask.sum()
        
        st.success(f"   ✅ Сохранено {had_na_mask.sum()} строк с Н/Д для отчета")
        
        return df_clean

    # ... остальные существующие методы остаются без изменений ...

    def update_field_projects_flag(self, google_df, autocoding_df):
        """
        Обновляет поле 'Полевой' в гугл таблице
        Полевой = 1 если код есть в АК И направление .01/.02
        """
        try:
            if google_df is None or google_df.empty:
                st.warning("⚠️ Гугл таблица пустая")
                return google_df
                
            if autocoding_df is None or autocoding_df.empty:
                st.warning("⚠️ Автокодификация пустая")
                return google_df
            
            google_df_clean = google_df.copy()
            autocoding_df_clean = autocoding_df.copy()
            
            # Находим ключевые колонки
            google_code_col = self._find_column(google_df_clean, [
                'Код проекта RU00.000.00.01SVZ24',
                'Код проекта',
                'Project Code',
                'Код'
            ])
            
            if not google_code_col:
                st.error("❌ В гугл таблице не найдена колонка с кодом проекта")
                return google_df
            
            ak_code_col = self._find_column(autocoding_df_clean, [
                'ИТОГО КОД',
                'ИтогоКод',
                'Код проекта',
                'Код'
            ])
            
            ak_direction_col = self._find_column(autocoding_df_clean, [
                'Направление',
                'Direction',
                'Напр'
            ])
            
            if not ak_code_col:
                st.error("❌ В автокодификации не найдена колонка 'ИТОГО КОД'")
                return google_df
                
            if not ak_direction_col:
                st.warning("⚠️ В автокодификации не найдена колонка 'Направление'")
                autocoding_df_clean[ak_direction_col] = ''
            
            # Предварительная очистка данных
            autocoding_df_clean[ak_code_col] = autocoding_df_clean[ak_code_col].astype(str).str.strip()
            autocoding_df_clean[ak_direction_col] = autocoding_df_clean[ak_direction_col].astype(str).str.strip()
            google_df_clean[google_code_col] = google_df_clean[google_code_col].astype(str).str.strip()
            
            # Создаем множество разрешенных направлений
            allowed_directions = {'.01', '.02', '01', '02', '0.01', '0.02', '1', '2'}
            
            # Создаем словарь для быстрого поиска {код: является_полевым}
            field_codes = set()
            
            for _, row in autocoding_df_clean.iterrows():
                try:
                    code = str(row[ak_code_col])
                    direction = str(row[ak_direction_col])
                    
                    if code and code.lower() not in ['nan', 'none', 'null', '']:
                        if direction in allowed_directions:
                            field_codes.add(code)
                except Exception:
                    continue
            
            st.info(f"🔍 Найдено {len(field_codes)} полевых кодов в автокодификации")
            
            # Инициализируем/обновляем колонку 'Полевой'
            if 'Полевой' not in google_df_clean.columns:
                google_df_clean['Полевой'] = 0
            
            # Векторизированная проверка
            google_codes = google_df_clean[google_code_col].astype(str)
            
            def check_field(code):
                if pd.isna(code) or str(code).lower() in ['nan', 'none', 'null', '']:
                    return 0
                return 1 if str(code) in field_codes else 0
            
            google_df_clean['Полевой'] = google_codes.apply(check_field).astype(int)
            
            updated_count = (google_df_clean['Полевой'] == 1).sum()
            st.success(f"✅ Обновлено поле 'Полевой': {updated_count} полевых проектов")
            
            return google_df_clean
            
        except Exception as e:
            st.error(f"❌ Ошибка в update_field_projects_flag: {str(e)[:100]}")
            return google_df

    def add_field_flag_to_array(self, array_df, google_df):
        """
        Добавляет 'Полевой' в массив на основе гугл таблицы
        """
        try:
            if array_df is None or array_df.empty:
                st.warning("⚠️ Массив пустой")
                return array_df
                
            if google_df is None or google_df.empty:
                st.warning("⚠️ Гугл таблиция пустая")
                return array_df
            
            array_df_clean = array_df.copy()
            google_df_clean = google_df.copy()
            
            # Находим колонки с кодами
            array_code_col = self._find_column(array_df_clean, [
                'Код анкеты',
                'Код проекта',
                'Project Code',
                'Код'
            ])
            
            google_code_col = self._find_column(google_df_clean, [
                'Код проекта RU00.000.00.01SVZ24',
                'Код проекта',
                'Project Code',
                'Код'
            ])
            
            if not array_code_col:
                st.error("❌ В массиве не найдена колонка 'Код анкеты'")
                return array_df
                
            if not google_code_col:
                st.error("❌ В гугл таблице не найдена колонка с кодом проекта")
                return array_df
            
            if 'Полевой' not in google_df_clean.columns:
                st.warning("⚠️ В гугл таблице нет колонки 'Полевой', создаю нулевую")
                google_df_clean['Полевой'] = 0
            
            # Очистка данных
            array_df_clean[array_code_col] = array_df_clean[array_code_col].astype(str).str.strip()
            google_df_clean[google_code_col] = google_df_clean[google_code_col].astype(str).str.strip()
            
            # Создаем словарь сопоставления {код: полевое_значение}
            code_to_field = {}
            
            for idx, row in google_df_clean.iterrows():
                try:
                    code = str(row[google_code_col])
                    if code and code.lower() not in ['nan', 'none', 'null', '']:
                        field_val = row.get('Полевой', 0)
                        try:
                            code_to_field[code] = int(field_val) if not pd.isna(field_val) else 0
                        except (ValueError, TypeError):
                            code_to_field[code] = 0
                except Exception:
                    continue
            
            st.info(f"🔍 Загружено {len(code_to_field)} сопоставлений кодов")
            
            # Добавляем колонку в массив
            array_df_clean['Полевой'] = 0
            
            def get_field_value(code):
                if pd.isna(code) or str(code).lower() in ['nan', 'none', 'null', '']:
                    return 0
                return code_to_field.get(str(code), 0)
            
            array_codes = array_df_clean[array_code_col].astype(str)
            array_df_clean['Полевой'] = array_codes.apply(get_field_value).astype(int)
            
            filled_count = (array_df_clean['Полевой'] == 1).sum()
            st.success(f"✅ Добавлен 'Полевой' в массив: {filled_count} полевых записей")
            
            return array_df_clean
            
        except Exception as e:
            st.error(f"❌ Ошибка в add_field_flag_to_array: {str(e)[:100]}")
            return array_df

    def split_array_by_field_flag(self, array_df):
        """
        Разделяет массив на Полевые и Неполевые проекты
        Возвращает только 8 указанных колонок
        """
        try:
            if array_df is None or array_df.empty:
                st.warning("⚠️ Массив пустой")
                return None, None
            
            array_df_clean = array_df.copy()
            
            if 'Полевой' not in array_df_clean.columns:
                st.error("❌ В массиве нет колонки 'Полевой'")
                return None, None
            
            # Определяем маппинг стандартных колонок
            column_mapping = {
                'Код проекта': ['Код проекта', 'Код анкеты', 'Project Code', 'Код'],
                'Имя клиента': ['Имя клиента', 'Клиент', 'Client', 'Client Name'],
                'Название проекта': ['Название проекта', 'Проект', 'Project', 'Project Name'],
                'ЗОД': ['ЗОД', 'ZOD', 'Зод', 'zod'],
                'АСС': ['АСС', 'ASS', 'Асс', 'ass'],
                'ЭМ': ['ЭМ', 'EM', 'Ем', 'em'],
                'Регион short': ['Регион short', 'Регион_short', 'Region_short', 'Short Region'],
                'Регион': ['Регион', 'Region', 'рег']
            }
            
            # Находим фактические названия колонок
            actual_columns = {}
            missing_columns = []
            
            for std_col, possible_names in column_mapping.items():
                found_col = self._find_column(array_df_clean, possible_names)
                if found_col:
                    actual_columns[std_col] = found_col
                else:
                    missing_columns.append(std_col)
            
            if missing_columns:
                st.warning(f"⚠️ Не найдены колонки: {', '.join(missing_columns)}")
                for col in missing_columns:
                    array_df_clean[col] = ''
                    actual_columns[col] = col
            
            # Отбираем нужные колонки + Полевой для фильтрации
            selected_cols = list(actual_columns.values()) + ['Полевой']
            
            # Фильтруем данные
            field_mask = array_df_clean['Полевой'] == 1
            field_projects = array_df_clean.loc[field_mask, selected_cols].copy()
            non_field_projects = array_df_clean.loc[~field_mask, selected_cols].copy()
            
            # Переименовываем колонки
            reverse_mapping = {v: k for k, v in actual_columns.items()}
            
            if not field_projects.empty:
                field_projects = field_projects.rename(columns=reverse_mapping)
                field_projects = field_projects.drop(columns=['Полевой'], errors='ignore')
            
            if not non_field_projects.empty:
                non_field_projects = non_field_projects.rename(columns=reverse_mapping)
                non_field_projects = non_field_projects.drop(columns=['Полевой'], errors='ignore')
            
            # Оставляем только 8 нужных колонок в правильном порядке
            final_columns = list(column_mapping.keys())
            
            if not field_projects.empty:
                field_projects = field_projects.reindex(columns=final_columns)
            
            if not non_field_projects.empty:
                non_field_projects = non_field_projects.reindex(columns=final_columns)
            
            st.success(f"✅ Разделение завершено:")
            st.info(f"   📊 Полевые: {len(field_projects)} записей")
            st.info(f"   📊 Неполевые: {len(non_field_projects)} записей")
            
            return field_projects, non_field_projects
            
        except Exception as e:
            st.error(f"❌ Ошибка в split_array_by_field_flag: {str(e)[:100]}")
            return None, None

    def export_split_array_to_excel(self, field_df, non_field_df, filename="разделенный_массив"):
        """
        Создает Excel с вкладками Полевые/Неполевые проекты
        """
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
                
                # Статистика
                stats_data = {
                    'Метрика': ['Всего записей', 'Полевые', 'Неполевые', 'Дата обработки'],
                    'Значение': [
                        (len(field_df) if field_df is not None else 0) + 
                        (len(non_field_df) if non_field_df is not None else 0),
                        len(field_df) if field_df is not None else 0,
                        len(non_field_df) if non_field_df is not None else 0,
                        pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    ]
                }
                pd.DataFrame(stats_data).to_excel(writer, sheet_name='СТАТИСТИКА', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"❌ Ошибка при создании Excel: {str(e)[:100]}")
            return None


# Глобальный экземпляр (ТОЛЬКО ОДИН РАЗ В КОНЦЕ)
data_cleaner = DataCleaner()
