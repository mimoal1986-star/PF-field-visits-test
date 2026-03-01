# utils/data_cleaner.py
# draft 4.0 - simplified
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
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
        Очистка Гугл таблицы (Проекты Сервизория)
        """
        if df is None or df.empty:
            return None
        
        df_clean = df.copy()
        
        # Удаление дубликатов
        code_field = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24'])
        start_date_field = self._find_column(df_clean, ['Дата старта'])
        end_date_field = self._find_column(df_clean, ['Дата финиша с продлением'])
        
        existing_fields = []
        if code_field:
            existing_fields.append(code_field)
        if start_date_field:
            existing_fields.append(start_date_field)
        if end_date_field:
            existing_fields.append(end_date_field)
        
        if len(existing_fields) >= 1:
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first')
        
        # Чистка пробелов в кодах
        code_col = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24'])
        if code_col:
            df_clean[code_col] = df_clean[code_col].astype(str).str.strip()
        
        # Форматирование Пилоты/Семплы/Мультикоды
        if code_col:
            target_values = ['пилот', 'семпл', 'мультикод']
            for idx, value in df_clean[code_col].items():
                if pd.isna(value):
                    continue
                str_value = str(value).strip()
                lower_value = str_value.lower()
                for target in target_values:
                    if target in lower_value:
                        df_clean.at[idx, code_col] = str_value.capitalize()
                        break
        
        # Конвертация дат
        date_cols = []
        for col_name in ['Дата старта', 'Дата финиша с продлением']:
            found = self._find_column(df_clean, [col_name])
            if found:
                date_cols.append(found)
        
        for col in date_cols:
            try:
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
            except:
                pass
        
        # Бизнес-правила для дат
        if 'plan_calc_params' in st.session_state:
            start_period = st.session_state.plan_calc_params['start_date']
            end_period = st.session_state.plan_calc_params['end_date']
            
            first_day = pd.Timestamp(year=end_period.year, month=end_period.month, day=1)
            last_day = first_day + pd.offsets.MonthEnd(1)
            
            start_cols = [col for col in date_cols if 'старт' in str(col).lower()]
            end_cols = [col for col in date_cols if 'финиш' in str(col).lower()]
            
            for col in start_cols:
                try:
                    mask = df_clean[col] < first_day
                    if mask.any():
                        df_clean.loc[mask, col] = first_day
                except:
                    pass
            
            for col in end_cols:
                try:
                    mask_too_early = df_clean[col] < first_day
                    if mask_too_early.any():
                        df_clean.loc[mask_too_early, col] = last_day
                    
                    mask_too_late = df_clean[col] > last_day
                    if mask_too_late.any():
                        df_clean.loc[mask_too_late, col] = last_day
                except:
                    pass
        
        return df_clean

    def clean_array(self, df):
        """Очистка файла Массив"""
        if df is None or df.empty:
            return None
        
        df_clean = df.copy()
        
        # Обработка дат
        DATE_COLUMNS = [
            'Дата визита',
            'Дата создания проверки', 
            'Дата назначения опроса за тайным покупателем',
            'Дата подтверждения опроса тайным покупателем',
            'Время окончания',
            'Время завершения ожидания статуса утверждения',
            'Время утверждения'
        ]
        
        existing_date_cols = [col for col in DATE_COLUMNS if col in df_clean.columns]
        
        if existing_date_cols:
            SURROGATE_DATE = pd.Timestamp('1900-01-01')
            
            for col in existing_date_cols:
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
                    nat_mask = df_clean[col].isna()
                    if nat_mask.any():
                        df_clean.loc[nat_mask, col] = SURROGATE_DATE
                except:
                    pass
        
        # Замена Н/Д на пустые значения
        na_values = ['Н/Д', 'н/д', 'N/A', 'n/a', '#Н/Д', '#н/д', 'NA', 'na', '-', '—', '–']
        
        for col in df_clean.columns:
            try:
                for na_val in na_values:
                    mask = df_clean[col].astype(str).str.strip() == na_val
                    if mask.any():
                        df_clean.loc[mask, col] = ''
                
                nan_mask = df_clean[col].astype(str).str.strip().str.lower().isin(['nan', 'none', 'null'])
                if nan_mask.any():
                    df_clean.loc[nan_mask, col] = ''
            except:
                pass
        
        # Добавление колонки ЗОД (пустая, будет заполнена позже)
        if 'ЗОД' not in df_clean.columns:
            df_clean['ЗОД'] = ''
        
        return df_clean
    
    def enrich_array_with_project_codes(self, array_df, projects_df):
        """
        Заполняет пустые 'Код анкеты' в Массиве из таблицы Проектов
        """
        array_df = array_df.copy()
        projects_df = projects_df.copy()
        
        # Находим строки с пустым кодом
        code_col = self._find_column(array_df, ['Код анкеты'])
        if not code_col:
            return array_df, pd.DataFrame(), {'filled': 0, 'total': 0}
        
        empty_code_mask = (
            array_df[code_col].isna() |
            (array_df[code_col].astype(str).str.strip() == '')
        )
        rows_to_process = array_df[empty_code_mask]
        total_empty = len(rows_to_process)
        
        if total_empty == 0:
            return array_df, pd.DataFrame(), {'filled': 0, 'total': 0}
        
        # Подготовка проектов для поиска
        project_client_col = self._find_column(projects_df, ['Проекты в  https://ru.checker-soft.com'])
        project_wave_col = self._find_column(projects_df, ['Название волны на Чекере/ином ПО'])
        project_code_col = self._find_column(projects_df, ['Код проекта RU00.000.00.01SVZ24'])
        
        if not all([project_client_col, project_wave_col, project_code_col]):
            return array_df, pd.DataFrame(), {'filled': 0, 'total': total_empty}
        
        projects_df['_match_client'] = projects_df[project_client_col].astype(str).str.strip()
        projects_df['_match_wave'] = projects_df[project_wave_col].astype(str).str.strip()
        
        # Поиск и заполнение
        filled_count = 0
        discrepancy_rows = []
        
        client_col = self._find_column(array_df, ['Имя клиента'])
        wave_col = self._find_column(array_df, ['Название проекта'])
        
        for idx, row in rows_to_process.iterrows():
            client_name = str(row[client_col]).strip() if pd.notna(row[client_col]) else ''
            wave_name = str(row[wave_col]).strip() if pd.notna(row[wave_col]) else ''
            
            match_mask = (
                (projects_df['_match_client'] == client_name) &
                (projects_df['_match_wave'] == wave_name)
            )
            
            matched_rows = projects_df[match_mask]
            
            if not matched_rows.empty:
                project_code = matched_rows.iloc[0][project_code_col]
                if pd.notna(project_code) and str(project_code).strip() != '':
                    array_df.at[idx, code_col] = str(project_code).strip()
                    filled_count += 1
                else:
                    discrepancy_rows.append(row.to_dict())
            else:
                discrepancy_rows.append(row.to_dict())
        
        discrepancy_df = pd.DataFrame(discrepancy_rows) if discrepancy_rows else pd.DataFrame()
        
        return array_df, discrepancy_df, {'filled': filled_count, 'total': total_empty}
    
    def update_field_projects_flag(self, google_df):
        """
        Обновляет поле 'Полевой' в гугл таблице на основе кода проекта
        """
        try:
            google_df = google_df.copy()
            
            code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24'])
            if not code_col:
                google_df['Полевой'] = 0
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
                        country = parts[0]
                        if len(parts[2]) >= 2:
                            direction = '.' + parts[2][:2]
                        else:
                            direction = ''
                        
                        if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
                            return 1
                            
                    return 0
                except:
                    return 0
                            
            google_df['Полевой'] = google_df[code_col].apply(is_field_project)
            
            return google_df
            
        except Exception as e:
            google_df['Полевой'] = 0
            return google_df

    def add_field_flag_to_array(self, array_df):
        """
        Добавляет 'Полевой' в массив на основе кода анкеты
        """
        try:
            array_df = array_df.copy()
            
            code_col = self._find_column(array_df, ['Код анкеты'])
            if not code_col:
                array_df['Полевой'] = 0
                return array_df
            
            def is_field_project(code):
                try:
                    if pd.isna(code):
                        return 0
                        
                    code_str = str(code).strip()
                    lower_code = code_str.lower()
                    
                    if any(word in lower_code for word in ['мультикод', 'пилот', 'семпл']):
                        return 1
                    
                    parts = code_str.split('.')
                    if len(parts) >= 4:
                        country = parts[0]
                        if len(parts[2]) >= 2:
                            direction = '.' + parts[2][:2]
                        else:
                            direction = ''
                        
                        if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
                            return 1
                            
                    return 0
                except:
                    return 0
            
            array_df['Полевой'] = array_df[code_col].apply(is_field_project)
            
            return array_df
            
        except Exception as e:
            array_df['Полевой'] = 0
            return array_df
    
    def add_portal_to_array(self, array_df, google_df):
        """
        Добавляет колонку 'ПО' в массив из гугл таблицы
        """
        try:
            array_df = array_df.copy()
            
            array_code_col = self._find_column(array_df, ['Код анкеты'])
            if not array_code_col:
                array_df['ПО'] = 'не определено'
                return array_df
            
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект', 'ПО'])
            
            if not google_code_col or not google_portal_col:
                array_df['ПО'] = 'не определено'
                return array_df
            
            portal_mapping = {}
            for _, row in google_df.iterrows():
                code = str(row.get(google_code_col, '')).strip()
                portal = str(row.get(google_portal_col, '')).strip()
                
                if code and code.lower() not in ['nan', 'none', 'null', '']:
                    portal_mapping[code] = portal
            
            def get_portal(code):
                if pd.isna(code):
                    return 'не определено'
                clean_code = str(code).strip()
                return portal_mapping.get(clean_code, 'не определено')
            
            array_df['ПО'] = array_df[array_code_col].apply(get_portal)
            
            return array_df
            
        except Exception as e:
            array_df['ПО'] = 'не определено'
            return array_df
    
    def add_zod_from_mapping(self, field_df, zod_acc_mapping):
        """
        Добавляет ЗОД в полевые проекты на основе встроенного справочника
        """
        try:
            if field_df is None or field_df.empty:
                return field_df
            
            field_df = field_df.copy()
            
            # Создаем обратный словарь {АСС: ЗОД}
            acc_to_zod = {}
            for zod, acc_list in zod_acc_mapping.items():
                for acc in acc_list:
                    acc_to_zod[acc] = zod
            
            # Находим колонку АСС
            acc_col = self._find_column(field_df, ['АСС', 'acc', 'ACC'])
            if not acc_col:
                return field_df
            
            def get_zod(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip() == '':
                    return ''
                clean_acc = str(acc_value).strip()
                return acc_to_zod.get(clean_acc, '')
            
            field_df['ЗОД'] = field_df[acc_col].apply(get_zod)
            
            return field_df
            
        except Exception as e:
            return field_df
    
    def split_array_by_field_flag(self, array_df):
        """
        Разделяет массив на Полевые и Неполевые проекты
        """
        try:
            if array_df is None or array_df.empty:
                return None, None
            
            array_df_clean = array_df.copy()
            
            if 'Полевой' not in array_df_clean.columns:
                return None, None
            
            # Определяем нужные колонки
            column_mapping = {
                'Код анкеты': ['Код анкеты'],
                'Имя клиента': ['Имя клиента'],
                'Название проекта': ['Название проекта'],
                'ЗОД': ['ЗОД'],
                'АСС': ['АСС', 'ACC'],
                'ЭМ': ['ЭМ рег', 'ЭМ'],
                'Регион short': ['Регион short', 'Регион'],
                'Регион': ['Регион '],
                'Полевой': ['Полевой'],
                'Статус': ['Статус'],
                'Дата визита': ['Дата визита'],
                'ПО': ['ПО']
            }
            
            actual_columns = {}
            for std_col, possible_names in column_mapping.items():
                found_col = self._find_column(array_df_clean, possible_names)
                if found_col:
                    actual_columns[std_col] = found_col
            
            if not actual_columns:
                return None, None
            
            selected_cols = list(actual_columns.values())
            
            field_mask = array_df_clean['Полевой'] == 1
            field_projects = array_df_clean.loc[field_mask, selected_cols].copy()
            non_field_projects = array_df_clean.loc[~field_mask, selected_cols].copy()
            
            # Переименовываем колонки
            reverse_mapping = {v: k for k, v in actual_columns.items()}
            
            if not field_projects.empty:
                field_projects = field_projects.rename(columns=reverse_mapping)
            
            if not non_field_projects.empty:
                non_field_projects = non_field_projects.rename(columns=reverse_mapping)
            
            return field_projects, non_field_projects
            
        except Exception as e:
            return None, None
    
    def check_problematic_projects(self, google_df, array_df):
        """
        Проверка проблемных проектов (без автокодификации)
        """
        if google_df is None or google_df.empty:
            return pd.DataFrame()
        
        result = google_df.copy()
        
        code_col = self._find_column(result, ['Код проекта RU00.000.00.01SVZ24'])
        project_col = self._find_column(result, ['Проекты в  https://ru.checker-soft.com'])
        wave_col = self._find_column(result, ['Название волны на Чекере/ином ПО'])
        portal_col = self._find_column(result, ['Портал на котором идет проект', 'ПО'])
        
        if not code_col:
            return pd.DataFrame()
        
        # 1. Код проекта пусто
        result['Код проекта пусто'] = (
            result[code_col].isna() |
            (result[code_col].astype(str).str.strip() == 'nan') |
            (result[code_col].astype(str).str.strip() == 'NaN') |
            (result[code_col].astype(str).str.strip() == 'None') |
            (result[code_col].astype(str).str.strip() == 'null') |
            (result[code_col].astype(str).str.strip() == '')
        )
        
        # 2. Проект есть в массиве, но не полевой
        if array_df is not None and not array_df.empty:
            array_code_col = self._find_column(array_df, ['Код анкеты'])
            array_project_col = self._find_column(array_df, ['Название проекта'])
            
            if array_code_col and array_project_col and 'Полевой' in array_df.columns:
                all_array_keys = set(zip(
                    array_df[array_code_col].astype(str).str.strip().fillna(''),
                    array_df[array_project_col].astype(str).str.strip().fillna('')
                ))
                
                non_field_mask = array_df['Полевой'] == 0
                non_field_keys = set(zip(
                    array_df.loc[non_field_mask, array_code_col].astype(str).str.strip().fillna(''),
                    array_df.loc[non_field_mask, array_project_col].astype(str).str.strip().fillna('')
                ))
                
                google_project_keys = list(zip(
                    result[code_col].astype(str).str.strip().fillna(''),
                    result[project_col].astype(str).str.strip().fillna('') if project_col else [''] * len(result)
                ))
                
                result['Проект есть в массиве, но не полевой'] = [
                    key in all_array_keys and key in non_field_keys 
                    for key in google_project_keys
                ]
            else:
                result['Проект есть в массиве, но не полевой'] = False
        else:
            result['Проект есть в массиве, но не полевой'] = False
        
        # 3. Проекта нет в массиве (только Чеккер)
        if array_df is not None and portal_col and wave_col:
            array_code_col = self._find_column(array_df, ['Код анкеты'])
            array_project_col = self._find_column(array_df, ['Название проекта'])
            
            if array_code_col and array_project_col:
                all_array_keys = set(zip(
                    array_df[array_code_col].astype(str).str.strip().fillna(''),
                    array_df[array_project_col].astype(str).str.strip().fillna('')
                ))
                
                checker_mask = result[portal_col].astype(str).str.strip() == 'Чеккер'
                google_checker_keys = list(zip(
                    result.loc[checker_mask, code_col].astype(str).str.strip().fillna(''),
                    result.loc[checker_mask, wave_col].astype(str).str.strip().fillna('')
                ))
                
                result['Проекта нет в массиве'] = False
                if len(google_checker_keys) > 0:
                    result.loc[checker_mask, 'Проекта нет в массиве'] = [
                        key not in all_array_keys for key in google_checker_keys
                    ]
            else:
                result['Проекта нет в массиве'] = False
        else:
            result['Проекта нет в массиве'] = False
        
        # Формируем результат
        columns = []
        if project_col:
            columns.append(project_col)
        if code_col:
            columns.append(code_col)
        if wave_col:
            columns.append(wave_col)
        if portal_col:
            columns.append(portal_col)
        
        columns.extend(['Код проекта пусто', 'Проект есть в массиве, но не полевой', 'Проекта нет в массиве'])
        
        existing_cols = [col for col in columns if col in result.columns]
        if not existing_cols:
            return pd.DataFrame()
        
        result = result[existing_cols]
        
        check_columns = ['Код проекта пусто', 'Проект есть в массиве, но не полевой', 'Проекта нет в массиве']
        existing_checks = [col for col in check_columns if col in result.columns]
        
        if existing_checks:
            problem_mask = result[existing_checks].any(axis=1)
            return result[problem_mask].copy()
        else:
            return pd.DataFrame()

# Глобальный экземпляр
data_cleaner = DataCleaner()
