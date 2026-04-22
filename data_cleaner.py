# utils/data_cleaner.py
# draft 4.1 - simplified
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import streamlit as st
import io

# Встроенный справочник {АСС: ЗОД}
ZOD_MAPPING = {
    'Аблязимова Екатерина': 'Авсейкова Елена',
    'Герасимова Светлана': 'Авсейкова Елена',
    'Карлышева Алиса': 'Герасименко Лика',
    'Механошина Елена': 'Герасименко Лика',
    'Солодникова Виктория': 'Герасименко Лика',
    'Шавлюк Юлия': 'Устинов Игорь',
    'Воронин Евгений': 'Устинов Игорь',
    'Яцевич Максим': 'Устинов Игорь'
}

# Встроенный справочник регионов {код: название}
REGION_MAPPING = {
    'AD': 'Республика Адыгея', 'AL': 'Алтайский край', 'AM': 'Амурская область',
    'AR': 'Архангельская область', 'AS': 'Астраханская область', 'BK': 'Республика Башкортостан',
    'BL': 'Белгородская область', 'BR': 'Брянская область', 'BU': 'Республика Бурятия',
    'CK': 'Чукотский автономный округ', 'CL': 'Челябинская область', 'CN': 'Чеченская Республика',
    'CV': 'Чувашская Республика', 'DA': 'Республика Дагестан', 'DN': 'Донецкая Народная Республика',
    'GA': 'Республика Алтай', 'IN': 'Республика Ингушетия', 'IR': 'Иркутская область',
    'IV': 'Ивановская область', 'KA': 'Камчатский край', 'KB': 'Кабардино-Балкарская Республика',
    'KC': 'Карачаево-Черкесская Республика', 'KD': 'Краснодарский край', 'KE': 'Кемеровская область',
    'KG': 'Калужская область', 'KH': 'Хабаровский край', 'KI': 'Республика Карелия',
    'KK': 'Республика Хакасия', 'KL': 'Республика Калмыкия', 'KM': 'Ханты-Мансийский автономный округ',
    'KN': 'Калининградская область', 'KO': 'Республика Коми', 'KS': 'Курская область',
    'KT': 'Костромская область', 'KU': 'Курганская область', 'KV': 'Кировская область',
    'KY': 'Красноярский край', 'LG': 'Луганская Народная Республика', 'LN': 'Ленинградская область',
    'LP': 'Липецкая область', 'MC': 'Московская область', 'ME': 'Республика Марий Эл',
    'MG': 'Магаданская область', 'MM': 'Мурманская область', 'MR': 'Республика Мордовия',
    'MS': 'Московская область', 'NG': 'Новгородская область', 'NN': 'Ненецкий автономный округ',
    'NO': 'Республика Северная Осетия', 'NS': 'Новосибирская область', 'NZ': 'Нижегородская область',
    'OB': 'Оренбургская область', 'OL': 'Орловская область', 'OM': 'Омская область',
    'PE': 'Пермский край', 'PR': 'Приморский край', 'PS': 'Псковская область',
    'PZ': 'Пензенская область', 'RK': 'Республика Крым', 'RO': 'Ростовская область',
    'RZ': 'Рязанская область', 'SA': 'Самарская область', 'SK': 'Республика Саха (Якутия)',
    'SL': 'Сахалинская область', 'SM': 'Смоленская область', 'SR': 'Саратовская область',
    'ST': 'Ставропольский край', 'SV': 'Свердловская область', 'TB': 'Тамбовская область',
    'TL': 'Тульская область', 'TO': 'Томская область', 'TT': 'Республика Татарстан',
    'TU': 'Республика Тыва', 'TV': 'Тверская область', 'TY': 'Тюменская область',
    'UD': 'Удмуртская Республика', 'UL': 'Ульяновская область', 'VG': 'Волгоградская область',
    'VL': 'Владимирская область', 'VO': 'Вологодская область', 'VR': 'Воронежская область',
    'YN': 'Ямало-Ненецкий автономный округ', 'YS': 'Ярославская область', 'YV': 'Еврейская автономная область',
    'ZK': 'Забайкальский край', 'ZO': 'Запорожская область'
}

# Обратный словарь {название: код}
REGION_NAME_TO_CODE = {v: k for k, v in REGION_MAPPING.items()}

@st.cache_data
def _enrich_array_with_project_codes_cached(array_df, projects_df):
    """Кэшируемая версия обогащения кодами проектов"""
    
    array_df = array_df.copy()
    projects_df = projects_df.copy()
    
    # Находим колонку кода
    code_col = 'Код анкеты' if 'Код анкеты' in array_df.columns else None
    if not code_col:
        return array_df, pd.DataFrame(), {'processed': 0, 'filled': 0, 'discrepancies': 0}
    
    # Находим пустые строки
    empty_mask = (
        array_df[code_col].isna() |
        (array_df[code_col].astype(str).str.strip() == '')
    )
    rows_to_process = array_df[empty_mask]
    total_empty = len(rows_to_process)
    
    if total_empty == 0:
        return array_df, pd.DataFrame(), {'processed': 0, 'filled': 0, 'discrepancies': 0}
    
    # Находим колонки в проектах
    client_col = None
    for name in ['Проекты в  https://ru.checker-soft.com', 'Проекты']:
        if name in projects_df.columns:
            client_col = name
            break
    
    wave_col = None
    for name in ['Название волны на Чекере/ином ПО', 'Волна']:
        if name in projects_df.columns:
            wave_col = name
            break
    
    code_proj_col = None
    for name in ['Код проекта RU00.000.00.01SVZ24', 'Код проекта']:
        if name in projects_df.columns:
            code_proj_col = name
            break
    
    if not all([client_col, wave_col, code_proj_col]):
        return array_df, pd.DataFrame(), {'processed': total_empty, 'filled': 0, 'discrepancies': total_empty}
    
    # Создаем словарь для быстрого поиска
    code_map = {}
    for _, row in projects_df.iterrows():
        client = str(row[client_col]).strip()
        wave = str(row[wave_col]).strip()
        code = str(row[code_proj_col]).strip()
        if client and wave and code and code.lower() not in ['nan', 'none', 'null', '']:
            code_map[(client, wave)] = code
    
    # Заполняем пустые коды
    client_array_col = 'Имя клиента' if 'Имя клиента' in array_df.columns else None
    wave_array_col = 'Название проекта' if 'Название проекта' in array_df.columns else None
    
    filled_count = 0
    if client_array_col and wave_array_col:
        for idx in rows_to_process.index:
            client = str(array_df.loc[idx, client_array_col]).strip()
            wave = str(array_df.loc[idx, wave_array_col]).strip()
            code = code_map.get((client, wave))
            if code:
                array_df.loc[idx, code_col] = code
                filled_count += 1
    
    # Собираем расхождения
    discrepancy_rows = []
    for idx in rows_to_process.index:
        if not array_df.loc[idx, code_col] or str(array_df.loc[idx, code_col]).strip() == '':
            discrepancy_rows.append(array_df.loc[idx].to_dict())
    
    discrepancy_df = pd.DataFrame(discrepancy_rows) if discrepancy_rows else pd.DataFrame()
    
    stats = {
        'processed': total_empty,
        'filled': filled_count,
        'discrepancies': len(discrepancy_df)
    }
    
    return array_df, discrepancy_df, stats
    

class DataCleaner:
    
    def is_non_unique_code(self, code):
        """Проверяет, является ли код неуникальным (Мультикод, Пилот, Семпл и т.д.)"""
        if pd.isna(code):
            return False
        code_str = str(code).strip().lower()
        non_unique_keywords = ['мультикод', 'пилот', 'семпл']
        return any(keyword in code_str for keyword in non_unique_keywords)
    
    def _log_samples(self, df, stage_name):
        """Вспомогательная функция для отладки семплов"""
        if df is None or df.empty:
            st.write(f"🔍 {stage_name}: 0 семплов (DataFrame пуст)")
            return
        
        # Проверяем, есть ли колонка с кодом
        code_col = None
        if 'Код анкеты' in df.columns:
            code_col = 'Код анкеты'
        elif 'Project Code' in df.columns:
            code_col = 'Project Code'
        else:
            st.write(f"🔍 {stage_name}: нет колонки с кодом, пропускаем")
            return
        
        sample_mask = df[code_col].astype(str).str.contains('семпл', case=False, na=False)
        sample_df = df[sample_mask]
        sample_count = len(sample_df)
        
        st.write(f"🔍 {stage_name}: {sample_count} семплов")
        
        if sample_count > 0:
            sample_codes = sample_df[code_col].unique()
            st.write(f"   Коды: {list(sample_codes)}")

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
            return None
        
        df_clean = df.copy()
        
        # === ШАГ 1: Удалить дубликаты записей ===
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
        
        # === ШАГ 2: Сжать пробелы в кодах проектов ===
        code_col = self._find_column(df_clean, ['Код проекта RU00.000.00.01SVZ24'])
        if code_col:
            df_clean[code_col] = df_clean[code_col].astype(str).str.strip()
        
        # === ШАГ 3: Проверка пустых кодов ===
        if code_col:
            empty_mask = (
                df_clean[code_col].isna() | 
                (df_clean[code_col].astype(str).str.strip() == '') |
                (df_clean[code_col].astype(str).str.strip() == 'nan') |
                (df_clean[code_col].astype(str).str.strip() == 'None')
            )
            empty_count = empty_mask.sum()
        
        # === ШАГ 4: Форматировать Пилоты/Семплы/Мультикоды ===
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
        
        # === ШАГ 5: Конвертация дат ===
        date_cols_to_process = []
        for col_name in ['Дата старта', 'Дата финиша с продлением']:
            found = self._find_column(df_clean, [col_name])
            if found:
                date_cols_to_process.append(found)
        
        date_cols = date_cols_to_process
        
        if date_cols:
            for col in date_cols:
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
                except:
                    pass
        
        # === ШАГ 6: Исправить даты по бизнес-правилам ===
        if 'plan_calc_params' in st.session_state:
            end_period = st.session_state['plan_calc_params']['end_date']
            first_day = pd.Timestamp(year=end_period.year, month=end_period.month, day=1)
            last_day = first_day + pd.offsets.MonthEnd(1)
            
            start_date_cols = []
            end_date_cols = []
            
            for col in date_cols:
                col_lower = str(col).lower()
                if any(word in col_lower for word in ['старт', 'начал', 'start']):
                    start_date_cols.append(col)
                elif any(word in col_lower for word in ['финиш', 'конец', 'end']):
                    end_date_cols.append(col)
            
            for col in start_date_cols:
                try:
                    mask = df_clean[col] < first_day
                    if mask.any():
                        df_clean.loc[mask, col] = first_day
                except:
                    pass
            
            for col in end_date_cols:
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
        
        # Удалить строки где Статус == "Удалено"
        status_col = self._find_column(df_clean, ['Статус', ' Статус', 'Статус '])
        if status_col:
            df_clean = df_clean[df_clean[status_col].astype(str).str.strip() != 'Удалено']
        
        # Удалить строки где Дата визита < первый день месяца
        date_col = self._find_column(df_clean, ['Дата визита', 'Date of Visit'])
        if date_col:
            if 'plan_calc_params' in st.session_state:
                first_day = pd.Timestamp(st.session_state['plan_calc_params']['start_date'])
            else:
                today = datetime.now()
                first_day = pd.Timestamp(year=today.year, month=today.month, day=1)
            
            df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors='coerce')
            df_clean = df_clean[pd.isna(df_clean[date_col]) | (df_clean[date_col] >= first_day)]
        
        
        # === Удалить нули в датах ===
        DATE_COLUMNS = [
            'Дата визита',
            'Дата создания проверки', 
            'Дата назначения опроса за тайным покупателем',
            'Дата подтверждения опроса тайным покупателем',
            'Время окончания',
            'Время завершения ожидания статуса утверждения (Дата проведения опроса?)',
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
        
        # === Проверить массив на Н/Д ===
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
        
        # === Добавить колонку ЗОД ===
        if 'ЗОД' not in df_clean.columns:
            df_clean['ЗОД'] = ''
        
        return df_clean
    
    def add_zod_from_hierarchy(self, array_df, hierarchy_df=None):
        try:
            if array_df is None or array_df.empty:
                return array_df
            
            array_clean = array_df.copy()
            
            # Ищем колонку АСС (только точное совпадение)
            acc_col = None
            for col in array_clean.columns:
                if col == 'АСС' or col == 'ACC' or col == 'АСМ':
                    acc_col = col
                    break
            
            if not acc_col:
                return array_clean
            
            # Создаем колонку ЗОД если её нет
            if 'ЗОД' not in array_clean.columns:
                array_clean['ЗОД'] = ''
            
            # Заполняем ТОЛЬКО пустые значения ЗОД (НЕ стираем существующие)
            mask_empty = (array_clean['ЗОД'] == '') | (array_clean['ЗОД'].isna())
            
            if mask_empty.any():
                def get_zod_by_acc(acc_value):
                    if pd.isna(acc_value) or str(acc_value).strip().lower() in ['nan', 'none', 'null', '']:
                        return ''
                    clean_acc = str(acc_value).strip()
                    return ZOD_MAPPING.get(clean_acc, '')
                
                array_clean.loc[mask_empty, 'ЗОД'] = array_clean.loc[mask_empty, acc_col].apply(get_zod_by_acc)
            
            return array_clean
            
        except Exception as e:
            return array_df
        

    def export_array_to_excel(self, cleaned_array_df, filename="очищенный_массив"):
        """Создает Excel файл для очищенного массива"""
        try:
            if cleaned_array_df is None or cleaned_array_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                cleaned_array_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ МАССИВ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None
        
    def enrich_array_with_project_codes(self, cleaned_array_df, projects_df):
        """Обертка для кэшируемой версии"""
        return _enrich_array_with_project_codes_cached(cleaned_array_df, projects_df)

    def export_discrepancies_to_excel(self, discrepancy_df, filename="Расхождение_Массив"):
        try:
            if discrepancy_df is None or discrepancy_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                info_df = pd.DataFrame({
                    'Информация': [
                        'Файл создан автоматически',
                        f'Дата создания: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}',
                        f'Количество строк: {len(discrepancy_df)}',
                        'Эти строки не удалось обогатить кодами проектов'
                    ]
                })
                info_df.to_excel(writer, sheet_name='ИНФО', index=False)
                discrepancy_df.to_excel(writer, sheet_name='РАСХОЖДЕНИЯ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None
    
    def export_to_excel(self, df, cleaned_df, filename="очищенные_данные"):
        try:
            if df is None or cleaned_df is None:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='ОРИГИНАЛ', index=False)
                cleaned_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None

    def update_field_projects_flag(self, google_df):
        """
        Обновляет поле 'Полевой' в гугл таблице
        """
        try:
            google_df = google_df.copy()
            
            code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24'])
            if not code_col:
                return google_df
            
            google_df['Полевой'] = self._is_field_project_vectorized(google_df[code_col])
            
            return google_df
            
        except Exception as e:
            return google_df


    def add_field_flag_to_array(self, array_df):
        """
        Добавляет 'Полевой' в массив
        """
        try:
            array_df = array_df.copy()
            # ВСЕ ПРОЕКТЫ СТАНОВЯТСЯ ПОЛЕВЫМИ
            array_df['Полевой'] = 1
            
            # code_col = self._find_column(array_df, ['Код анкеты'])
            # if not code_col:
            #     return array_df
            
            # def is_field_project(code):
            #     try:
            #         if pd.isna(code):
            #             return 0
                        
            #         code_str = str(code).strip()
            #         lower_code = code_str.lower()
                    
            #         if any(word in lower_code for word in ['мультикод', 'пилот', 'семпл']):
            #             return 1
                    
            #         parts = code_str.split('.')
            #         if len(parts) >= 4:
            #             country = parts[0]
            #             if len(parts[2]) >= 2:
            #                 direction = '.' + parts[2][:2]
            #             else:
            #                 direction = ''
                        
            #             if country in ['RU00', 'RU01', 'RU02', 'RU03', 'RU04'] and direction in ['.01', '.02']:
            #                 return 1
                            
            #         return 0
            #     except:
            #         return 0
            
            # array_df['Полевой'] = self._is_field_project_vectorized(array_df[code_col])
            
            return array_df
            
        except Exception as e:
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
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            
            if not google_code_col or not google_portal_col:
                array_df['ПО'] = 'не определено'
                return array_df
            
            portal_mapping = {}
            cxway_mapping = {}
            for _, row in google_df.iterrows():
                code = str(row.get(google_code_col, '')).strip()
                portal = str(row.get(google_portal_col, '')).strip()
                
                if code and code.lower() not in ['nan', 'none', 'null', '']:
                    portal_mapping[code] = portal
                    if portal.upper() == 'CXWAY':
                        cxway_mapping[code] = 'CXWAY'
            
            def get_portal(code):
                if pd.isna(code):
                    return 'Чеккер'
                clean_code = str(code).strip()
                # Если код неуникальный — не переопределяем, оставляем Чеккер
                if self.is_non_unique_code(clean_code):
                    return 'Чеккер'
                # Переопределяем только если в Google проект отмечен как CXWAY
                if clean_code in cxway_mapping:
                    return 'CXWAY'
                return portal_mapping.get(clean_code, 'Чеккер')
            
            array_df['ПО'] = array_df[array_code_col].apply(get_portal)
            
            return array_df
            
        except Exception as e:
            array_df['ПО'] = 'не определено'
            return array_df
    
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
            
            # Маппинг стандартных колонок
            column_mapping = {
                'Код анкеты': ['Код анкеты'],
                'Имя клиента': ['Имя клиента'],
                'Название проекта': ['Название проекта',' Название проекта'],
                'ЗОД': ['ЗОД', 'ZOD', 'Зод', 'zod'],
                'АСС': ['АСС', 'ACC'],
                'ЭМ': ['ЭМ рег'],
                'Регион short': ['Регион'],
                'Регион': ['Регион '],
                'Полевой': ['Полевой'],
                'Статус': [' Статус', 'Статус'],
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
                           'ЗОД', 'АСС', 'ЭМ', 'Регион short', 'Регион', 'ПО', 'Полевой', 'Статус', 'Дата визита']
            
            # Реорганизуем колонки
            if not field_projects.empty:
                for col in final_columns:
                    if col not in field_projects.columns:
                        field_projects[col] = '' if col != 'Полевой' else 0
                field_projects = field_projects.reindex(columns=final_columns)
                field_projects['Источник'] = 'Портал'
            
            if not non_field_projects.empty:
                for col in final_columns:
                    if col not in non_field_projects.columns:
                        non_field_projects[col] = '' if col != 'Полевой' else 0
                non_field_projects = non_field_projects.reindex(columns=final_columns)
                non_field_projects['Источник'] = 'Портал'
            
            return field_projects, non_field_projects
            
        except Exception as e:
            return None, None
    
    def export_split_array_to_excel(self, field_df, non_field_df, filename="разделенный_массив"):
        """Создает Excel с вкладками Полевые/Неполевые проекты"""
        try:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if field_df is not None and not field_df.empty:
                    field_df.to_excel(writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет полевых проектов']}).to_excel(
                        writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
                
                if non_field_df is not None and not non_field_df.empty:
                    non_field_df.to_excel(writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет неполевых проектов']}).to_excel(
                        writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None
            
    def export_field_projects_only(self, field_df, filename="полевые_проекты"):
        """Создает Excel файл ТОЛЬКО с полевыми проектами"""
        try:
            if field_df is None or field_df.empty:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not field_df.empty:
                    field_df.to_excel(writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет полевых проектов']}).to_excel(
                        writer, sheet_name='ПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None

    def export_non_field_projects_only(self, non_field_df, filename="неполевые_проекты"):
        """Создает Excel файл ТОЛЬКО с неполевыми проектами"""
        try:
            if non_field_df is None or non_field_df.empty:
                return None
            
            non_field_clean = non_field_df.copy()
            
            required_cols = ['Код анкеты', 'Имя клиента', 'Название проекта']
            missing_cols = [col for col in required_cols if col not in non_field_clean.columns]
            
            if missing_cols:
                non_field_unique = non_field_clean
            else:
                non_field_unique = non_field_clean.drop_duplicates(
                    subset=required_cols, 
                    keep='first'
                )
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not non_field_unique.empty:
                    columns_to_drop = ['ЗОД', 'АСС', 'ЭМ', 'Регион short', 'Регион', 'Полевой']
                    columns_exist = [col for col in columns_to_drop if col in non_field_unique.columns]
                    
                    if columns_exist:
                        non_field_export = non_field_unique.drop(columns=columns_exist, errors='ignore')
                    else:
                        non_field_export = non_field_unique
                    
                    non_field_export.to_excel(writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False)
                else:
                    pd.DataFrame({'Сообщение': ['Нет неполевых проектов']}).to_excel(
                        writer, sheet_name='НЕПОЛЕВЫЕ_ПРОЕКТЫ', index=False
                    )
            
            output.seek(0)
            return output
            
        except Exception as e:
            return None
            
    def check_problematic_projects(self, google_df, field_df):
        """
        Проверка проблемных проектов после очистки и обогащения
        
        Сравнивает данные из Google-таблицы (план) с данными визитов
        и находит несоответствия.
        """
        
        # ============================================================
        # ЭТАП 1: ПОДГОТОВКА ДАННЫХ
        # ============================================================
        
        if google_df is None or google_df.empty:
            return pd.DataFrame()
        
        # Шаг 1.1: Поиск колонок в Google-таблице
        google_cols = {
            'name': self._find_column(google_df, ['Проекты в  https://ru.checker-soft.com', 'Проекты']),
            'wave': self._find_column(google_df, ['Название волны на Чекере/ином ПО', 'Волна']),
            'code': self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта']),
            'portal': self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО']),
            'fio_om': self._find_column(google_df, ['ФИО ОМ'])  # 🔴 НОВАЯ КОЛОНКА
        }
        
        if not all([google_cols['name'], google_cols['wave'], google_cols['code']]):
            return pd.DataFrame()
        
        # Шаг 1.2: Собираем все коды из данных
        all_codes = set()
        if field_df is not None and not field_df.empty:
            field_code_col = self._find_column(field_df, ['Код анкеты', 'Код'])
            if field_code_col:
                codes = field_df[field_code_col].astype(str).str.strip().fillna('').values
                all_codes = {c for c in codes if c and c not in ['', 'nan', 'None', 'null']}
        
        # ============================================================
        # ЭТАП 2: СБОР ПРОБЛЕМ
        # ============================================================
        
        problems = []
        
        # ------------------------------------------------------------
        # ПРОВЕРКА 1: Проекты из Google, отсутствующие в данных
        # ------------------------------------------------------------
        
        for _, row in google_df.iterrows():
            # Извлечение данных
            name = str(row.get(google_cols['name'], '')).strip()
            wave = str(row.get(google_cols['wave'], '')).strip()
            code = str(row.get(google_cols['code'], '')).strip()
            portal = str(row.get(google_cols['portal'], '')).strip() if google_cols['portal'] else ''
            fio_om = str(row.get(google_cols['fio_om'], '')).strip() if google_cols['fio_om'] else ''  # 🔴 НОВОЕ
            is_field = row.get('Полевой', 0)
            
            if not name or name in ['nan', 'None', '']:
                continue
            
            # Проверка проблем
            code_empty = pd.isna(row.get(google_cols['code'])) or code in ['', 'nan', 'None', 'null']
            non_field = (is_field == 0)
            missing_in_data = (is_field == 1 and code and code not in ['', 'nan', 'None', 'null'] 
                              and code not in all_codes)
            
            if code_empty or non_field or missing_in_data:
                problems.append({
                    'source': 'google',
                    'name': name, 'wave': wave, 'code': code if not code_empty else '',
                    'portal': portal, 'fio_om': fio_om,  # 🔴 НОВОЕ
                    'code_empty': code_empty,
                    'non_field': non_field, 'missing_in_data': missing_in_data
                })
        
        # ------------------------------------------------------------
        # ПРОВЕРКА 2: Проекты из данных, отсутствующие в Google
        # ------------------------------------------------------------
        
        if field_df is not None and not field_df.empty:
            # Поиск колонок в данных
            field_cols = {
                'name': self._find_column(field_df, ['Имя клиента', 'Client']),
                'wave': self._find_column(field_df, ['Название проекта', 'Wave Name']),
                'code': self._find_column(field_df, ['Код анкеты', 'Код'])
            }
            
            if all(field_cols.values()):
                # Коды из Google для сравнения
                google_codes = {str(c).strip() for c in google_df[google_cols['code']].values 
                              if pd.notna(c) and str(c).strip() not in ['', 'nan', 'None', 'null']}
                
                for _, row in field_df.iterrows():
                    name = str(row.get(field_cols['name'], '')).strip()
                    wave = str(row.get(field_cols['wave'], '')).strip()
                    code = str(row.get(field_cols['code'], '')).strip()
                    
                    if not name or name in ['nan', 'None', '']:
                        continue
                    
                    if code and code not in ['', 'nan', 'None', 'null'] and code not in google_codes:
                        problems.append({
                            'source': 'data',
                            'name': name, 'wave': wave, 'code': code,
                            'portal': '', 'fio_om': '',  # 🔴 НОВОЕ (пустое значение)
                            'code_empty': False,
                            'non_field': False, 'missing_in_data': False,
                            'missing_in_google': True
                        })
        
        # ============================================================
        # ЭТАП 3: ФОРМИРОВАНИЕ РЕЗУЛЬТАТА
        # ============================================================
        
        if not problems:
            return pd.DataFrame()
        
        # Преобразование в итоговый формат
        result = []
        for p in problems:
            result.append({
                'Название проекта': p['name'],
                'Волна': p['wave'],
                'Код проекта': p['code'],
                'ПО': p['portal'],
                'ФИО ОМ': p.get('fio_om', ''),  # 🔴 НОВАЯ КОЛОНКА
                'Код проекта пусто': p.get('code_empty', False),
                'Проект неполевой, есть в гугл': p.get('non_field', False),
                'Проект есть в гугл, нет в массиве': p.get('missing_in_data', False),
                'Проект есть в массиве, нет в гугл': p.get('missing_in_google', False)
            })
        
        df_result = pd.DataFrame(result)
        df_result = df_result.drop_duplicates(subset=['Название проекта', 'Волна'])
        df_result = df_result.sort_values('Название проекта')
        
        return df_result
    
    def clean_cxway(self, df, hierarchy_df, google_df):
        """Очистка файла CXWAY и приведение к структуре полевых проектов"""
        if df is None or df.empty:
            return pd.DataFrame()
        
        self._log_samples(df, "1. Исходные данные")
        df_clean = df.copy()
        
        # Удалить строки где Status == "Удалено"
        status_col = self._find_column(df_clean, ['Status', 'Статус', 'status'])
        if status_col:
            df_clean = df_clean[df_clean[status_col].astype(str).str.strip() != 'Удалено']
       
        self._log_samples(df_clean, "2. После удаления статуса")
        
        # Удалить строки где Date of Visit < первый день месяца
        date_col = self._find_column(df_clean, ['Date of Visit', 'Дата визита', 'Visit Date'])
        if date_col:
            first_day = None
            if 'plan_calc_params' in st.session_state:
                first_day = pd.Timestamp(st.session_state['plan_calc_params']['start_date'])
            else:
                today = datetime.now()
                first_day = pd.Timestamp(year=today.year, month=today.month, day=1)
            
            df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors='coerce')
            mask = pd.isna(df_clean[date_col]) | (df_clean[date_col] >= first_day)
            df_clean = df_clean[mask]
        
        self._log_samples(df_clean, "3. После фильтра дат")
        
        # Базовые колонки
        column_mapping = {
            'Код анкеты': ['Project Code', 'Код', 'Code'],
            'Имя клиента': ['Client', 'Имя клиента', 'Клиент имя'],
            'Название проекта': ['Wave Name'],
            'АСС': ['АСС', 'ACC'],
            'ЭМ': ['ЭМ рег', 'Эксперт', 'Эксперт менеджер'],
            'Регион short': ['Region short'],
            'Статус': ['Status'],
            'Дата визита': ['Date of Visit']
        }    
        
        result = pd.DataFrame()
        
        for std_col, possible_names in column_mapping.items():
            source_col = self._find_column(df_clean, possible_names)
            if source_col:
                result[std_col] = df_clean[source_col].astype(str).fillna('')
            else:
                result[std_col] = ''
        
        # Обогащение кодов проектов из гугл-таблицы
        if google_df is not None and not google_df.empty:
            empty_code_mask = (
                result['Код анкеты'].isna() | 
                (result['Код анкеты'].astype(str).str.strip() == '') |
                (result['Код анкеты'].astype(str).str.strip() == 'nan')
            )
            rows_to_process = result[empty_code_mask]
            total_empty = len(rows_to_process)
            
            if total_empty > 0:
                google_df = google_df.copy()
                project_client_col = self._find_column(google_df, ['Проекты в  https://ru.checker-soft.com'])
                project_wave_col = self._find_column(google_df, ['Название волны на Чекере/ином ПО'])
                project_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24'])
                
                if all([project_client_col, project_wave_col, project_code_col]):
                    google_df['_match_client'] = google_df[project_client_col].astype(str).str.strip()
                    google_df['_match_wave'] = google_df[project_wave_col].astype(str).str.strip()
                    
                    filled_count = 0
                    for idx, row in rows_to_process.iterrows():
                        client_name = str(row['Имя клиента']).strip() if pd.notna(row['Имя клиента']) else ''
                        wave_name = str(row['Название проекта']).strip() if pd.notna(row['Название проекта']) else ''
                        
                        match_mask = (
                            (google_df['_match_client'] == client_name) &
                            (google_df['_match_wave'] == wave_name)
                        )
                        
                        matched_rows = google_df[match_mask]
                        
                        if not matched_rows.empty:
                            project_code = matched_rows.iloc[0][project_code_col]
                            if pd.notna(project_code) and str(project_code).strip() != '':
                                result.at[idx, 'Код анкеты'] = str(project_code).strip()
                                filled_count += 1
        
        self._log_samples(result, "4. После обогащения кодами")
        
        # все CXWAY проекты полевые
        result['Полевой'] = 1
        
        # # Определение "Полевой" (векторно)
        # if 'Код анкеты' in result.columns and not result.empty:
        #     result['Полевой'] = self._is_field_project_vectorized(result['Код анкеты'])

        # self._log_samples(result, "5. После определения полевого")
        
        # Добавление ЗОД из встроенного справочника (по АСС)
        if 'АСС' in result.columns:
            def get_zod_by_acc(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip() == '':
                    return ''
                clean_acc = str(acc_value).strip()
                return ZOD_MAPPING.get(clean_acc, '')
            
            result['ЗОД'] = result['АСС'].apply(get_zod_by_acc)
        else:
            result['ЗОД'] = ''
        
        # Добавление ПО
        result['ПО'] = 'CXWAY'  # значение по умолчанию
        
        # Если есть Google, проверяем только на 'Чеккер'
        if google_df is not None and 'Код анкеты' in result.columns:
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            
            if google_code_col and google_portal_col:
                # Создаем словарь только для проектов с ПО 'Чеккер'
                checker_mapping = {}
                for _, row in google_df.iterrows():
                    code = str(row.get(google_code_col, '')).strip()
                    portal = str(row.get(google_portal_col, '')).strip()
                    if code and portal.upper() == 'ЧЕККЕР':
                        checker_mapping[code] = 'Чеккер'
                
                def get_portal_from_google(code_value):
                    if pd.isna(code_value) or str(code_value).strip() == '':
                        return 'CXWAY'
                    clean_code = str(code_value).strip()
                    # Если код неуникальный — не переопределяем, оставляем CXWAY
                    if self.is_non_unique_code(clean_code):
                        return 'CXWAY'
                    # Переопределяем только если в Google проект отмечен как Чеккер
                    if clean_code in checker_mapping:
                        return 'Чеккер'
                    return 'CXWAY'
                
                result['ПО'] = result['Код анкеты'].apply(get_portal_from_google)
        
        # Удаляем проекты, которые в Google отмечены как Чеккер
        if google_df is not None and not google_df.empty:
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            google_client_col = self._find_column(google_df, ['Проекты в  https://ru.checker-soft.com', 'Проекты'])
            
            if google_code_col and google_portal_col and google_client_col:
                # Находим проекты с ПО Чеккер
                checker_mask = google_df[google_portal_col].astype(str).str.strip().str.upper() == 'ЧЕККЕР'
                checker_df = google_df[checker_mask].copy()
                
                if not checker_df.empty:
                    # Создаем ключи (клиент + код) для проектов Чеккер
                    checker_df['_key'] = (
                        checker_df[google_client_col].astype(str).str.strip() + '|' +
                        checker_df[google_code_col].astype(str).str.strip()
                    )
                    
                    # Создаем ключи в CXWAY (клиент + код)
                    result['_key'] = (
                        result['Имя клиента'].astype(str).str.strip() + '|' +
                        result['Код анкеты'].astype(str).str.strip()
                    )
                    
                    # Определяем семплы и пилоты
                    result['_is_sample_pilot'] = result['Код анкеты'].astype(str).str.contains('семпл|пилот', case=False, na=False)
                    
                    # Удаляем ТОЛЬКО обычные проекты (не семплы/пилоты)
                    mask_to_remove = result['_key'].isin(checker_df['_key']) & ~result['_is_sample_pilot']
                    result = result[~mask_to_remove]
                    
                    # Удаляем временные колонки
                    result = result.drop(['_key', '_is_sample_pilot'], axis=1)
        
        # Добавление полного региона
        region_mapping = {
            'AD': 'Республика Адыгея', 'AL': 'Алтайский край', 'AM': 'Амурская область',
            'AR': 'Архангельская область', 'AS': 'Астраханская область', 'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область', 'BR': 'Брянская область', 'BU': 'Республика Бурятия',
            'CK': 'Чукотский автономный округ', 'CL': 'Челябинская область', 'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика', 'DA': 'Республика Дагестан', 'DN': 'Донецкая Народная Республика',
            'GA': 'Республика Алтай', 'IN': 'Республика Ингушетия', 'IR': 'Иркутская область',
            'IV': 'Ивановская область', 'KA': 'Камчатский край', 'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика', 'KD': 'Краснодарский край', 'KE': 'Кемеровская область',
            'KG': 'Калужская область', 'KH': 'Хабаровский край', 'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия', 'KL': 'Республика Калмыкия', 'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область', 'KO': 'Республика Коми', 'KS': 'Курская область',
            'KT': 'Костромская область', 'KU': 'Курганская область', 'KV': 'Кировская область',
            'KY': 'Красноярский край', 'LG': 'Луганская Народная Республика', 'LN': 'Ленинградская область',
            'LP': 'Липецкая область', 'MC': 'Московская область', 'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область', 'MM': 'Мурманская область', 'MR': 'Республика Мордовия',
            'MS': 'Московская область', 'NG': 'Новгородская область', 'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия', 'NS': 'Новосибирская область', 'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область', 'OL': 'Орловская область', 'OM': 'Омская область',
            'PE': 'Пермский край', 'PR': 'Приморский край', 'PS': 'Псковская область',
            'PZ': 'Пензенская область', 'RK': 'Республика Крым', 'RO': 'Ростовская область',
            'RZ': 'Рязанская область', 'SA': 'Самарская область', 'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область', 'SM': 'Смоленская область', 'SR': 'Саратовская область',
            'ST': 'Ставропольский край', 'SV': 'Свердловская область', 'TB': 'Тамбовская область',
            'TL': 'Тульская область', 'TO': 'Томская область', 'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва', 'TV': 'Тверская область', 'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика', 'UL': 'Ульяновская область', 'VG': 'Волгоградская область',
            'VL': 'Владимирская область', 'VO': 'Вологодская область', 'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ', 'YS': 'Ярославская область', 'YV': 'Еврейская автономная область',
            'ZK': 'Забайкальский край', 'ZO': 'Запорожская область'
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
        
        result['Источник'] = 'CXWAY'
        
        return result
    
    def clean_easymerch(self, df, google_df):
        """
        Очистка файла Easymerch и приведение к структуре полевых проектов
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        df_clean = df.copy()
        
        # Маппинг колонок Easymerch → стандартные названия
        column_mapping = {
            'Код анкеты': ['код проекта'],
            'Имя клиента': ['Имя клиента'],
            'Название проекта': ['Название волны'],
            'АСС': ['АСМ'],
            'ЭМ': ['ЭМ', 'ТП'],
            'Регион short': ['Регион'],
            'Статус': ['Статус'],
            'Дата визита': ['Дата визита']
        }
        
        result = pd.DataFrame()
        
        for std_col, possible_names in column_mapping.items():
            source_col = self._find_column(df_clean, possible_names)
            if source_col:
                result[std_col] = df_clean[source_col].astype(str).fillna('')
            else:
                result[std_col] = ''
        
        # Обработка региона (берем первые 2 символа)
        if 'Регион short' in result.columns:
            result['Регион short'] = result['Регион short'].str[:2]
        
        # Все записи Easymerch - полевые
        result['Полевой'] = 1
        
        # Добавление ПО из гугл таблицы по коду проекта
        result['ПО'] = 'не определено'  # значение по умолчанию
        
        if google_df is not None and 'Код анкеты' in result.columns:
            # Находим колонки в гугл таблице
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            
            if google_code_col and google_portal_col:
                # Создаем словарь {код проекта: ПО}
                portal_mapping = {}
                for _, row in google_df.iterrows():
                    code = str(row.get(google_code_col, '')).strip()
                    portal = str(row.get(google_portal_col, '')).strip()
                    if code and code.lower() not in ['nan', 'none', 'null', '']:
                        portal_mapping[code] = portal
                
                def get_portal(code_value):
                    if pd.isna(code_value) or str(code_value).strip() == '':
                        return 'Easymerch'
                    clean_code = str(code_value).strip()
                    # Если код неуникальный — не переопределяем
                    if self.is_non_unique_code(clean_code):
                        return 'Easymerch'
                    return portal_mapping.get(clean_code, 'Easymerch')
                
                result['ПО'] = result['Код анкеты'].apply(get_portal)
        
        # Добавление полного региона (из справочника)
        region_mapping = {
            'AD': 'Республика Адыгея', 'AL': 'Алтайский край', 'AM': 'Амурская область',
            'AR': 'Архангельская область', 'AS': 'Астраханская область', 'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область', 'BR': 'Брянская область', 'BU': 'Республика Бурятия',
            'CK': 'Чукотский автономный округ', 'CL': 'Челябинская область', 'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика', 'DA': 'Республика Дагестан', 'DN': 'Донецкая Народная Республика',
            'GA': 'Республика Алтай', 'IN': 'Республика Ингушетия', 'IR': 'Иркутская область',
            'IV': 'Ивановская область', 'KA': 'Камчатский край', 'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика', 'KD': 'Краснодарский край', 'KE': 'Кемеровская область',
            'KG': 'Калужская область', 'KH': 'Хабаровский край', 'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия', 'KL': 'Республика Калмыкия', 'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область', 'KO': 'Республика Коми', 'KS': 'Курская область',
            'KT': 'Костромская область', 'KU': 'Курганская область', 'KV': 'Кировская область',
            'KY': 'Красноярский край', 'LG': 'Луганская Народная Республика', 'LN': 'Ленинградская область',
            'LP': 'Липецкая область', 'MC': 'Московская область', 'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область', 'MM': 'Мурманская область', 'MR': 'Республика Мордовия',
            'MS': 'Московская область', 'NG': 'Новгородская область', 'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия', 'NS': 'Новосибирская область', 'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область', 'OL': 'Орловская область', 'OM': 'Омская область',
            'PE': 'Пермский край', 'PR': 'Приморский край', 'PS': 'Псковская область',
            'PZ': 'Пензенская область', 'RK': 'Республика Крым', 'RO': 'Ростовская область',
            'RZ': 'Рязанская область', 'SA': 'Самарская область', 'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область', 'SM': 'Смоленская область', 'SR': 'Саратовская область',
            'ST': 'Ставропольский край', 'SV': 'Свердловская область', 'TB': 'Тамбовская область',
            'TL': 'Тульская область', 'TO': 'Томская область', 'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва', 'TV': 'Тверская область', 'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика', 'UL': 'Ульяновская область', 'VG': 'Волгоградская область',
            'VL': 'Владимирская область', 'VO': 'Вологодская область', 'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ', 'YS': 'Ярославская область', 'YV': 'Еврейская автономная область',
            'ZK': 'Забайкальский край', 'ZO': 'Запорожская область'
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
        
        
        # Добавление ЗОД из встроенного справочника (по АСС) с учетом новых сотрудников
        if 'АСС' in result.columns:            
            def get_zod(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip() == '':
                    return ''
                clean_acc = str(acc_value).strip()
                return ZOD_MAPPING.get(clean_acc, '')
            
            result['ЗОД'] = result['АСС'].apply(get_zod)
        else:
            result['ЗОД'] = ''
        
        # Добавляем источник
        result['Источник'] = 'Easymerch'
        
        return result

    def clean_optima(self, df, google_df):
        """
        Очистка файла Optima и приведение к структуре полевых проектов
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        df_clean = df.copy()
        
        # Словарь для преобразования длинных названий регионов в короткие коды
        long_to_short_region = {
            'Республика Адыгея': 'AD', 'Алтайский край': 'AL','Алтай': 'AL', 'Амурская область': 'AM',
            'Архангельская область': 'AR', 'Астраханская область': 'AS', 'Республика Башкортостан': 'BK',
            'Белгородская область': 'BL', 'Брянская область': 'BR', 'Республика Бурятия': 'BU',
            'Чукотский автономный округ': 'CK', 'Челябинская область': 'CL', 'Чеченская Республика': 'CN',
            'Чувашская Республика': 'CV', 'Республика Дагестан': 'DA', 'Донецкая Народная Республика': 'DN',
            'Республика Алтай': 'GA', 'Республика Ингушетия': 'IN', 'Иркутская область': 'IR',
            'Ивановская область': 'IV', 'Камчатский край': 'KA', 'Кабардино-Балкарская Республика': 'KB',
            'Карачаево-Черкесская Республика': 'KC', 'Краснодарский край': 'KD', 'Кемеровская область': 'KE',
            'Калужская область': 'KG', 'Хабаровский край': 'KH', 'Республика Карелия': 'KI',
            'Республика Хакасия': 'KK', 'Республика Калмыкия': 'KL', 'Ханты-Мансийский автономный округ': 'KM',
            'Калининградская область': 'KN', 'Республика Коми': 'KO', 'Курская область': 'KS',
            'Костромская область': 'KT', 'Курганская область': 'KU', 'Кировская область': 'KV',
            'Красноярский край': 'KY', 'Луганская Народная Республика': 'LG', 'Ленинградская область': 'LN',
            'Липецкая область': 'LP', 'Московская область': 'MS','Москва': 'MC', 'Республика Марий Эл': 'ME',
            'Магаданская область': 'MG', 'Мурманская область': 'MM', 'Республика Мордовия': 'MR',
            'Новгородская область': 'NG', 'Ненецкий автономный округ': 'NN', 'Республика Северная Осетия': 'NO',
            'Новосибирская область': 'NS', 'Нижегородская область': 'NZ', 'Оренбургская область': 'OB',
            'Орловская область': 'OL', 'Омская область': 'OM', 'Пермский край': 'PE',
            'Приморский край': 'PR', 'Псковская область': 'PS', 'Пензенская область': 'PZ',
            'Республика Крым': 'RK', 'Ростовская область': 'RO', 'Рязанская область': 'RZ',
            'Самарская область': 'SA', 'Республика Саха (Якутия)': 'SK', 'Сахалинская область': 'SL',
            'Смоленская область': 'SM', 'Саратовская область': 'SR', 'Ставропольский край': 'ST',
            'Свердловская область': 'SV', 'Тамбовская область': 'TB', 'Тульская область': 'TL',
            'Томская область': 'TO', 'Республика Татарстан': 'TT', 'Республика Тыва': 'TU',
            'Тверская область': 'TV', 'Тюменская область': 'TY', 'Удмуртская Республика': 'UD',
            'Ульяновская область': 'UL', 'Волгоградская область': 'VG', 'Владимирская область': 'VL',
            'Вологодская область': 'VO', 'Воронежская область': 'VR', 'Ямало-Ненецкий автономный округ': 'YN',
            'Ярославская область': 'YS', 'Еврейская автономная область': 'YV', 'Забайкальский край': 'ZK',
            'Запорожская область': 'ZO'
        }
        
        # Маппинг колонок Optima → стандартные названия
        column_mapping = {
            'Код анкеты': ['Project Code'],
            'Имя клиента': ['Проект'],
            'Название проекта': ['Wave Name'],
            'АСС': ['АСС'],
            'ЭМ': ['Координатор'],
            'Регион long': ['Регион Чекер'],
            'Статус': ['Статус анкеты'],
            'Дата визита': ['Дата визита']
        }
        
        result = pd.DataFrame()
        
        for std_col, possible_names in column_mapping.items():
            source_col = self._find_column(df_clean, possible_names)
            if source_col:
                if std_col == 'Код анкеты':
                    # Обработка Project Code (может содержать обратный слеш)
                    codes = df_clean[source_col].astype(str).fillna('')
                    result[std_col] = codes.apply(
                        lambda x: x.split('\\')[0].strip() if '\\' in x else x.strip()
                    )
                elif std_col == 'Регион long':
                    result['Регион long'] = df_clean[source_col].astype(str).fillna('')
                else:
                    result[std_col] = df_clean[source_col].astype(str).fillna('')
            else:
                if std_col == 'Регион long':
                    result['Регион long'] = ''
                else:
                    result[std_col] = ''
        
        # Конвертация даты
        if 'Дата визита' in result.columns:
            result['Дата визита'] = pd.to_datetime(
                result['Дата визита'], 
                errors='coerce',
                dayfirst=True
            )
        
        # Преобразуем длинные названия регионов в короткие коды
        def get_short_region(long_name):
            if pd.isna(long_name) or str(long_name).strip() in ['', 'nan', 'none', 'null']:
                return 'не определен'
            clean_name = str(long_name).strip()
            # Прямое соответствие
            if clean_name in long_to_short_region:
                return long_to_short_region[clean_name]
            # Поиск по вхождению
            for long, short in long_to_short_region.items():
                if long in clean_name or clean_name in long:
                    return short
            return 'не определен'
        
        result['Регион short'] = result['Регион long'].apply(get_short_region)
        result['Регион'] = result['Регион long']  # сохраняем полное название для отчетов
        
        # Все записи Optima - полевые
        result['Полевой'] = 1
        
        # Источник
        result['Источник'] = 'Оптима'
        
        # ПО (портал) - по умолчанию Optima
        result['ПО'] = 'Оптима'
        

        
        # Добавление ЗОД из встроенного справочника (по АСС)
        if 'АСС' in result.columns:
            
            def get_zod(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip() in ['', 'nan', 'none', 'null']:
                    return ''
                clean_acc = str(acc_value).strip()
                return ZOD_MAPPING.get(clean_acc, '')
            
            result['ЗОД'] = result['АСС'].apply(get_zod)
        else:
            result['ЗОД'] = ''
        
        result['Источник'] = 'Оптима'
        
        return result
    
    def clean_prodata(self, df, google_df):
        """
        Очистка файла ПроДата (Мониторинги) и приведение к структуре полевых проектов
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        df_clean = df.copy()
        
        # Маппинг колонок ПроДата → стандартные названия
        column_mapping = {
            'Код анкеты': ['Код проекта'],
            'Имя клиента': ['Имя клиента'],
            'Название проекта': ['Направление'],
            'Тип мониторинга': ['Тип мониторинга'],
            'АСС': ['АСС'],
            'Регион short': ['Кластер'],
            'Статус': ['Статус'],
            'Дата визита': ['Дата визита']
        }
        
        result = pd.DataFrame()
        
        for std_col, possible_names in column_mapping.items():
            source_col = self._find_column(df_clean, possible_names)
            if source_col:
                result[std_col] = df_clean[source_col].astype(str).fillna('')
            else:
                result[std_col] = ''
                
        # ЗАМЕНЯЕМ ВСЕ ЗНАЧЕНИЯ ВОЛНЫ НА "Все волны"
        if 'Название проекта' in result.columns:
            result['Название проекта'] = 'Все волны'
        
        # ЭМ всегда пусто (нет в исходных данных)
        result['ЭМ'] = ''
        
        # --- КОНВЕРТАЦИЯ ДАТЫ (важно!) ---
        if 'Дата визита' in result.columns:
            result['Дата визита'] = pd.to_datetime(
                result['Дата визита'], 
                errors='coerce',
                dayfirst=True
            )

        # Заменяем фейковые даты (1900-01-01) на NaT
        fake_date = pd.Timestamp('1900-01-01')
        result['Дата визита'] = result['Дата визита'].replace(fake_date, pd.NaT)
        
        # --- СПЕЦИАЛЬНАЯ ОБРАБОТКА РЕГИОНА ---
        # Словарь для поиска региона по ключевым словам
        # (перенесен из DataVisualizer для использования в DataCleaner)
        region_keywords = {
            'AD': ['адыг'], 'AL': ['алтай'], 'AM': ['амур'],
            'AR': ['архангельск'], 'AS': ['астрахан'], 'BK': ['башкортостан', 'башкир'],
            'BL': ['белгород'], 'BR': ['брянск'], 'BU': ['бурят'],
            'CL': ['челябин'], 'CN': ['чечен'], 'CV': ['чуваш'],
            'DA': ['дагестан'], 'DN': ['донецк', 'донецкая'], 'GA': ['горный алтай', 'республика алтай'],
            'IN': ['ингуш'], 'IR': ['иркут'], 'IV': ['иван'],
            'KA': ['камчат'], 'KB': ['кабард'], 'KC': ['карача'],
            'KD': ['краснодар'], 'KE': ['кемер'], 'KG': ['калуж'],
            'KH': ['хабаров'], 'KI': ['карел'], 'KK': ['хакас'],
            'KL': ['калмы'], 'KM': ['хант', 'манс'], 'KN': ['калинин'],
            'KO': ['коми'], 'KS': ['курск'], 'KT': ['костр'],
            'KU': ['курган'], 'KV': ['киров'], 'KY': ['краснояр'],
            'LG': ['луганск', 'луганская'], 'LN': ['ленинград', 'питер'], 'LP': ['липец'],
            'MC': ['моск'], 'ME': ['марий'], 'MG': ['магадан'],
            'MM': ['мурман'], 'MR': ['мордов'], 'MS': ['моск'],
            'NG': ['новгород'], 'NN': ['ненец'], 'NO': ['осет'],
            'NS': ['новосиб'], 'NZ': ['нижегород'], 'OB': ['оренбург'],
            'OL': ['орлов'], 'OM': ['омск'], 'PE': ['перм'],
            'PR': ['примор'], 'PS': ['псков'], 'PZ': ['пенз'],
            'RK': ['крым'], 'RO': ['ростов'], 'RZ': ['ряз'],
            'SA': ['самар'], 'SK': ['саха', 'якут'], 'SL': ['сахалин'],
            'SM': ['смол'], 'SR': ['саратов'], 'ST': ['ставроп'],
            'SV': ['свердлов'], 'TB': ['тамбов'], 'TL': ['туль'],
            'TO': ['томск'], 'TT': ['татар'], 'TU': ['тыва'],
            'TV': ['твер'], 'TY': ['тюмен'], 'UD': ['удмурт'],
            'UL': ['ульян'], 'VG': ['волгоград'], 'VL': ['владимир'],
            'VO': ['волог'], 'VR': ['воронеж'], 'YN': ['ямал'],
            'YS': ['ярослав'], 'YV': ['еврей'], 'ZK': ['забайкал'],
            'ZO': ['запорожье', 'запорожская']
        }
        
        def extract_region_code(cluster_value):
            """
            Из значения типа "Архангельск_ГМ_1" извлекает первую часть до "_"
            и находит соответствующий код региона по ключевым словам
            """
            if pd.isna(cluster_value) or str(cluster_value).strip() in ['', 'nan', 'none', 'null']:
                return 'не определен'
            
            cluster_str = str(cluster_value).strip()
            
            # Берем первую часть до разделителя "_"
            first_part = cluster_str.split('_')[0].strip().lower()
            
            if not first_part:
                return 'не определен'
            
            # Ищем по ключевым словам
            for region_code, keywords in region_keywords.items():
                for keyword in keywords:
                    if keyword in first_part:
                        return region_code
            
            return 'не определен'
        
        # Применяем обработку региона
        result['Регион short'] = result['Регион short'].apply(extract_region_code)
        
        # Все записи ПроДата - полевые
        result['Полевой'] = 1
        
        # ПО (портал) - по умолчанию "Мониторинги"
        result['ПО'] = 'Мониторинги'
        
        # Обогащение ПО из Google-таблицы по коду проекта
        if google_df is not None and 'Код анкеты' in result.columns:
            google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
            google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
            
            if google_code_col and google_portal_col:
                portal_mapping = {}
                for _, row in google_df.iterrows():
                    code = str(row.get(google_code_col, '')).strip()
                    portal = str(row.get(google_portal_col, '')).strip()
                    if code and code.lower() not in ['nan', 'none', 'null', '']:
                        portal_mapping[code] = portal
                
                def get_portal(code_value):
                    if pd.isna(code_value) or str(code_value).strip() == '':
                        return 'Мониторинги'
                    clean_code = str(code_value).strip()
                    return portal_mapping.get(clean_code, 'Мониторинги')
                
                result['ПО'] = result['Код анкеты'].apply(get_portal)
        
        # Добавление полного региона (из словаря - нужно добавить в класс)
        # Временно используем упрощенный вариант
        full_region_names = {
            'AD': 'Республика Адыгея', 'AL': 'Алтайский край', 'AM': 'Амурская область',
            'AR': 'Архангельская область', 'AS': 'Астраханская область', 'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область', 'BR': 'Брянская область', 'BU': 'Республика Бурятия',
            'CK': 'Чукотский автономный округ', 'CL': 'Челябинская область', 'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика', 'DA': 'Республика Дагестан', 'DN': 'Донецкая Народная Республика',
            'GA': 'Республика Алтай', 'IN': 'Республика Ингушетия', 'IR': 'Иркутская область',
            'IV': 'Ивановская область', 'KA': 'Камчатский край', 'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика', 'KD': 'Краснодарский край', 'KE': 'Кемеровская область',
            'KG': 'Калужская область', 'KH': 'Хабаровский край', 'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия', 'KL': 'Республика Калмыкия', 'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область', 'KO': 'Республика Коми', 'KS': 'Курская область',
            'KT': 'Костромская область', 'KU': 'Курганская область', 'KV': 'Кировская область',
            'KY': 'Красноярский край', 'LG': 'Луганская Народная Республика', 'LN': 'Ленинградская область',
            'LP': 'Липецкая область', 'MC': 'Московская область', 'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область', 'MM': 'Мурманская область', 'MR': 'Республика Мордовия',
            'MS': 'Московская область', 'NG': 'Новгородская область', 'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия', 'NS': 'Новосибирская область', 'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область', 'OL': 'Орловская область', 'OM': 'Омская область',
            'PE': 'Пермский край', 'PR': 'Приморский край', 'PS': 'Псковская область',
            'PZ': 'Пензенская область', 'RK': 'Республика Крым', 'RO': 'Ростовская область',
            'RZ': 'Рязанская область', 'SA': 'Самарская область', 'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область', 'SM': 'Смоленская область', 'SR': 'Саратовская область',
            'ST': 'Ставропольский край', 'SV': 'Свердловская область', 'TB': 'Тамбовская область',
            'TL': 'Тульская область', 'TO': 'Томская область', 'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва', 'TV': 'Тверская область', 'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика', 'UL': 'Ульяновская область', 'VG': 'Волгоградская область',
            'VL': 'Владимирская область', 'VO': 'Вологодская область', 'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ', 'YS': 'Ярославская область', 'YV': 'Еврейская автономная область',
            'ZK': 'Забайкальский край', 'ZO': 'Запорожская область'
        }
        
        def get_full_region(short):
            if pd.isna(short) or str(short).strip() == '' or short == 'не определен':
                return 'не определен'
            short_clean = str(short).strip().upper()
            return full_region_names.get(short_clean, 'не определен')
        
        result['Регион'] = result['Регион short'].apply(get_full_region)
        
        # Добавление ЗОД из встроенного справочника (по АСС)
        if 'АСС' in result.columns:
            
            def get_zod(acc_value):
                if pd.isna(acc_value) or str(acc_value).strip() in ['', 'nan', 'none', 'null']:
                    return ''
                clean_acc = str(acc_value).strip()
                return ZOD_MAPPING.get(clean_acc, '')
            
            result['ЗОД'] = result['АСС'].apply(get_zod)
        else:
            result['ЗОД'] = ''
        
        # Добавляем источник
        result['Источник'] = 'Мониторинги'
        
        # ===== НОВАЯ ЛОГИКА ДЛЯ ПРОДАТА =====
        if not result.empty:
            # Группируем по клиенту и ТИПУ МОНИТОРИНГА
            prodata_agg = result.groupby(['Имя клиента', 'Тип мониторинга']).size().reset_index(name='count')
            
            # Факт = количество / 70
            prodata_agg['Факт проекта, шт.'] = (prodata_agg['count'] / 70).round(1)
            
            # План = Факт
            prodata_agg['План проекта, шт.'] = prodata_agg['Факт проекта, шт.']
            prodata_agg['План на дату, шт.'] = prodata_agg['План проекта, шт.']
            prodata_agg['Факт на дату, шт.'] = prodata_agg['Факт проекта, шт.']
            
            # Переименовываем для отображения
            result = prodata_agg.rename(columns={
                'Имя клиента': 'Клиент'
            })
            
            # Добавляем служебные колонки
            result['ПО'] = 'Мониторинги'
            result['Полевой'] = 1
            result['Источник'] = 'ПроДата'
            result['ЗОД'] = ''
            result['АСС'] = ''
            result['ЭМ'] = ''
            result['Регион short'] = ''
            result['Регион'] = ''
            result['Статус'] = ''
            result['Дата визита'] = pd.NaT
        
        return result
    
    def _is_field_project_vectorized(self, codes_series):
        """Векторизованное определение полевых проектов (быстро)"""
        codes_str = codes_series.astype(str).str.strip()
        
        # Ключевые слова: мультикод, пилот, семпл
        keyword_mask = codes_str.str.contains('мультикод|пилот|семпл', case=False, na=False)
        
        # Коды с направлением .01 или .02
        pattern = r'^RU0[0-4]\..*\.(01|02)\..*'
        code_mask = codes_str.str.match(pattern, na=False)
        
        return (keyword_mask | code_mask).astype(int)
        
    
    def merge_field_projects(self, array_field_df, cxway_field_df):
        """Объединяет полевые проекты из массива и CXWAY"""
        try:
            if array_field_df is not None and not array_field_df.empty:
                mask_array = (array_field_df['ПО'] == 'Чеккер') | (array_field_df['ПО'] == 'не определено')
                array_filtered = array_field_df[mask_array].copy()
            else:
                array_filtered = pd.DataFrame()
            
            if cxway_field_df is not None and not cxway_field_df.empty:
                mask_cxway = (cxway_field_df['ПО'] == 'CXWAY') | (cxway_field_df['ПО'] == 'не определено')
                cxway_filtered = cxway_field_df[mask_cxway].copy()
            else:
                cxway_filtered = pd.DataFrame()
            
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
            return array_field_df if array_field_df is not None else pd.DataFrame()

    def remove_cxway_from_portal(self, portal_df, google_df):
        """Удаляет из портала проекты, которые в google отмечены как CXWAY (только для обычных проектов)"""
        if portal_df is None or portal_df.empty or google_df is None or google_df.empty:
            return portal_df
        
        code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
        portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
        client_col = self._find_column(google_df, ['Проекты в  https://ru.checker-soft.com', 'Проекты'])
        
        if code_col is None or portal_col is None or client_col is None:
            return portal_df
        
        # Находим проекты с ПО CXWAY
        cxway_mask = google_df[portal_col].astype(str).str.strip().str.upper() == 'CXWAY'
        cxway_df = google_df[cxway_mask].copy()
        
        if cxway_df.empty:
            return portal_df
        
        # Создаем ключи (клиент + код) для CXWAY-проектов
        cxway_df['_key'] = (
            cxway_df[client_col].astype(str).str.strip() + '|' +
            cxway_df[code_col].astype(str).str.strip()
        )
        
        # Создаем ключи в портале (клиент + код)
        portal_code_col = self._find_column(portal_df, ['Код анкеты', 'Код'])
        portal_client_col = self._find_column(portal_df, ['Имя клиента', 'Client'])
        
        if portal_code_col is None or portal_client_col is None:
            return portal_df
        
        portal_df = portal_df.copy()
        portal_df['_key'] = (
            portal_df[portal_client_col].astype(str).str.strip() + '|' +
            portal_df[portal_code_col].astype(str).str.strip()
        )
        
        # Определяем семплы и пилоты в портале
        portal_df['_is_sample_pilot'] = portal_df[portal_code_col].astype(str).str.contains('семпл|пилот', case=False, na=False)
        
        # Удаляем ТОЛЬКО обычные проекты (не семплы/пилоты)
        mask_to_remove = portal_df['_key'].isin(cxway_df['_key']) & ~portal_df['_is_sample_pilot']
        portal_df = portal_df[~mask_to_remove]
        
        # Удаляем временные колонки
        portal_df = portal_df.drop(['_key', '_is_sample_pilot'], axis=1)
        
        return portal_df
    
    def enrich_array_batch(self, array_df, google_df):
        """Быстрое обогащение: Полевой, ПО, ЗОД за один проход"""
        df = array_df.copy()
        
        # Полевой (векторно)
        df['Полевой'] = self._is_field_project_vectorized(df['Код анкеты'])
        
        # ПО из Google
        google_code_col = self._find_column(google_df, ['Код проекта RU00.000.00.01SVZ24', 'Код проекта'])
        google_portal_col = self._find_column(google_df, ['Портал на котором идет проект (для работы полевой команды)', 'ПО'])
        
        portal_map = {}
        if google_code_col and google_portal_col:
            for _, row in google_df.iterrows():
                code = str(row.get(google_code_col, '')).strip()
                portal = str(row.get(google_portal_col, '')).strip()
                if code and code.lower() not in ['nan', 'none', 'null', '']:
                    portal_map[code] = portal
        
        df['ПО'] = df['Код анкеты'].map(portal_map).fillna('не определено')
        
        # ЗОД из справочника
        if 'АСС' in df.columns:
            df['ЗОД'] = df['АСС'].map(ZOD_MAPPING).fillna('')
        else:
            df['ЗОД'] = ''
        
        return df
        
# Глобальный экземпляр
data_cleaner = DataCleaner()






