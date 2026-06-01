# utils/visit_calculator.py
# draft 4.1 - simplified
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io
import calendar
from github_settings import get_plan_adjustment_manager
from github_settings import get_multon_plan_manager
from github_settings import get_plan_adjustment_manager, get_multon_plan_manager, get_multibrand_plan_manager
from data_cleaner import REGION_NAME_TO_CODE


class VisitCalculator:
    
    def _calculate_rs_weights(self, visits_df, project_code, wave_name, region):
        """
        Доли RS = визиты RS в проекте+волне+регионе / все визиты проекта+волны+региона
        """
        try:
            # Ищем колонку RS
            rs_col = None
            for col in visits_df.columns:
                if any(name in str(col).lower() for name in ['эм', 'rs']):
                    rs_col = col
                    break
            
            if not rs_col:
                return {}
            
            
            # Все визиты проекта+волны+региона
            project_wave_region_mask = (
                (visits_df['Код анкеты'] == project_code) &
                (visits_df['Название проекта'] == wave_name) &
                (visits_df['Регион short'] == region)
            )
            filtered_visits = visits_df[project_wave_region_mask]
            
            if filtered_visits.empty:
                return {}
            
            # Визиты по RS
            rs_counts = filtered_visits.groupby(rs_col).size()
            total_visits = rs_counts.sum()
            
            if total_visits == 0:
                return {}
            
            # Доли
            rs_weights = (rs_counts / total_visits).to_dict()
            return rs_weights
            
        except Exception as e:
            print(f"[DEBUG] Ошибка расчета долей RS: {e}")
            return {}
    

    def _get_working_days_in_range(self, start_date, end_date):
        """Возвращает количество рабочих дней (пн-пт) в диапазоне"""
        if pd.isna(start_date) or pd.isna(end_date):
            return 0
        if hasattr(start_date, 'date'):
            start_date = start_date.date()
        if hasattr(end_date, 'date'):
            end_date = end_date.date()
        
        days = 0
        current = start_date
        while current <= end_date:
            if current.weekday() < 5:
                days += 1
            current += timedelta(days=1)
        return days
        
    
    def calculate_plan_with_stages(self, total_plan, duration, coefficients, start_date, finish_date, period_start, period_end):
        
        if total_plan == 0 or duration == 0:
            return 0.0, 0.0
        
        # ПРИВОДИМ ВСЕ ДАТЫ К ТИПУ datetime.date
        if hasattr(start_date, 'date'):
            start_date = start_date.date()
        if hasattr(finish_date, 'date'):
            finish_date = finish_date.date()
        if hasattr(period_start, 'date'):
            period_start = period_start.date()
        if hasattr(period_end, 'date'):
            period_end = period_end.date()
        
        # 1. Разбиваем проект на этапы
        stage_days = duration // 4
        extra_days = duration % 4
        
        stages = []  # каждый элемент: (days, plan, daily_plan)
        plan_remaining = total_plan
        
        for i in range(4):
            days = stage_days + (1 if i < extra_days else 0)
            if i < 3:
                plan = total_plan * coefficients[i]
            else:
                plan = plan_remaining
            plan_remaining -= plan
            daily_plan = plan / days if days > 0 else 0
            stages.append((days, plan, daily_plan))
        
        # 2. Считаем план на дату
        plan_on_date = 0.0
        current_day = 0
        
        for days, plan, daily_plan in stages:
            stage_start = start_date + timedelta(days=current_day)
            stage_end = start_date + timedelta(days=current_day + days - 1)
            
            # Пересечение с отчетным периодом
            intersect_start = max(period_start, stage_start)
            intersect_end = min(period_end, stage_end)
            
            if intersect_start <= intersect_end:
                days_in_period = (intersect_end - intersect_start).days + 1
                plan_on_date += daily_plan * days_in_period
            
            current_day += days
        
        # Средний дневной план
        period_days = (period_end - period_start).days + 1
        daily_plan_avg = plan_on_date / period_days if period_days > 0 else 0
        
        return plan_on_date, daily_plan_avg
            
    def extract_hierarchical_data(self, visits_df, google_df=None, google_df_original=None):
        """
        Создаёт полную иерархию Проект→Клиент→Волна→Регион→DSM→ASM→RS
        с базовой информацией о проекте
        """
        import time
        start_total = time.time()
        
        try:
            # 1. Определяем колонку региона
            region_col = 'Регион short'
            
            # Создаём иерархию из visits_df
            start = time.time()
            hierarchy = pd.DataFrame({
                'Проект': visits_df['Код анкеты'].fillna('Не указано'),
                'Клиент': visits_df['Имя клиента'].fillna('Не указано'),
                'Волна': visits_df['Название проекта'].fillna('Не указано'),
                'Регион': visits_df[region_col].fillna('Не указано'),
                'DSM': visits_df['ЗОД'].fillna('Не указано'),
                'ASM': visits_df['АСС'].fillna('Не указано'),
                'RS': visits_df['ЭМ'].fillna('Не указано'),
                'ПО': visits_df['ПО'].fillna('не определено'),
                'Полевой': visits_df['Полевой']
            })
            # st.write(f"[DETAIL] Создание DataFrame: {time.time() - start:.2f} сек")
            
            # ТОЛЬКО ПОЛЕВЫЕ ПРОЕКТЫ
            start = time.time()
            hierarchy = hierarchy[hierarchy['Полевой'] == 1]
            hierarchy = hierarchy.drop('Полевой', axis=1)
            # st.write(f"[DETAIL] Фильтр полевых: {time.time() - start:.2f} сек")
            
            # Удаляем дубликаты
            start = time.time()
            hierarchy = hierarchy.drop_duplicates().reset_index(drop=True)
            # st.write(f"[DETAIL] Удаление дубликатов: {time.time() - start:.2f} сек")
            
            # Даты - по умолчанию пустые
            hierarchy['Дата старта'] = pd.NaT
            hierarchy['Дата финиша'] = pd.NaT
            

            # ============================================
            # 1. ОСНОВНЫЕ ДАТЫ (из очищенного google_df)
            # ============================================
            if google_df is not None and not google_df.empty:
                start = time.time()
                try:
                    start_mapping = {}
                    finish_mapping = {}
                    
                    for idx, row in google_df.iterrows():
                        code_raw = str(row.get('Код проекта RU00.000.00.01SVZ24', '')).strip()
                        if code_raw and code_raw not in ['nan', '']:
                            start_date = row.get('Дата старта')
                            finish_date = row.get('Дата финиша с продлением')
                            
                            # Разделяем составной код по '/'
                            codes = code_raw.split('/')
                            for code in codes:
                                code = code.strip()
                                if not code:
                                    continue
                                if pd.notna(start_date):
                                    start_mapping[code] = start_date
                                if pd.notna(finish_date):
                                    finish_mapping[code] = finish_date

                    # Функция для получения даты из start_mapping с учетом составных кодов
                    def get_start_date(project_code):
                        code_str = str(project_code)
                        # Если нет слеша — ищем как есть
                        if '/' not in code_str:
                            return start_mapping.get(code_str, pd.NaT)
                        
                        # Разделяем на части и перебираем
                        parts = code_str.split('/')
                        for part in parts:
                            part = part.strip()
                            if not part:
                                continue
                            # Как только нашли часть в start_mapping — берем дату
                            if part in start_mapping:
                                return start_mapping[part]
                        # Если ни одной части нет в mapping — возвращаем пустоту
                        return pd.NaT
                    
                    def get_finish_date(project_code):
                        code_str = str(project_code)
                        if '/' not in code_str:
                            return finish_mapping.get(code_str, pd.NaT)
                        
                        parts = code_str.split('/')
                        for part in parts:
                            part = part.strip()
                            if not part:
                                continue
                            if part in finish_mapping:
                                return finish_mapping[part]
                        return pd.NaT
                    
                    hierarchy['Дата старта'] = hierarchy['Проект'].apply(get_start_date)
                    hierarchy['Дата финиша'] = hierarchy['Проект'].apply(get_finish_date)
                    
                    # Если дат нет, ставим первый и последний день месяца
                    if 'plan_calc_params' in st.session_state:
                        first_day = pd.Timestamp(st.session_state['plan_calc_params']['start_date'])
                        last_day = first_day + pd.offsets.MonthEnd(1)
                    else:
                        today = datetime.now()
                        first_day = pd.Timestamp(year=today.year, month=today.month, day=1)
                        last_day = first_day + pd.offsets.MonthEnd(1)
                    
                    hierarchy['Дата старта'] = hierarchy['Дата старта'].fillna(first_day)
                    hierarchy['Дата финиша'] = hierarchy['Дата финиша'].fillna(last_day)
                    
                except Exception as e:
                    hierarchy['Дата старта'] = pd.NaT
                    hierarchy['Дата финиша'] = pd.NaT
                    
                    if 'plan_calc_params' in st.session_state:
                        first_day = pd.Timestamp(st.session_state['plan_calc_params']['start_date'])
                        last_day = first_day + pd.offsets.MonthEnd(1)
                    else:
                        today = datetime.now()
                        first_day = pd.Timestamp(year=today.year, month=today.month, day=1)
                        last_day = first_day + pd.offsets.MonthEnd(1)
                    
                    hierarchy['Дата старта'] = hierarchy['Дата старта'].fillna(first_day)
                    hierarchy['Дата финиша'] = hierarchy['Дата финиша'].fillna(last_day)
                    pass
                # st.write(f"[DETAIL] Основные даты: {time.time() - start:.2f} сек")
            else:
                # Если google_df нет, ставим даты по умолчанию
                if 'plan_calc_params' in st.session_state:
                    first_day = pd.Timestamp(st.session_state['plan_calc_params']['start_date'])
                    last_day = first_day + pd.offsets.MonthEnd(1)
                else:
                    today = datetime.now()
                    first_day = pd.Timestamp(year=today.year, month=today.month, day=1)
                    last_day = first_day + pd.offsets.MonthEnd(1)
                
                hierarchy['Дата старта'] = first_day
                hierarchy['Дата финиша'] = last_day
            
            # ============================================
            # 2. ОРИГИНАЛЬНЫЕ ДАТЫ ИЗ GOOGLE (для информации и коэффициента)
            # ============================================
            if google_df_original is not None and not google_df_original.empty:
                start = time.time()
                try:
                    # Словарь с приоритетом: (код, волна) → (старт, финиш, метод)
                    date_mapping = {}
                    
                    for idx, row in google_df_original.iterrows():
                        code_raw = str(row.get('Код проекта RU00.000.00.01SVZ24', '')).strip()
                        
                        if code_raw and code_raw not in ['nan', '']:
                            start_date = row.get('Дата старта')
                            finish_date = row.get('Дата финиша с продлением')
                            
                            if pd.notna(start_date) and pd.notna(finish_date):
                                # Получаем волны
                                wave_checker = str(row.get('Название волны на Чекере/ином ПО', '')).strip()
                                wave_dummy = str(row.get('Название волны холостой', '')).strip()
                                
                                waves = []
                                if wave_checker and wave_checker not in ['nan', '']:
                                    waves.append(wave_checker)
                                if wave_dummy and wave_dummy not in ['nan', '']:
                                    waves.append(wave_dummy)
                                
                                # ✅ РАЗДЕЛЯЕМ СОСТАВНОЙ КОД ПО '/' 
                                codes = code_raw.split('/')
                                
                                for single_code in codes:
                                    single_code = single_code.strip()
                                    if not single_code:
                                        continue
                                        
                                    # Для каждой волны
                                    for wave in waves:
                                        key = (single_code, wave)
                                        if key not in date_mapping:
                                            date_mapping[key] = (start_date, finish_date, 'ВК')
                                    
                                    # Только по коду (без волны)
                                    key = (single_code, None)
                                    if key not in date_mapping:
                                        date_mapping[key] = (start_date, finish_date, 'К')
                    
                    # Функция для получения дат по строке иерархии
                    def get_dates(row):
                        code = row['Проект']
                        wave = row['Волна']
                        
                        # Точное совпадение по коду и волне
                        if (code, wave) in date_mapping:
                            start, finish, method = date_mapping[(code, wave)]
                            return pd.Series([start, finish, method])
                        
                        # ✅ ЕСЛИ В КОДЕ ЕСТЬ '/', ПРОВЕРЯЕМ КАЖДУЮ ЧАСТЬ
                        if '/' in code:
                            parts = code.split('/')
                            for part in parts:
                                part = part.strip()
                                if not part:
                                    continue
                                
                                # Ищем совпадение по части кода и волне
                                if (part, wave) in date_mapping:
                                    start, finish, method = date_mapping[(part, wave)]
                                    return pd.Series([start, finish, 'ВК'])
                                
                                # Ищем только по части кода (без волны)
                                for (c, w), (start, finish, _) in date_mapping.items():
                                    if c == part:
                                        return pd.Series([start, finish, 'К'])
                        
                        # Код есть, но волна не совпала (или пустая)
                        for (c, w), (start, finish, _) in date_mapping.items():
                            if c == code:
                                return pd.Series([start, finish, 'К'])
                        
                        # Код не найден
                        return pd.Series([pd.NaT, pd.NaT, 'МП'])
                    
                    # Применяем маппинг
                    hierarchy[['Дата старта_гугл', 'Дата финиша_гугл', 'Метод подбора дат']] = hierarchy.apply(get_dates, axis=1)
                    
                    # Преобразуем в datetime
                    hierarchy['Дата старта_гугл'] = pd.to_datetime(hierarchy['Дата старта_гугл'], errors='coerce')
                    hierarchy['Дата финиша_гугл'] = pd.to_datetime(hierarchy['Дата финиша_гугл'], errors='coerce')
                    
                except Exception as e:
                    hierarchy['Дата старта_гугл'] = pd.NaT
                    hierarchy['Дата финиша_гугл'] = pd.NaT
                    hierarchy['Метод подбора дат'] = 'МП'
                    pass
                # st.write(f"[DETAIL] Оригинальные даты: {time.time() - start:.2f} сек")
            else:
                hierarchy['Дата старта_гугл'] = pd.NaT
                hierarchy['Дата финиша_гугл'] = pd.NaT
                hierarchy['Метод подбора дат'] = 'МП'
                
            
            # Рассчитываем длительность
            start = time.time()
            hierarchy['Длительность'] = 0
            mask_valid_dates = hierarchy['Дата старта'].notna() & hierarchy['Дата финиша'].notna()
            
            if mask_valid_dates.any():
                hierarchy.loc[mask_valid_dates, 'Длительность'] = (
                    hierarchy.loc[mask_valid_dates, 'Дата финиша'] - 
                    hierarchy.loc[mask_valid_dates, 'Дата старта']
                ).dt.days + 1
            # st.write(f"[DETAIL] Расчет длительности: {time.time() - start:.2f} сек")
            
            # Сортируем
            start = time.time()
            hierarchy = hierarchy.sort_values(['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM', 'RS'])
            hierarchy = hierarchy[hierarchy['RS'] != 'Итого']
            # st.write(f"[DETAIL] Сортировка: {time.time() - start:.2f} сек")
            
            # st.write(f"[DETAIL] ВСЕГО Иерархия: {time.time() - start_total:.2f} сек")
            
            return hierarchy
            
        except KeyError as e:
            return pd.DataFrame()
        except Exception as e:
            return pd.DataFrame()
    
    def calculate_hierarchical_plan_on_date(self, hierarchy_df, visits_df, calc_params, google_df=None, optima_df=None):
        
        coefficients = calc_params.get('coefficients', [0.25, 0.25, 0.25, 0.25])
        
        try:
            if hierarchy_df.empty or visits_df.empty:
                return pd.DataFrame()
            
            start_period = calc_params['start_date']
            end_period = calc_params['end_date']
            month_days = calendar.monthrange(start_period.year, start_period.month)[1]
            
            # ============================================
            # 1. ПРЕДВАРИТЕЛЬНЫЙ РАСЧЕТ ВСЕХ НУЖНЫХ ДАННЫХ
            # ============================================
            
            # КВОТЫ МУЛТОН
            multon_quotas = {}
            if google_df is not None and not google_df.empty:
                project_col = 'Проекты в  https://ru.checker-soft.com'
                code_col = 'Код проекта RU00.000.00.01SVZ24'
                kvota_col = 'Квота'
                if all(col in google_df.columns for col in [project_col, code_col, kvota_col]):
                    multon_mask = google_df[project_col].astype(str).str.strip() == 'Мултон'
                    for _, row in google_df[multon_mask].iterrows():
                        code = str(row.get(code_col, '')).strip()
                        if code and code not in ['', 'nan', 'None', 'null']:
                            try:
                                multon_quotas[code] = float(row.get(kvota_col, 0))
                            except:
                                pass
            
            
            # КВОТЫ ПРОДАТА
            prodata_quotas = {}
            if google_df is not None and not google_df.empty:
                project_col = 'Проекты в  https://ru.checker-soft.com'
                code_col = 'Код проекта RU00.000.00.01SVZ24'
                kvota_col = 'Квота'
                if all(col in google_df.columns for col in [project_col, code_col, kvota_col]):
                    prodata_mask = google_df[project_col].astype(str).str.strip().str.startswith('Мониторинг')
                    for _, row in google_df[prodata_mask].iterrows():
                        code = str(row.get(code_col, '')).strip()
                        if code and code not in ['', 'nan', 'None', 'null']:
                            try:
                                prodata_quotas[code] = float(row.get(kvota_col, 0))
                            except:
                                pass
            
            # Планы клиентов+проектов+волн+регионов
            project_wave_region_plans = visits_df.groupby([
                'Имя клиента',
                'Код анкеты', 
                'Название проекта',
                'Регион short'
            ]).size().to_dict()

            # ============================================
            # РАСЧЕТ ВЕСОВ RS ДЛЯ РАСПРЕДЕЛЕНИЯ ПЛАНА
            # ============================================
            rs_weights_for_plan = {}
            
            # Фильтруем только УДАЛЕННЫЕ записи (исключаем их)
            status_col = None
            for col in visits_df.columns:
                if col.strip() == 'Статус':
                    status_col = col
                    break
            
            if status_col:
                # Исключаем только 'Удалено'
                not_deleted_mask = visits_df[status_col].astype(str).str.strip() != 'Удалено'
                visits_for_weights = visits_df[not_deleted_mask].copy()
            else:
                visits_for_weights = visits_df.copy()
            
            # Группируем по (клиент, код, волна, регион, RS)
            rs_counts = visits_for_weights.groupby([
                'Имя клиента',
                'Код анкеты',
                'Название проекта',
                'Регион short',
                'ЭМ'  # RS
            ]).size().reset_index(name='count')
            
            # Для каждой группы (клиент, код, волна, регион) считаем доли
            if not rs_counts.empty:
                for _, group in rs_counts.groupby(['Имя клиента', 'Код анкеты', 'Название проекта', 'Регион short']):
                    key = (group['Имя клиента'].iloc[0], 
                           group['Код анкеты'].iloc[0], 
                           group['Название проекта'].iloc[0], 
                           group['Регион short'].iloc[0])
                    total = group['count'].sum()
                    if total > 0:
                        for _, row in group.iterrows():
                            weight_key = key + (row['ЭМ'],)
                            rs_weights_for_plan[weight_key] = row['count'] / total

            # ============================================
            # ЗАГРУЗКА КОРРЕКТИРОВОК ПЛАНА (ОДИН РАЗ)
            # ============================================
            plan_adjustments = {}
            try:
                adj_manager = get_plan_adjustment_manager()
                all_adjustments = adj_manager.get_adjustments()
                for adj in all_adjustments:
                    key = f"{adj.get('project_name', '')}|{adj.get('wave_name', '')}|{adj.get('project_code', '')}"
                    current = plan_adjustments.get(key, 0)
                    plan_adjustments[key] = current + adj.get('adjustment_value', 0)
            except Exception as e:
                plan_adjustments = {}


        
            # ============================================
            # 2. ПРЕДВАРИТЕЛЬНЫЙ РАСЧЕТ ВЕСОВ RS
            # ============================================
            
            # Создаем словарь весов RS: ключ = (проект, волна, регион)
            rs_weights_cache = {}
            
            # Группируем все визиты по проекту, волне, региону и RS
            visits_grouped = visits_df.groupby([
                'Код анкеты', 
                'Название проекта', 
                'Регион short',
                'ЭМ'  # RS колонка
            ]).size().reset_index(name='count')
            
            # Для каждой комбинации считаем общее количество и долю
            for _, group in visits_grouped.groupby(['Код анкеты', 'Название проекта', 'Регион short']):
                key = (group['Код анкеты'].iloc[0], group['Название проекта'].iloc[0], group['Регион short'].iloc[0])
                total = group['count'].sum()
                if total > 0:
                    rs_weights_cache[key] = {}
                    for _, row in group.iterrows():
                        rs_weights_cache[key][row['ЭМ']] = row['count'] / total
            
            # ============================================
            # 3. ПРЕДВАРИТЕЛЬНЫЙ РАСЧЕТ РАСПРЕДЕЛЕНИЯ ПО РЕГИОНАМ
            # ============================================
            
            # Для Мултон и ПроДата считаем количество регионов
            multon_regions = {}
            prodata_regions = {}
            
            for _, row in hierarchy_df.iterrows():
                project_code = row['Проект']
                po = row['ПО']
                client = row['Клиент']
                region = row['Регион']
                
                if po == 'ПО клиента' and client == 'Мултон':
                    if project_code not in multon_regions:
                        multon_regions[project_code] = set()
                    multon_regions[project_code].add(region)

                elif po == 'Мониторинги':
                    if project_code not in prodata_regions:
                        prodata_regions[project_code] = set()
                    prodata_regions[project_code].add(region)
                    

            # === РАСШИРЕНИЕ ИЕРАРХИИ ДЛЯ МУЛТОН (добавляем проекты без визитов) ===
            from github_settings import get_multon_plan_manager
            multon_manager = get_multon_plan_manager()
            plan_df = multon_manager.load_plan()
            
            if not plan_df.empty:
                # Получаем существующие комбинации из иерархии
                existing_combinations = set(zip(
                    hierarchy_df[hierarchy_df['Клиент'] == 'Мултон']['Проект'],
                    hierarchy_df[hierarchy_df['Клиент'] == 'Мултон']['Регион'],
                    hierarchy_df[hierarchy_df['Клиент'] == 'Мултон']['ASM']
                ))
                
                # Получаем все комбинации из JSON
                json_combinations = set(zip(plan_df['project_code'], plan_df['region'], plan_df['rs']))
                
                # Находим недостающие комбинации
                missing_combinations = json_combinations - existing_combinations
                
                if missing_combinations:
                    # Создаем строки для недостающих комбинаций
                    new_rows = []
                    for project_code, region, asm in missing_combinations:
                        new_rows.append({
                            'Проект': project_code,
                            'Клиент': 'Мултон',
                            'Волна': 'не указано',
                            'Регион': region,
                            'DSM': 'не указано',
                            'ASM': asm,
                            'RS': 'нет',
                            'ПО': 'ПО клиента',
                            'Полевой': 1,
                            'Дата старта': pd.Timestamp(start_period),
                            'Дата финиша': pd.Timestamp(end_period),
                            'Длительность': (pd.Timestamp(end_period) - pd.Timestamp(start_period)).days + 1,
                            'Метод подбора дат': 'План (нет визитов)'
                        })
                    
                    # Добавляем в иерархию
                    new_rows_df = pd.DataFrame(new_rows)
                    hierarchy_df = pd.concat([hierarchy_df, new_rows_df], ignore_index=True)
        
            
            # ============================================
            # 4. ОСНОВНОЙ ЦИКЛ (ОПТИМИЗИРОВАННЫЙ)
            # ============================================
            
            results = []
            
            for _, row in hierarchy_df.iterrows():
                region = row['Регион']
                project_code = row['Проект']
                wave_name = row['Волна']
                po = row['ПО']
                client = row['Клиент']
                rs_name = row['RS']

                # ПРОВЕРКА ДАТ
                start_date = row['Дата старта']
                finish_date = row['Дата финиша']
                duration = row['Длительность']
                
                if pd.isna(start_date) or pd.isna(finish_date) or duration <= 0:
                    continue
                
                if end_period < start_date.date() or start_period > finish_date.date():
                    continue
                
                period_start = max(start_period, start_date.date())
                period_end = min(end_period, finish_date.date())
                days_in_period = max(0, (period_end - period_start).days + 1)

                
                # ============================================
                # РАСЧЕТ КОЭФФИЦИЕНТА МЕСЯЦА
                # ============================================
                start_date_google = row.get('Дата старта_гугл', None)
                finish_date_google = row.get('Дата финиша_гугл', None)
                
                if pd.isna(start_date_google):
                    start_date_google = start_date
                if pd.isna(finish_date_google):
                    finish_date_google = finish_date
                
                # Приводим все к Timestamp (чтобы можно было сравнивать)
                if hasattr(start_date_google, 'date'):
                    start_ts = pd.Timestamp(start_date_google.date())
                elif hasattr(start_date_google, 'year'):
                    start_ts = pd.Timestamp(start_date_google)
                else:
                    start_ts = pd.Timestamp(start_date.date())
                
                if hasattr(finish_date_google, 'date'):
                    finish_ts = pd.Timestamp(finish_date_google.date())
                elif hasattr(finish_date_google, 'year'):
                    finish_ts = pd.Timestamp(finish_date_google)
                else:
                    finish_ts = pd.Timestamp(finish_date.date())
                
                # Границы текущего месяца (как Timestamp)
                month_start_ts = pd.Timestamp(calc_params['start_date'])
                month_end_ts = month_start_ts + pd.offsets.MonthEnd(1)
                
                # Рабочие дни проекта в текущем месяце (числитель)
                project_start_in_month = max(start_ts, month_start_ts)
                project_end_in_month = min(finish_ts, month_end_ts)
                working_days_in_month = self._get_working_days_in_range(project_start_in_month, project_end_in_month)
                
                # Общие рабочие дни проекта (знаменатель)
                total_working_days = self._get_working_days_in_range(start_ts, finish_ts)
                
                if total_working_days > 0:
                    month_coefficient = working_days_in_month / total_working_days
                else:
                    month_coefficient = 1.0
                # ============================================
                
                
                if days_in_period == 0:
                    continue
                
                # Определяем total_plan и считаем план на дату
                if po == 'Мониторинги':
                    total_plan = prodata_quotas.get(project_code, 0)
                    if total_plan <= 0:
                        continue
                    num_regions = len(prodata_regions.get(project_code, []))
                    if num_regions > 0:
                        total_plan = total_plan / num_regions
                    # Мониторинги: старая логика, без этапов
                    rs_plan_on_date = total_plan
                    rs_daily_plan = total_plan
                
                else:
                    # Для всех остальных типов (Чеккер, CXWAY, Easymerch, Мултон, Оптима)
                    # Сначала рассчитываем total_plan (как раньше)
                    
                    if po == 'ПО клиента' and client == 'Мултон':
                        # Проверяем, загружен ли Easymerch
                        if 'easymerch' not in st.session_state.uploaded_files:
                            continue
                            
                        # Загружаем распределение плана из JSON
                        from github_settings import get_multon_plan_manager
                        multon_manager = get_multon_plan_manager()
                        plan_df = multon_manager.load_plan()
                        
                        if plan_df.empty:
                            # Нет распределения → проект не существует
                            total_plan = 0
                        else:
                            # Ищем план по (код_проекта, регион, RS)
                            mask = (plan_df['project_code'] == project_code) & \
                                   (plan_df['region'] == region) & \
                                   (plan_df['rs'] == row['ASM'])
                            
                            if mask.any():
                                total_plan = plan_df.loc[mask, 'plan'].iloc[0]
                            else:
                                total_plan = 0
                        
                        if total_plan <= 0:
                            continue
                    
                    elif client == 'Мультибренд 2024' and po == 'CXWAY':
                        # Специальная логика для Мультибренд 2024
                        
                        # 1. Определяем тип волны по названию (после последнего '_')
                        wave_parts = wave_name.split('_')
                        if len(wave_parts) >= 2:
                            wave_type = '_'.join(wave_parts[1:])
                        else:
                            wave_type = wave_name
                        
                        # 2. Загружаем распределение плана
                        multibrand_manager = get_multibrand_plan_manager()
                        dilers_df, pronto_df = multibrand_manager.load_plan()
                        
                        # 3. Определяем план в зависимости от типа волны
                        asm_from_plan = row['ASM']
                        rs_from_plan = row['RS']
                        skip_plan_correction = False  # ← значение по умолчанию
                        
                        if wave_type == 'Нерезультативные_Пронто_Дилеры':
                            total_plan = 0
                            skip_plan_correction = True
                            
                        elif wave_type == 'Дилеры':
                            plan_row = dilers_df[dilers_df['region_code'] == region]
                            if not plan_row.empty:
                                total_plan = plan_row.iloc[0]['plan']
                                asm_from_plan = plan_row.iloc[0]['asm']
                                rs_from_plan = plan_row.iloc[0]['rs']
                            else:
                                total_plan = 0
                            
                        elif wave_type == 'Пронто' or wave_type == 'Пронто М':
                            mapped_region = region
                            if region == 'MS':
                                if wave_type == 'Пронто':
                                    mapped_region = 'Москва и Московская область'
                                else:
                                    mapped_region = 'МСК дистр.'
                            
                            plan_row = pronto_df[pronto_df['region_code'] == mapped_region]
                            if not plan_row.empty:
                                total_plan = plan_row.iloc[0]['plan']
                                asm_from_plan = plan_row.iloc[0]['asm']
                                rs_from_plan = plan_row.iloc[0]['rs']
                            else:
                                plan_row = pronto_df[pronto_df['region_code'] == region]
                                if not plan_row.empty:
                                    total_plan = plan_row.iloc[0]['plan']
                                    asm_from_plan = plan_row.iloc[0]['asm']
                                    rs_from_plan = plan_row.iloc[0]['rs']
                                else:
                                    total_plan = 0
                            
                        else:
                            total_plan = 0
                        
                        if total_plan <= 0:
                            continue
                        
                        # Рассчитываем план на дату с учетом этапов
                        rs_plan_on_date, rs_daily_plan = self.calculate_plan_with_stages(
                            total_plan,
                            duration,
                            coefficients,
                            start_date,
                            finish_date,
                            period_start,
                            period_end
                        )
    
                    else:  # Чеккер, CXWAY (другие), Easymerch, Optima
                        client_name = row['Клиент']
                        plan_key = (client_name, project_code, wave_name, region)
                        total_plan = project_wave_region_plans.get(plan_key, 0)
                        
                        if total_plan <= 0:
                            continue
                        
                        # ПРИМЕНЯЕМ КОЭФФИЦИЕНТ МЕСЯЦА
                        total_plan = round(total_plan * month_coefficient, 1)
                        
                        # Распределяем план по RS с помощью весов
                        weight_key = (client_name, project_code, wave_name, region, rs_name)
                        weight = rs_weights_for_plan.get(weight_key, 0)
                        
                        if weight > 0:
                            total_plan = round(total_plan * weight, 1)
                        
                        # Рассчитываем план на дату с учетом этапов
                        rs_plan_on_date, rs_daily_plan = self.calculate_plan_with_stages(
                            total_plan,
                            duration,
                            coefficients,
                            start_date,
                            finish_date,
                            period_start,
                            period_end
                        )
                        
                        asm_from_plan = row['ASM']
                        rs_from_plan = rs_name
                        skip_plan_correction = False
                    


                results.append({
                    'Проект': project_code,
                    'Клиент': row['Клиент'],
                    'Волна': wave_name,
                    'Регион': region,
                    'DSM': row['DSM'],
                    'ASM': asm_from_plan,
                    'RS': rs_from_plan,
                    'ПО': po,
                    'Уровень': 'RS',
                    'План проекта, шт.': total_plan,
                    'План на дату, шт.': round(rs_plan_on_date, 1),
                    'Длительность': int(duration),
                    'Дата старта': start_date,
                    'Дата финиша': finish_date,
                    'Дата старта_гугл': start_date_google,     
                    'Дата финиша_гугл': finish_date_google,    
                    'Коэффициент месяца': month_coefficient,
                    'Метод подбора дат': row['Метод подбора дат'],
                    'Дней в периоде': days_in_period,
                    'Дневной план RS, шт.': round(rs_daily_plan, 2),
                    'skip_plan_correction': skip_plan_correction
                })


            # ============================================
            # ПРИМЕНЕНИЕ КОРРЕКТИРОВОК (ПОСЛЕ СБОРА ВСЕХ ДАННЫХ)
            # ============================================
            
            if plan_adjustments and results:
                # Преобразуем в DataFrame
                results_df = pd.DataFrame(results)
                
                # Создаем ключ проекта (клиент + волна + код)
                results_df['_project_key'] = (
                    results_df['Клиент'].astype(str).str.strip() + '|' +
                    results_df['Волна'].astype(str).str.strip() + '|' +
                    results_df['Проект'].astype(str).str.strip()
                )
                
                # 1. Считаем общий план проекта
                project_totals = results_df.groupby('_project_key')['План проекта, шт.'].sum().to_dict()
                
                # 2. Применяем корректировки
                adjusted_totals = {}
                for key, total in project_totals.items():
                    adjustment = plan_adjustments.get(key, 0)
                    new_total = total + adjustment
                    if new_total < 0:
                        st.warning(f"⚠️ Корректировка {adjustment} для проекта {key} делает план отрицательным ({total} + {adjustment} = {new_total}). Корректировка НЕ применена.")
                        adjusted_totals[key] = total
                    else:
                        adjusted_totals[key] = new_total
                
                # 3. Вычисляем коэффициенты корректировки
                correction_factors = {}
                for key in project_totals:
                    if project_totals[key] > 0:
                        correction_factors[key] = adjusted_totals[key] / project_totals[key]
                    else:
                        correction_factors[key] = 1
                
                # 4. Применяем коэффициенты к каждой строке
                for idx, row in results_df.iterrows():
                    # Пропускаем строки с флагом skip_plan_correction
                    if row.get('skip_plan_correction', False):
                        continue
                    key = row['_project_key']
                    coeff = correction_factors.get(key, 1)
                    
                    if coeff != 1:
                        old_plan = row['План проекта, шт.']
                        new_plan = old_plan * coeff
                        
                        old_plan_date = row['План на дату, шт.']
                        new_plan_date = old_plan_date * coeff
                        
                        results_df.at[idx, 'План проекта, шт.'] = float(new_plan)
                        results_df.at[idx, 'План на дату, шт.'] = round(new_plan_date, 1)
                        
                        if 'Дней в периоде' in results_df.columns:
                            days = row['Дней в периоде']
                            if days > 0:
                                results_df.at[idx, 'Дневной план RS, шт.'] = round(new_plan_date / days, 2)
                
                # Удаляем временную колонку
                results_df = results_df.drop('_project_key', axis=1)
                
                # Обновляем results
                results = results_df.to_dict('records')
            if not results:
                return pd.DataFrame()
            return pd.DataFrame(results)
        
        except Exception as e:
            print(f"Ошибка в calculate_hierarchical_plan_on_date: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()
        
    def calculate_dynamics_fact(self, visits_df, calc_params, group_cols):
        """
        Рассчитывает факт визитов в динамике по дням
        """
        if visits_df is None or visits_df.empty:
            return pd.DataFrame()
        
        df = visits_df.copy()
        
        # 1. Проверка наличия даты визита
        if 'Дата визита' not in df.columns:
            return pd.DataFrame()
        
        # 2. Приводим дату к типу date
        df['Дата визита'] = pd.to_datetime(df['Дата визита'], errors='coerce', dayfirst=True)
        df['Дата'] = df['Дата визита'].dt.date
        
        df = df[df['Дата'].notna()].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        # 3. Фильтруем по периоду
        start_date = calc_params['start_date']
        end_date = calc_params['end_date']
        
        if hasattr(start_date, 'date'):
            start_date = start_date.date()
        if hasattr(end_date, 'date'):
            end_date = end_date.date()
        
        mask = (df['Дата'] >= start_date) & (df['Дата'] <= end_date)
        df = df[mask].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        # 4. Фильтруем только выполненные визиты
        status_col = None
        for col in df.columns:
            col_clean = col.strip()
            if col_clean == 'Статус':
                status_col = col
                break
        
        if status_col:
            completed_statuses = [
                'Выполнено', 'выполнен',
                'Заполнена', 'Проверена',
                'Завершено', 'Готово'
            ]
            completed_mask = df[status_col].isin(completed_statuses)
            df = df[completed_mask].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        # 5. Проверяем наличие колонок для группировки
        existing_group_cols = [col for col in group_cols if col in df.columns]
        
        if not existing_group_cols:
            result = df.groupby(['Дата']).size().reset_index(name='Факт')
            return result
        
        # 6. Группировка
        groupby_cols = existing_group_cols + ['Дата']
        result = df.groupby(groupby_cols).size().reset_index(name='Факт')
        
        return result
    
    def calculate_hierarchical_fact_on_date(self, plan_df, visits_df, calc_params, status_filter='completed'):
        try:
            if plan_df.empty or visits_df.empty:
                return pd.DataFrame()
            
            result_df = plan_df.copy()
            region_col = 'Регион short'
            
            # Ищем колонку статуса
            
            status_col = None
            for col in visits_df.columns:
                col_clean = col.strip()
                if col_clean == 'Статус':
                    status_col = col
                    break
            
            if not status_col:
                result_df['Факт проекта, шт.'] = 0
                result_df['Факт на дату, шт.'] = 0
                return result_df
            
            # Ищем колонку RS
            rs_col = None
            for col in visits_df.columns:
                if any(name in str(col).lower() for name in ['эм', 'rs']):
                    rs_col = col
                    break
            
            if not rs_col:
                result_df['Факт проекта, шт.'] = 0
                result_df['Факт на дату, шт.'] = 0
                return result_df
            
            # ФИЛЬТРЫ
            # КОНВЕРТИРУЕМ ДАТУ ВИЗИТА
            if 'Дата визита' in visits_df.columns:
                # ПРОСТО: pandas сам определит формат
                visits_df['Дата визита'] = pd.to_datetime(
                    visits_df['Дата визита'], 
                    errors='coerce',
                    dayfirst=True
                )
                
                # Нормализуем к началу дня (00:00:00)
                visits_df['Дата визита'] = visits_df['Дата визита'].dt.normalize()
            
            # ФИЛЬТРЫ - выбор статуса в зависимости от параметра
            if status_filter == 'completed':
                status_mask = visits_df[status_col].isin([
                    'Выполнено', 'выполнен',
                    'Заполнена', 'Проверена', 'Принята',
                    'Завершено', 'Готово'
                ])
                suffix = ''
            elif status_filter == 'assigned':
                status_mask = visits_df[status_col].astype(str).str.strip() == 'Поручено'
                suffix = '_поручено'
            elif status_filter == 'not_assigned':
                status_mask = visits_df[status_col].astype(str).str.strip() == 'Не поручено'
                suffix = '_не_поручено'
            else:
                status_mask = None
                suffix = ''
                
            start_date = pd.Timestamp(calc_params['start_date'])
            end_date = pd.Timestamp(calc_params['end_date'])
            period_mask = (
                (visits_df['Дата визита'] >= start_date) &
                (visits_df['Дата визита'] <= end_date)
            )

            # СЧИТАЕМ ФАКТЫ
            filtered_df = visits_df[status_mask]
            rs_facts_total = filtered_df.groupby([
                'Имя клиента',
                'Код анкеты',
                'Название проекта',
                region_col,
                'АСС',
                rs_col
            ]).size().to_dict()

            # Суммируем оплату по группам
            payment_sum = filtered_df.groupby([
                'Имя клиента',
                'Код анкеты',
                'Название проекта',
                region_col,
                'АСС',
                rs_col
            ])['Оплата факт'].sum().to_dict()
            
            # Если нет колонки "Дата визита" — факт на дату = факт проекта
            if 'Дата визита' in visits_df.columns:
                start_date = pd.Timestamp(calc_params['start_date'])
                end_date = pd.Timestamp(calc_params['end_date'])
                period_mask = (
                    (visits_df['Дата визита'] >= start_date) &
                    (visits_df['Дата визита'] <= end_date)
                )
                filtered_in_period = visits_df[status_mask & period_mask]
                rs_facts_period = filtered_in_period.groupby([
                    'Имя клиента',
                    'Код анкеты',
                    'Название проекта',
                    region_col,
                    'АСС',
                    rs_col
                ]).size().to_dict()
            else:
                # Для Optima и других без дат
                rs_facts_period = rs_facts_total.copy()
    
            # ✅ СОЗДАЁМ КОЛОНКИ
            result_df[f'Факт проекта{suffix}, шт.'] = 0
            result_df[f'Факт на дату{suffix}, шт.'] = 0
            
            # ✅ ЗАПОЛНЯЕМ
            for idx in result_df[result_df['Уровень'] == 'RS'].index:
                row = result_df.loc[idx]
                project = str(row['Проект']).strip()
                wave = str(row['Волна']).strip()
                region = str(row['Регион']).strip()
                rs = str(row['RS']).strip()
                
                client_name = str(row['Клиент']).strip()
                asm = str(row['ASM']).strip()
                key = (client_name, project, wave, region, asm, rs)
                result_df.at[idx, f'Факт проекта{suffix}, шт.'] = rs_facts_total.get(key, 0)
                result_df.at[idx, f'Факт на дату{suffix}, шт.'] = rs_facts_period.get(key, 0)
                if suffix == '':
                    result_df.at[idx, 'Оплата факт'] = payment_sum.get(key, 0)
                else:
                    result_df.at[idx, f'Оплата{suffix}'] = payment_sum.get(key, 0)
    
            return result_df
            
        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
            return pd.DataFrame()


    # 6. РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ ПОКАЗАТЕЛЕЙ
    def _calculate_metrics(self, fact_df, calc_params=None, plan_df=None):
        """
        Расчет всех метрик план/факта
        """
        df = fact_df.copy()
        
        # 1. ПЛАН (из plan_df если есть)
        if plan_df is not None:
            if 'План на дату, шт.' in plan_df.columns:
                df['План на дату, шт.'] = plan_df['План на дату, шт.']
            # Сохраняем дополнительные колонки из plan_df
            if 'Дата старта_гугл' in plan_df.columns:
                df['Дата старта_гугл'] = plan_df['Дата старта_гугл']
            if 'Дата финиша_гугл' in plan_df.columns:
                df['Дата финиша_гугл'] = plan_df['Дата финиша_гугл']
            if 'Коэффициент месяца' in plan_df.columns:
                df['Коэффициент месяца'] = plan_df['Коэффициент месяца']
            if 'Метод подбора дат' in plan_df.columns:
                df['Метод подбора дат'] = plan_df['Метод подбора дат']
        elif 'План на дату, шт.' not in df.columns:
            df['План на дату, шт.'] = 0
        
        # 2. ФАКТ
        if 'Факт на дату, шт.' not in df.columns:
            df['Факт на дату, шт.'] = 0

        # 2.5 КОРРЕКТИРОВКА ПЛАНА: если факт > план
        mask_fact_gt_plan_project = df['Факт проекта, шт.'] > df['План проекта, шт.']
        if mask_fact_gt_plan_project.any():
            df.loc[mask_fact_gt_plan_project, 'План проекта, шт.'] = df.loc[mask_fact_gt_plan_project, 'Факт проекта, шт.'].copy()
        
        mask_fact_gt_plan_date = df['Факт на дату, шт.'] > df['План на дату, шт.']
        if mask_fact_gt_plan_date.any():
            df.loc[mask_fact_gt_plan_date, 'План на дату, шт.'] = df.loc[mask_fact_gt_plan_date, 'Факт на дату, шт.'].copy()
        
        # 3. БАЗОВЫЕ МЕТРИКИ
        mask = df['План на дату, шт.'] > 0
        
        # % выполнения плана
        df['%ПФ на дату'] = 0.0
        if mask.any():
            df.loc[mask, '%ПФ на дату'] = (
                df.loc[mask, 'Факт на дату, шт.'] / 
                df.loc[mask, 'План на дату, шт.'] * 100
            ).round(1)

        # Прогноз ВП для краткого отчета
        
        if 'Факт проекта, шт.' in df.columns and 'План проекта, шт.' in df.columns and 'Длительность' in df.columns and 'Дата старта' in df.columns:
            # Получаем end_date из calc_params
            if calc_params:
                end_date = pd.Timestamp(calc_params['end_date'])
            else:
                end_date = pd.Timestamp(datetime.now())
            
            # Преобразуем дату старта в datetime
            start_date_series = pd.to_datetime(df['Дата старта'])
            
            # Рассчитываем истекшие дни (от старта до end_date)
            days_passed = (end_date - start_date_series).dt.days + 1
            
            # Ограничиваем: не меньше 0, не больше длительности
            days_passed = days_passed.clip(lower=0, upper=df['Длительность'])
            
            # Маска для расчета (истекшие дни > 0)
            mask_forecast = days_passed > 0
            
            # Прогноз, шт.
            df['Прогноз, шт.'] = 0.0
            if mask_forecast.any():
                df.loc[mask_forecast, 'Прогноз, шт.'] = (
                    df.loc[mask_forecast, 'Факт проекта, шт.'] * 
                    (df.loc[mask_forecast, 'Длительность'] / days_passed.loc[mask_forecast])
                ).round(1)
            
            # Прогноз ВП, %
            df['Прогноз ВП, %'] = 0.0
            mask_plan = df['План проекта, шт.'] > 0
            if mask_plan.any():
                df.loc[mask_plan, 'Прогноз ВП, %'] = (
                    df.loc[mask_plan, 'Прогноз, шт.'] / df.loc[mask_plan, 'План проекта, шт.'] * 100
                ).round(1)
            
            # Удаляем временную колонку Прогноз, шт.
            # df = df.drop('Прогноз, шт.', axis=1)
        
        # Отклонение в штуках
        df['△План/Факт на дату, шт.'] = (
            df['План на дату, шт.'] - df['Факт на дату, шт.']
        ).round(1)
        
        # Отклонение в процентах 🔴 ПРАВИЛЬНАЯ ФОРМУЛА
        df['△План/Факт,%'] = 0.0
        if mask.any():
            df.loc[mask, '△План/Факт,%'] = (
                (df.loc[mask, 'Факт на дату, шт.'] / 
                 df.loc[mask, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # Исполнение проекта
        df['Исполнение Проекта,%'] = df['%ПФ на дату']

        # План/Факт проекта, %
        mask_project = df['План проекта, шт.'] > 0
        df['План/Факт проекта,%'] = 0.0
        if mask_project.any():
            df.loc[mask_project, 'План/Факт проекта,%'] = (
                df.loc[mask_project, 'Факт проекта, шт.'] / 
                df.loc[mask_project, 'План проекта, шт.'] * 100
            ).round(1)
        
        # △План/Факт проекта, шт
        df['△План/Факт проекта, шт'] = (
            df['Факт проекта, шт.'] - df['План проекта, шт.']
        ).round(1)
        
        # △План/Факт проекта, %
        df['△План/Факт проекта, %'] = 0.0
        if mask_project.any():
            df.loc[mask_project, '△План/Факт проекта, %'] = (
                (df.loc[mask_project, 'Факт проекта, шт.'] / 
                 df.loc[mask_project, 'План проекта, шт.']) - 1
            ).round(3) * 100
        
        # 4. МЕТРИКИ ПО ДНЯМ (если есть calc_params)
        if calc_params and 'Дата старта' in df.columns and 'Дата финиша' in df.columns:
            end_period = pd.Timestamp(calc_params['end_date'])
            
            # # Конвертируем даты
            # df['Дата старта'] = pd.to_datetime(df['Дата старта'], errors='coerce')
            # df['Дата финиша'] = pd.to_datetime(df['Дата финиша'], errors='coerce')
            
            # Дней потрачено / до конца
            df['Дней потрачено'] = 0
            df['Дней до конца проекта'] = 0
            
            mask_dates = df['Дата старта'].notna() & df['Дата финиша'].notna() & (df['Длительность'] > 0)
            
            if mask_dates.any():
                # Дней от старта до конца периода
                days_spent = (end_period - df.loc[mask_dates, 'Дата старта']).dt.days + 1
                df.loc[mask_dates, 'Дней потрачено'] = days_spent.clip(lower=0, upper=df.loc[mask_dates, 'Длительность'])
                
                # Дней до конца
                days_left = (df.loc[mask_dates, 'Дата финиша'] - end_period).dt.days
                df.loc[mask_dates, 'Дней до конца проекта'] = days_left.clip(lower=0)
            
            # Утилизация времени
            df['Утилизация тайминга, %'] = 0.0
            mask_duration = df['Длительность'] > 0
            if mask_duration.any():
                df.loc[mask_duration, 'Утилизация тайминга, %'] = (
                    df.loc[mask_duration, 'Дней потрачено'] / 
                    df.loc[mask_duration, 'Длительность'] * 100
                ).round(1)
            

            # # Средний план на день для 100% плана
            # df['Ср. план на день для 100% плана'] = 0.0
            # if mask_duration.any() and mask.any():
            #     remaining_plan = df.loc[mask & mask_duration, 'План проекта, шт.'] - df.loc[mask & mask_duration, 'Факт проекта, шт.']
            #     remaining_plan = remaining_plan.clip(lower=0)  # защита от отрицательных
            #     days_left = df.loc[mask & mask_duration, 'Дней до конца проекта'].replace(0, 1)
            #     df.loc[mask & mask_duration, 'Ср. план на день для 100% плана'] = (
            #         np.ceil(remaining_plan / days_left)
            #     ).astype(int)

            # # === ОТЛАДКА ===
            # st.write("### 🔍 Отладка: Средний план на день (план ≥ 200)")
            
            # # Фильтруем строки с планом ≥ 200
            # high_plan_mask = (mask & mask_duration) & (df['План проекта, шт.'] >= 200)
            
            # if high_plan_mask.any():
            #     # Получаем первые 3 индекса
            #     sample_indices = df.index[high_plan_mask][:3]
                
            #     for idx in sample_indices:
            #         st.write(f"**Проект:** {df.loc[idx, 'Проект']}")
            #         st.write(f"**Клиент:** {df.loc[idx, 'Клиент']}")
            #         st.write(f"  План проекта: {df.loc[idx, 'План проекта, шт.']}")
            #         st.write(f"  Факт проекта: {df.loc[idx, 'Факт проекта, шт.']}")
            #         st.write(f"  remaining_plan: {remaining_plan.loc[idx]}")
            #         st.write(f"  Дней до конца проекта: {df.loc[idx, 'Дней до конца проекта']}")
            #         st.write(f"  days_left (после replace): {days_left.loc[idx]}")
            #         st.write(f"  Ср. план на день: {df.loc[idx, 'Ср. план на день для 100% плана']}")
            #         st.write("---")
            # else:
            #     st.write("Нет проектов с планом ≥ 200")
            #  # === ОТЛАДКА ===
        
        # Расчет средней оплаты
        if 'Оплата факт' in df.columns and 'Факт проекта, шт.' in df.columns:
            mask = df['Факт проекта, шт.'] > 0
            df['Оплата факт средн., руб.'] = 0.0
            df.loc[mask, 'Оплата факт средн., руб.'] = (
                df.loc[mask, 'Оплата факт'] / df.loc[mask, 'Факт проекта, шт.']
            ).round(2)
            
        return df
        

# Глобальный экземпляр
visit_calculator = VisitCalculator()







