# utils/visit_calculator.py
# draft 1.8
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

class VisitCalculator:
    
    def extract_base_data(self, field_projects_df, google_df_clean=None):
        """Извлекает базовые данные из полевых проектов (ВСЕХ, не только Чеккер)"""
        
        try:
            if field_projects_df is None or field_projects_df.empty:
                return pd.DataFrame()
            
            # Базовые колонки (всегда)
            base = pd.DataFrame()
            base['Код проекта'] = field_projects_df['Код проекта']
            base['Имя клиента'] = field_projects_df['Имя клиента']
            base['Название проекта'] = field_projects_df['Название проекта']
            base['ЗОД'] = field_projects_df['ЗОД']
            base['АСС'] = field_projects_df['АСС']
            base['ЭМ'] = field_projects_df['ЭМ']
            base['Регион short'] = field_projects_df['Регион short']
            base['Регион'] = field_projects_df['Регион']
            
            # ПО из исходных данных (если есть)
            if 'ПО' in field_projects_df.columns:
                base['ПО'] = field_projects_df['ПО']
            elif google_df_clean is not None and not google_df_clean.empty:
                # Пытаемся определить ПО из Google
                base['ПО'] = 'не определено'
                portal_col = 'Портал на котором идет проект (для работы полевой команды)'
                if portal_col in google_df_clean.columns:
                    # Для каждого проекта ищем ПО в Google
                    for idx, row in base.iterrows():
                        mask = (
                            google_df_clean['Код проекта RU00.000.00.01SVZ24'] == row['Код проекта']
                        )
                        if mask.any():
                            po_value = google_df_clean.loc[mask, portal_col].iloc[0]
                            if pd.notna(po_value):
                                base.at[idx, 'ПО'] = po_value
            else:
                base['ПО'] = 'не определено'
            
            # ====== ДОБАВЛЯЕМ ДАТЫ СТАРТА И ФИНИША ======
            base['Дата старта'] = None
            base['Дата финиша с продлением'] = None
            base['Длительность проекта, кол-во дней'] = 0
            
            if google_df_clean is not None and not google_df_clean.empty:
                # Колонки с датами в Google таблице
                start_col = 'Дата старта'
                end_col = 'Дата финиша с продлением'
                code_col = 'Код проекта RU00.000.00.01SVZ24'
                
                if start_col in google_df_clean.columns and end_col in google_df_clean.columns:
                    for idx, row in base.iterrows():
                        mask = google_df_clean[code_col] == row['Код проекта']
                        if mask.any():
                            # Берем ПЕРВУЮ найденную запись
                            start_date = pd.to_datetime(
                                google_df_clean.loc[mask, start_col].iloc[0], 
                                errors='coerce'
                            )
                            end_date = pd.to_datetime(
                                google_df_clean.loc[mask, end_col].iloc[0], 
                                errors='coerce'
                            )
                            
                            if pd.notna(start_date) and pd.notna(end_date):
                                base.at[idx, 'Дата старта'] = start_date
                                base.at[idx, 'Дата финиша с продлением'] = end_date
                                
                                # РАСЧЕТ ДЛИТЕЛЬНОСТИ ПРОЕКТА
                                duration_days = (end_date - start_date).days + 1
                                base.at[idx, 'Длительность проекта, кол-во дней'] = max(0, duration_days)
            # =============================================
            
            # ====== ПЕРЕСТАНОВКА КОЛОНОК ======
            # Желаемый порядок
            desired_order = [
                'Код проекта',
                'Имя клиента',
                'Название проекта',
                'ЗОД',
                'АСС',
                'ЭМ',
                'Регион short',
                'Регион',
                'ПО',
                'Дата старта',           # ← теперь эти колонки есть
                'Дата финиша с продлением',
                'Длительность проекта, кол-во дней'  # ← добавляем длительность тоже
            ]
            
            # Оставляем только существующие колонки
            existing_cols = [col for col in desired_order if col in base.columns]
            other_cols = [col for col in base.columns if col not in existing_cols]
            final_order = existing_cols + other_cols
            
            # Применяем новый порядок
            base = base[final_order]
            # =================================
            
            # Удаляем дубликаты и возвращаем
            base = base.drop_duplicates(subset=['Код проекта', 'Название проекта'], keep='first')
            return base
            
        except Exception:
            return pd.DataFrame()

    def _get_project_dates(self, project_code, google_df):
        """Находит даты проекта в Google ТОЛЬКО по коду"""
        try:
            # Ищем ВСЕ записи с этим кодом
            mask = google_df['Код проекта RU00.000.00.01SVZ24'] == project_code
            
            if not mask.any():
                return None, None
            
            # Берем все совпадающие строки
            matching_rows = google_df[mask]
            
            # Преобразуем даты
            start_dates = pd.to_datetime(matching_rows['Дата старта'], errors='coerce')
            end_dates = pd.to_datetime(matching_rows['Дата финиша с продлением'], errors='coerce')
            
            # Находим самую раннюю дату старта
            min_start = start_dates.min() if not start_dates.isna().all() else None
            
            # Находим самую позднюю дату финиша
            max_end = end_dates.max() if not end_dates.isna().all() else None
            
            return min_start, max_end
            
        except Exception:
            return None, None

    def _calculate_stages_plan(self, total_plan, duration_days, coefficients):
        """Рассчитывает план по этапам"""
        if total_plan == 0 or duration_days == 0:
            return [], []
        
        # Делим на 4 этапа
        stage_days = duration_days // 4
        extra_days = duration_days % 4
        
        stages_plan = []
        stages_days = []
        
        # Первые 3 этапа
        for i in range(3):
            days_in_stage = stage_days + (1 if i < extra_days else 0)
            stage_plan = total_plan * coefficients[i]
            stages_plan.append(stage_plan)
            stages_days.append(days_in_stage)
        
        # 4-й этап (остаток)
        days_in_stage = stage_days + (1 if 3 < extra_days else 0)
        stage_plan = total_plan - sum(stages_plan)
        stages_plan.append(stage_plan)
        stages_days.append(days_in_stage)
        
        return stages_plan, stages_days
    
    def calculate_plan_on_date_full(self, base_data, google_df, array_df, cxway_df, calc_params):
        """Рассчитывает 'План на дату, шт.' для всех проектов"""
        
        result = base_data.copy()
        result['План проекта, шт.'] = 0
        result['План на дату, шт.'] = 0.0
        result['Дата старта проекта'] = None
        result['Дата финиша проекта'] = None
        
        start_period = calc_params['start_date']
        end_period = calc_params['end_date']
        coeffs = calc_params['coefficients']
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            project_po = row['ПО']
            
            # 1. Ищем даты проекта в Google ТОЛЬКО ПО КОДУ
            start_date, end_date = self._get_project_dates(project_code, google_df)
            
            # ✅ СОХРАНЯЕМ ДАТЫ В РЕЗУЛЬТАТ
            result.at[idx, 'Дата старта проекта'] = start_date
            result.at[idx, 'Дата финиша проекта'] = end_date
            
            if start_date is None or end_date is None:
                # Если дат нет - пропускаем проект
                continue
            
            # 2. Считаем общий план проекта из всех источников
            total_plan = 0
            
            # ПЛАН из МАССИВА (для проектов на Чеккере)
            if project_po in ['Чеккер', 'не определено']:
                project_rows_array = array_df[
                    (array_df['Код анкеты'] == project_code) & 
                    (array_df['Название проекта'] == project_name)
                ]
                total_plan += len(project_rows_array)
            
            # ПЛАН из CXWAY (для проектов на CXWAY)
            if project_po in ['CXWAY', 'не определено'] and cxway_df is not None:
                project_rows_cxway = cxway_df[
                    (cxway_df['Project Code'] == project_code) &
                    (cxway_df['Project Name'] == project_name)
                ]
                total_plan += len(project_rows_cxway)
            
            if total_plan == 0:
                continue
                    
            result.at[idx, 'План проекта, шт.'] = total_plan
            
            # 3. Считаем длительность проекта
            duration_days = (end_date - start_date).days + 1
            
            # 4. Распределяем план по этапам
            stages_plan, stages_days = self._calculate_stages_plan(total_plan, duration_days, coeffs)
            
            # 5. Считаем план на дату
            plan_on_date = 0.0
            current_date = start_date
            
            for stage_idx in range(4):
                stage_plan = stages_plan[stage_idx]
                stage_days = stages_days[stage_idx]
                
                if stage_plan > 0 and stage_days > 0:
                    daily_plan = stage_plan / stage_days
                    
                    for day_offset in range(stage_days):
                        current_day = current_date + timedelta(days=day_offset)
                        
                        if start_period <= current_day.date() <= end_period:
                            plan_on_date += daily_plan
                
                current_date += timedelta(days=stage_days)
            
            result.at[idx, 'План на дату, шт.'] = round(plan_on_date, 1)
        
        return result
    
    def calculate_fact_on_date_full(self, base_data, google_df, array_df, cxway_df, calc_params):
        """Рассчитывает 'Факт на дату, шт.' и 'Факт проекта' (Массив + CXWAY)."""
        
        result = base_data.copy()
        result['Факт проекта, шт.'] = 0
        result['Факт на дату, шт.'] = 0
        
        # ✅ УБЕДИМСЯ ЧТО КОЛОНКИ ДАТ ЕСТЬ
        if 'Дата старта проекта' not in result.columns:
            result['Дата старта проекта'] = None
        if 'Дата финиша проекта' not in result.columns:
            result['Дата финиша проекта'] = None
        
        start_date_period = calc_params['start_date']
        end_date_period = calc_params['end_date']
        surrogate_date = pd.Timestamp('1900-01-01')
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            project_po = row['ПО']
            
            # ✅ ЕСЛИ ДАТЫ ЕЩЕ НЕ СОХРАНЕНЫ - ИЩЕМ ИХ
            if pd.isna(row['Дата старта проекта']) or pd.isna(row['Дата финиша проекта']):
                start_date, finish_date = self._get_project_dates(project_code, google_df)
                
                if start_date is not None and finish_date is not None:
                    result.at[idx, 'Дата старта проекта'] = start_date
                    result.at[idx, 'Дата финиша проекта'] = finish_date
                    
                    # Длительность проекта (для расчета)
                    duration_days = (finish_date - start_date).days + 1
                    result.at[idx, 'Длительность проекта, кол-во дней'] = duration_days
                else:
                    result.at[idx, 'Дата старта проекта'] = None
                    result.at[idx, 'Дата финиша проекта'] = None
                    result.at[idx, 'Длительность проекта, кол-во дней'] = 0
            
            # 1. Факт проекта из ВСЕХ источников
            fact_total = 0
            project_visits_array = None  # для расчета факта на дату
            
            # ФАКТ из МАССИВА (только для проектов на Чеккер или не определено ПО)
            if project_po in ['Чеккер', 'не определено']:
                # Определяем правильное название колонки статуса
                status_col_array = None
                if 'Статус' in array_df.columns:
                    status_col_array = 'Статус'
                elif 'Status' in array_df.columns:
                    status_col_array = 'Status'
                elif ' Статус' in array_df.columns:  # с пробелом в начале
                    status_col_array = ' Статус'
                
                if status_col_array:
                    project_visits_array = array_df[
                        (array_df['Код анкеты'] == project_code) &
                        (array_df['Название проекта'] == project_name) &
                        (array_df[status_col_array] == 'Выполнено')
                    ]
                    fact_total += len(project_visits_array)
                else:
                    # Если нет колонки статуса, используем старую логику (дата визита)
                    project_visits_array = array_df[
                        (array_df['Код анкеты'] == project_code) &
                        (array_df['Название проекта'] == project_name) &
                        (array_df['Дата визита'] != surrogate_date)
                    ]
                    fact_total += len(project_visits_array)
            
            # ФАКТ из CXWAY (только для проектов на CXWAY или не определено ПО)
            if project_po in ['CXWAY', 'не определено'] and cxway_df is not None:
                # Ищем записи в CXWAY со статусом 'Выполнено'
                status_col = None
                if 'Status' in cxway_df.columns:
                    status_col = 'Status'
                elif 'Статус' in cxway_df.columns:
                    status_col = 'Статус'
                
                if status_col:
                    project_visits_cxway = cxway_df[
                        (cxway_df['Project Code'] == project_code) &
                        (cxway_df['Project Name'] == project_name) &
                        (cxway_df[status_col] == 'Выполнено')
                    ]
                    fact_total += len(project_visits_cxway)
            
            if fact_total == 0:
                continue
                
            result.at[idx, 'Факт проекта, шт.'] = fact_total
            
            # 2. Распределяем факт по датам (только для массива)
            if project_po in ['Чеккер', 'не определено'] and project_visits_array is not None:
                # Находим даты проекта из Google (уже сохранены выше)
                start_date = result.at[idx, 'Дата старта проекта']
                finish_date = result.at[idx, 'Дата финиша проекта']
                
                if pd.notna(start_date) and pd.notna(finish_date):
                    # Те же 4 этапа что для плана
                    proj_duration = (finish_date - start_date).days + 1
                    stage_days = proj_duration // 4
                    extra_days = proj_duration % 4
                    
                    # Распределяем визиты по этапам
                    day_pointer = start_date
                    
                    for stage in range(4):
                        days_in_stage = stage_days + (1 if stage < extra_days else 0)
                        stage_end = day_pointer + timedelta(days=days_in_stage - 1)
                        
                        # Визиты в этом этапе
                        stage_visits = project_visits_array[
                            (project_visits_array['Дата визита'] >= day_pointer) &
                            (project_visits_array['Дата визита'] <= stage_end)
                        ]
                        
                        # Считаем визиты в периоде календаря
                        for _, visit_row in stage_visits.iterrows():
                            visit_date = visit_row['Дата визита']
                            if start_date_period <= visit_date.date() <= end_date_period:
                                result.at[idx, 'Факт на дату, шт.'] += 1
                        
                        day_pointer = stage_end + timedelta(days=1)
        
        # 3. Добавляем % после расчета факта
        result['%ПФ проекта'] = 0.0
        result['%ПФ на дату'] = 0.0
        
        for idx, row in result.iterrows():
            plan_project = row['План проекта, шт.']
            fact_project = row['Факт проекта, шт.']
            plan_date = row['План на дату, шт.']
            fact_date = row['Факт на дату, шт.']
            
            if plan_project > 0:
                result.at[idx, '%ПФ проекта'] = round((fact_project / plan_project) * 100, 1)
            
            if plan_date > 0:
                result.at[idx, '%ПФ на дату'] = round((fact_date / plan_date) * 100, 1)
    
        # 4. РАСЧЕТ ПРОГНОЗА НА МЕСЯЦ
        result['Прогноз на месяц, шт.'] = 0.0
        result['Прогноз на месяц, %'] = 0.0
        
        for idx, row in result.iterrows():
            pf_percent = row['%ПФ проекта']
            plan_project = row['План проекта, шт.']
            
            if pf_percent > 100:
                # Если %ПФ > 100%, прогноз = план
                result.at[idx, 'Прогноз на месяц, шт.'] = plan_project
                result.at[idx, 'Прогноз на месяц, %'] = 100.0
            elif pf_percent > 0:
                # Прогноз = %ПФ × План проекта
                forecast = (pf_percent / 100) * plan_project
                if forecast > 0:
                    result.at[idx, 'Прогноз на месяц, шт.'] = round(forecast, 1)
                result.at[idx, 'Прогноз на месяц, %'] = round(pf_percent, 1)
        
        # 5. ДОПОЛНИТЕЛЬНЫЕ РАСЧЕТЫ
        result['Кол-во визитов до 100% плана, шт.'] = 0
        result['Поручено'] = 0
        result['Доля Поручено, %'] = 0.0
        
        for idx, row in result.iterrows():
            plan_project = row['План проекта, шт.']
            fact_project = row['Факт проекта, шт.']
            
            # 1. Кол-во визитов до 100% плана
            visits_to_100 = max(0, plan_project - fact_project)
            result.at[idx, 'Кол-во визитов до 100% плана, шт.'] = visits_to_100
            
            # 2. Поручено - считаем из array_df где Статус == 'Поручено'
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            
            porucheno_count = 0
            if project_code in array_df['Код анкеты'].values:
                # Определяем колонку статуса для Поручено
                status_col_array = None
                if 'Статус' in array_df.columns:
                    status_col_array = 'Статус'
                elif ' Статус' in array_df.columns:
                    status_col_array = ' Статус'
                
                if status_col_array:
                    porucheno_mask = (
                        (array_df['Код анкеты'] == project_code) &
                        (array_df['Название проекта'] == project_name) &
                        (array_df[status_col_array] == 'Поручено')
                    )
                    porucheno_count = porucheno_mask.sum()
            
            result.at[idx, 'Поручено'] = porucheno_count
            
            # 3. Доля Поручено, %
            if visits_to_100 > 0 and porucheno_count > 0:
                result.at[idx, 'Доля Поручено, %'] = round((porucheno_count / visits_to_100) * 100, 1)
        
        # 6. РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ ПОКАЗАТЕЛЕЙ ПО ДНЯМ (используем сохраненные даты)
        result['Дней потрачено'] = 0
        result['Дней до конца проекта'] = 0
        result['Ср. план на день для 100% плана'] = 0.0
        
        for idx, row in result.iterrows():
            start_date = row['Дата старта проекта']
            finish_date = row['Дата финиша проекта']
            duration_days = row['Длительность проекта, кол-во дней']
            
            if pd.notna(start_date) and pd.notna(finish_date) and duration_days > 0:
                # Дней потрачено
                days_spent = (end_date_period - start_date.date()).days + 1
                result.at[idx, 'Дней потрачено'] = max(0, min(days_spent, duration_days))
                
                # Дней до конца проекта
                days_left = (finish_date.date() - end_date_period).days
                result.at[idx, 'Дней до конца проекта'] = max(0, days_left)
                
                # Средний план на день
                plan_project = row['План проекта, шт.']
                result.at[idx, 'Ср. план на день для 100% плана'] = round(plan_project / duration_days, 1)
        
        # 7. РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ ПОКАЗАТЕЛЕЙ
        result['Исполнение Проекта,%'] = 0.0
        result['Утилизация тайминга, %'] = 0.0
        result['Фокус'] = 'Нет'
        
        for idx, row in result.iterrows():
            fact_date = row['Факт на дату, шт.']
            plan_month = row['План проекта, шт.']
            days_spent = row['Дней потрачено']
            duration_days = row['Длительность проекта, кол-во дней']
            
            # Исполнение Проекта
            if plan_month > 0:
                result.at[idx, 'Исполнение Проекта,%'] = round((fact_date / plan_month) * 100, 1)
            
            # Утилизация тайминга
            if duration_days > 0:
                result.at[idx, 'Утилизация тайминга, %'] = round((days_spent / duration_days) * 100, 1)
            
            # Фокус
            if (row['Исполнение Проекта,%'] < 80 and 
                row['Утилизация тайминга, %'] > 80 and 
                row['Утилизация тайминга, %'] < 100):
                result.at[idx, 'Фокус'] = 'Да'
        
        # 8. ДЕЛЬТЫ
        result['△План/Факт на дату, шт.'] = result['План на дату, шт.'] - result['Факт на дату, шт.']
        result['△План/Факт проекта, шт.'] = result['План проекта, шт.'] - result['Факт проекта, шт.']
        
        return result

# Глобальный экземпляр
visit_calculator = VisitCalculator()



