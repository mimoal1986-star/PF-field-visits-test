# utils/visit_calculator.py
# draft 1.9 
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

class VisitCalculator:
    
    def extract_base_data(self, field_projects_df, google_df_clean=None):
        """Извлекает базовые данные из полевых проектов с ПРАВИЛЬНЫМИ датами"""
        
        try:
            if field_projects_df is None or field_projects_df.empty:
                return pd.DataFrame()
            
            # 1. Создаем базовую таблицу
            base = pd.DataFrame()
            base['Код проекта'] = field_projects_df['Код проекта']
            base['Имя клиента'] = field_projects_df['Имя клиента']
            base['Название проекта'] = field_projects_df['Название проекта']
            base['ПО'] = field_projects_df['ПО']
            
            # 2. Добавляем колонки для дат (пока пустые)
            base['Дата старта'] = None
            base['Дата финиша с продлением'] = None
            base['Длительность проекта, кол-во дней'] = 0
            
            # 3. Если Google таблица есть - заполняем даты ПРАВИЛЬНО
            if google_df_clean is not None and not google_df_clean.empty:
                # Убедимся, что названия колонок правильные
                google_code_col = 'Код проекта RU00.000.00.01SVZ24'
                start_col = 'Дата старта'
                end_col = 'Дата финиша с продлением'
                
                if all(col in google_df_clean.columns for col in [google_code_col, start_col, end_col]):
                    
                    # 4. Для КАЖДОГО проекта находим ВСЕ записи в Google
                    for idx, row in base.iterrows():
                        project_code = str(row['Код проекта']).strip()
                        
                        # Находим ВСЕ строки с этим кодом
                        mask = google_df_clean[google_code_col].astype(str).str.strip() == project_code
                        
                        if mask.any():
                            # Берем ВСЕ совпадающие строки
                            matching_rows = google_df_clean[mask]
                            
                            # 5. Находим САМУЮ РАННЮЮ дату старта
                            all_start_dates = pd.to_datetime(
                                matching_rows[start_col], 
                                errors='coerce'
                            )
                            earliest_start = all_start_dates.min()  # самая ранняя
                            
                            # 6. Находим САМУЮ ПОЗДНЮЮ дату финиша  
                            all_end_dates = pd.to_datetime(
                                matching_rows[end_col], 
                                errors='coerce'
                            )
                            latest_end = all_end_dates.max()  # самая поздняя
                            
                            # 7. Сохраняем если обе даты найдены
                            if pd.notna(earliest_start) and pd.notna(latest_end):
                                base.at[idx, 'Дата старта'] = earliest_start
                                base.at[idx, 'Дата финиша с продлением'] = latest_end
                                
                                # 8. Считаем длительность
                                duration_days = (latest_end - earliest_start).days + 1
                                base.at[idx, 'Длительность проекта, кол-во дней'] = max(0, duration_days)
                                
            # 9. Удаляем дубликаты
            base = base.drop_duplicates(subset=['Код проекта', 'Название проекта'], keep='first')
            
            return base
            
        except Exception as e:
            st.error(f"❌ Ошибка в extract_base_data: {str(e)[:200]}")
            return pd.DataFrame()


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
    
    def calculate_plan_on_date_full(self, base_data, array_df, cxway_df, calc_params):
        """Рассчитывает 'План на дату, шт.' для всех проектов - ИСПРАВЛЕННАЯ ВЕРСИЯ"""
        
        result = base_data.copy()
        result['План проекта, шт.'] = 0
        result['План на дату, шт.'] = 0.0
        
        start_period = calc_params['start_date']
        end_period = calc_params['end_date']
        coeffs = calc_params['coefficients']
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            project_po = row['ПО']
            
            # ✅ ИСПРАВЛЕНИЕ: Берем даты ИЗ БАЗОВЫХ ДАННЫХ
            start_date = row.get('Дата старта')
            end_date = row.get('Дата финиша с продлением')
            
            if pd.isna(start_date) or pd.isna(end_date):
                # Если дат нет - пропускаем проект
                continue
            
            # ✅ Длительность уже рассчитана в base_data
            duration_days = row.get('Длительность проекта, кол-во дней', 0)
            
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
            
            # 3. Используем длительность из base_data (уже рассчитана)
            if duration_days <= 0:
                continue
                
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
    
    def calculate_fact_on_date_full(self, base_data, array_df, cxway_df, calc_params):
        """Рассчитывает 'Факт на дату, шт.' и 'Факт проекта' (Массив + CXWAY)."""
        
        result = base_data.copy()
        result['Факт проекта, шт.'] = 0
        result['Факт на дату, шт.'] = 0
        
        # ✅ УБЕДИМСЯ ЧТО КОЛОНКИ ДАТ ЕСТЬ (КОСЯК!!!)
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
            
            # ✅ ИСПРАВЛЕНИЕ: Берем даты ИЗ БАЗОВЫХ ДАННЫХ (из extract_base_data)
            start_date = row.get('Дата старта')
            finish_date = row.get('Дата финиша с продлением')
            duration_days = row.get('Длительность проекта, кол-во дней', 0)
            
            # Сохраняем в старые колонки для обратной совместимости
            if pd.notna(start_date):
                result.at[idx, 'Дата старта проекта'] = start_date
            if pd.notna(finish_date):
                result.at[idx, 'Дата финиша проекта'] = finish_date
            
            # Если дат нет в новых колонках, используем старые
            if pd.isna(start_date):
                start_date = row.get('Дата старта проекта')
            if pd.isna(finish_date):
                finish_date = row.get('Дата финиша проекта')
            if duration_days == 0:
                if pd.notna(start_date) and pd.notna(finish_date):
                    duration_days = (finish_date - start_date).days + 1
            
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
            
            # 2. ФАКТ НА ДАТУ: визиты от 1-го числа месяца до вчера
            if project_po in ['Чеккер', 'не определено'] and project_visits_array is not None:
                # Даты из extract_base_data
                if pd.isna(start_date):
                    continue
                
                # 1. Первый день месяца проекта
                first_day = pd.Timestamp(start_date.year, start_date.month, 1)
                
                # 2. Вчерашний день
                yesterday = pd.Timestamp.now() - pd.Timedelta(days=1)
                
                # 3. Последний день месяца проекта
                last_day = pd.Timestamp(start_date.year, start_date.month, 1) + pd.offsets.MonthEnd(1)
                
                # 4. Ограничиваем период
                end_date = yesterday if yesterday <= last_day else last_day
                
                # 5. Считаем визиты в этом периоде
                period_visits = project_visits_array[
                    (project_visits_array['Дата визита'] >= first_day) &
                    (project_visits_array['Дата визита'] <= end_date)
                ]
                
                result.at[idx, 'Факт на дату, шт.'] = len(period_visits)
        
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
            # Используем даты из base_data (новые или старые колонки)
            start_date = row.get('Дата старта') or row.get('Дата старта проекта')
            finish_date = row.get('Дата финиша с продлением') or row.get('Дата финиша проекта')
            duration_days = row.get('Длительность проекта, кол-во дней', 0)
            
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










