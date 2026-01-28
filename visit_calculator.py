# utils/visit_calculator.py
# draft 1.6
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

# ==============================================
# Расчет ПФ визитов
# ==============================================

"""Калькулятор плана/факта визитов (воспроизводит логику 'Датасет ПФ1')"""
class VisitCalculator:
    
    """Извлекает базовые данные ТОЛЬКО из полевых проектов (столбцы A-H)"""
    def extract_base_data(self, field_projects_df, google_df_clean=None):
        """Извлекает базовые данные с фильтрацией по ПО (Чеккер)"""
        
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
            
            # Фильтрация по ПО (если есть google_df_clean)
            if google_df_clean is not None and not google_df_clean.empty:
                portal_col = 'Портал на котором идет проект (для работы полевой команды)'
                
                if portal_col in google_df_clean.columns:
                    # Только проекты на Чеккере
                    checker_mask = google_df_clean[portal_col] == 'Чеккер'
                    checker_df = google_df_clean[checker_mask]
                    
                    # Связывание по коду и названию
                    checker_codes = set(checker_df['Код проекта RU00.000.00.01SVZ24'])
                    checker_names = set(checker_df['Название волны на Чекере/ином ПО'])
                    
                    base = base[
                        base['Код проекта'].isin(checker_codes) & 
                        base['Название проекта'].isin(checker_names)
                    ]
                
                # Добавляем колонку ПО
                base['ПО'] = 'Чеккер'
            
            # Удаляем дубликаты и возвращаем
            base = base.drop_duplicates(subset=['Код проекта', 'Название проекта'], keep='first')
            return base
            
        except Exception:
            return pd.DataFrame()


    """Рассчитывает план на каждый этап, день проекта. Возвращает план на дату"""
    def calculate_plan_on_date_full(self, base_data, google_df, array_df, calc_params):
        """Рассчитывает 'План на дату, шт.' по полной логике."""
        
        result = base_data.copy()
        result['План проекта, шт.'] = 0
        result['План на дату, шт.'] = 0.0
        
        start_date = calc_params['start_date']
        end_date = calc_params['end_date']
        coeffs = calc_params['coefficients']
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            
            # 1. План проекта = кол-во строк в массиве
            project_rows = array_df[
                (array_df['Код анкеты'] == project_code) & 
                (array_df['Название проекта'] == project_name)
            ]
            plan_total = len(project_rows)
            result.at[idx, 'План проекта, шт.'] = plan_total
            
            # 2. Даты проекта из google
            google_mask = (
                (google_df['Код проекта RU00.000.00.01SVZ24'] == project_code) &
                (google_df['Название волны на Чекере/ином ПО'] == project_name)
            )
            
            if google_mask.any():
                proj_start = pd.to_datetime(google_df.loc[google_mask, 'Дата старта'].iloc[0])
                proj_end = pd.to_datetime(google_df.loc[google_mask, 'Дата финиша с продлением'].iloc[0])
                
                # 3. 4 этапа (равные отрезки)
                proj_duration = (proj_end - proj_start).days + 1
                stage_days = proj_duration // 4
                extra_days = proj_duration % 4
                
                stages = []
                day_pointer = proj_start
                
                for i in range(4):
                    days_in_stage = stage_days + (1 if i < extra_days else 0)
                    stage_end = day_pointer + timedelta(days=days_in_stage - 1)
                    
                    # 4. План этапов 1-3
                    if i < 3:
                        stage_plan = plan_total * coeffs[i]
                    else:  # Этап 4
                        stage_plan = plan_total - sum(stages)
                    
                    stages.append(stage_plan)
                    
                    # 5. План по дням (равномерно)
                    daily_plan = stage_plan / days_in_stage
                    
                    # 6. План на дату = сумма за период
                    for day_offset in range(days_in_stage):
                        current_day = day_pointer + timedelta(days=day_offset)
                        if start_date <= current_day.date() <= end_date:
                            result.at[idx, 'План на дату, шт.'] += daily_plan
                    
                    day_pointer = stage_end + timedelta(days=1)
        
        result['План на дату, шт.'] = result['План на дату, шт.'].round(1)
        return result
    
    def calculate_fact_on_date_full(self, base_data, google_df, array_df, calc_params):
        """Рассчитывает 'Факт на дату, шт.' и 'Факт проекта'."""
        
        result = base_data.copy()
        result['Факт проекта, шт.'] = 0
        result['Факт на дату, шт.'] = 0
        
        start_date = calc_params['start_date']
        end_date = calc_params['end_date']
        surrogate_date = pd.Timestamp('1900-01-01')
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']  # ← ДОБАВИЛИ
            project_name = row['Название проекта']  # ← ДОБАВИЛИ
            
            # Все фактические визиты проекта
            project_visits = array_df[
                (array_df['Код анкеты'] == project_code) &
                (array_df['Название проекта'] == project_name) &
                (array_df['Дата визита'] != surrogate_date)
            ]
            
            # 1. Факт проекта (все визиты)
            fact_total = len(project_visits)
            result.at[idx, 'Факт проекта, шт.'] = fact_total
            
            if fact_total > 0:
                # 2. Даты проекта из google (те же что для плана)
                google_mask = (
                    (google_df['Код проекта RU00.000.00.01SVZ24'] == project_code) &
                    (google_df['Название волны на Чекере/ином ПО'] == project_name)
                )
                
                if google_mask.any():
                    proj_start = pd.to_datetime(google_df.loc[google_mask, 'Дата старта'].iloc[0])
                    proj_end = pd.to_datetime(google_df.loc[google_mask, 'Дата финиша с продлением'].iloc[0])
                    
                    # 3. Те же 4 этапа что для плана
                    proj_duration = (proj_end - proj_start).days + 1
                    stage_days = proj_duration // 4
                    extra_days = proj_duration % 4
                    
                    # 4. Распределяем визиты по этапам
                    day_pointer = proj_start
                    
                    for stage in range(4):
                        days_in_stage = stage_days + (1 if stage < extra_days else 0)
                        stage_end = day_pointer + timedelta(days=days_in_stage - 1)
                        
                        # 5. Визиты в этом этапе
                        stage_visits = project_visits[
                            (project_visits['Дата визита'] >= day_pointer) &
                            (project_visits['Дата визита'] <= stage_end)
                        ]
                        
                        # 6. Считаем визиты в периоде календаря
                        for _, visit_row in stage_visits.iterrows():
                            visit_date = visit_row['Дата визита']
                            if start_date <= visit_date.date() <= end_date:
                                result.at[idx, 'Факт на дату, шт.'] += 1
                        
                        day_pointer = stage_end + timedelta(days=1)
        
        # 7. Добавляем % после расчета факта
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
    
        # 8. РАСЧЕТ ПРОГНОЗА НА МЕСЯЦ
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
            # Если %ПФ = 0 или отрицательный, прогноз остается 0
        
        # 9. ДОПОЛНИТЕЛЬНЫЕ РАСЧЕТЫ
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
                porucheno_mask = (
                    (array_df['Код анкеты'] == project_code) &
                    (array_df['Название проекта'] == project_name) &
                    (array_df[' Статус'] == 'Поручено')
                )
                porucheno_count = porucheno_mask.sum()
            
            result.at[idx, 'Поручено'] = porucheno_count
            
            # 3. Доля Поручено, %
            if visits_to_100 > 0 and porucheno_count > 0:
                result.at[idx, 'Доля Поручено, %'] = round((porucheno_count / visits_to_100) * 100, 1)
        
        # 10. РАСЧЕТ ПОКАЗАТЕЛЕЙ ПО ДНЯМ
        result['Длительность проекта, кол-во дней'] = 0
        result['Дней потрачено'] = 0
        result['Дней до конца проекта'] = 0
        result['Ср. план на день для 100% плана'] = 0.0
        
        # Даты из параметров
        end_date_period = calc_params['end_date']
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            
            # Находим даты проекта из google_df
            google_mask = (
                (google_df['Код проекта RU00.000.00.01SVZ24'] == project_code) &
                (google_df['Название волны на Чекере/ином ПО'] == project_name)
            )
            
            if google_mask.any():
                start_date = pd.to_datetime(google_df.loc[google_mask, 'Дата старта'].iloc[0])
                finish_date = pd.to_datetime(google_df.loc[google_mask, 'Дата финиша с продлением'].iloc[0])
                
                # Длительность проекта
                duration_days = (finish_date - start_date).days + 1
                result.at[idx, 'Длительность проекта, кол-во дней'] = duration_days
                
                # Дней потрачено
                days_spent = (end_date_period - start_date.date()).days + 1
                result.at[idx, 'Дней потрачено'] = max(0, min(days_spent, duration_days))
                
                # Дней до конца проекта
                days_left = (finish_date.date() - end_date_period).days
                result.at[idx, 'Дней до конца проекта'] = max(0, days_left)
                
                # Средний план на день
                plan_project = row['План проекта, шт.']
                if duration_days > 0:
                    result.at[idx, 'Ср. план на день для 100% плана'] = round(plan_project / duration_days, 1)
        
        return result 
    
# Глобальный экземпляр
visit_calculator = VisitCalculator()















