# utils/visit_calculator.py
# draft 1.0
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
    def extract_base_data(self, field_projects_df, google_df_clean):
        """Извлекает базовые данные только для проектов на Чеккере"""
        
        # ========== ПРОВЕРКА КОЛОНОК ==========
        portal_col = 'Портал на котором идет проект (для работы полевой команды)'
        
        if portal_col not in google_df_clean.columns:
            # Если нет колонки ПО, возвращаем все проекты с меткой "Неизвестно"
            base = pd.DataFrame()
            base['Код проекта'] = field_projects_df['Код проекта']
            base['Имя клиента'] = field_projects_df['Имя клиента']
            base['Название проекта'] = field_projects_df['Название проекта']
            base['ПО'] = 'Неизвестно'  # Ставим заглушку
            base['ЗОД'] = field_projects_df['ЗОД']
            base['АСС'] = field_projects_df['АСС']
            base['ЭМ'] = field_projects_df['ЭМ']
            base['Регион short'] = field_projects_df['Регион short']
            base['Регион'] = field_projects_df['Регион']
            return base
        
        # ========== ФИЛЬТРАЦИЯ ПО ЧЕККЕРУ ==========
        checker_mask = google_df_clean[portal_col] == 'Чеккер'
        checker_projects = google_df_clean[checker_mask]
        
        # Создаем базовый датафрейм
        base = pd.DataFrame()
        
        # Базовые колонки из полевых проектов
        base['Код проекта'] = field_projects_df['Код проекта']
        base['Имя клиента'] = field_projects_df['Имя клиента']
        base['Название проекта'] = field_projects_df['Название проекта']
        
        # Создаем ключ для связывания
        base['key'] = base['Код проекта'] + '|' + base['Название проекта']
        checker_projects['key'] = (
            checker_projects['Код проекта RU00.000.00.01SVZ24'] + '|' + 
            checker_projects['Название волны на Чекере/ином ПО']
        )
        
        # Добавляем ПО
        po_mapping = checker_projects.set_index('key')[portal_col].to_dict()
        base['ПО'] = base['key'].map(po_mapping)
        
        # Фильтруем только проекты на Чеккере
        base = base[base['ПО'] == 'Чеккер']
        
        # Добавляем остальные колонки
        base['ЗОД'] = field_projects_df['ЗОД']
        base['АСС'] = field_projects_df['АСС']
        base['ЭМ'] = field_projects_df['ЭМ']
        base['Регион short'] = field_projects_df['Регион short']
        base['Регион'] = field_projects_df['Регион']
        
        # Удаляем временный ключ
        base = base.drop('key', axis=1)
        
        return base


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
        
        return result 
    
# Глобальный экземпляр
visit_calculator = VisitCalculator()









