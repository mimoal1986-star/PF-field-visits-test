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
    def extract_base_data(self, field_projects_df):
        try:
            if field_projects_df is None or field_projects_df.empty:
                st.warning("⚠️ Нет полевых проектов для анализа")
                return pd.DataFrame()
            
            # Создаем DataFrame с нужными колонками
            result = pd.DataFrame()
            
            # Берем колонки напрямую из полевых проектов
            result['Код проекта'] = field_projects_df['Код проекта']
            result['Имя клиента'] = field_projects_df['Имя клиента']
            result['Название проекта'] = field_projects_df['Название проекта']
            result['ЗОД'] = field_projects_df['ЗОД']
            result['АСС'] = field_projects_df['АСС']
            result['ЭМ'] = field_projects_df['ЭМ']
            result['Регион short'] = field_projects_df['Регион short']
            result['Регион'] = field_projects_df['Регион']
            
            # Удаляем дубликаты по Названию проекта (как в Excel)
            result = result.drop_duplicates(subset=['Название проекта', 'Код проекта'], keep='first')
            
            st.info(f"📊 Извлечено базовых данных: {len(result)} уникальных полевых проектов")
            return result
            
        except KeyError as e:
            st.error(f"❌ В полевых проектах нет колонки: {e}")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"❌ Ошибка извлечения базовых данных: {str(e)[:100]}")
            return pd.DataFrame()
        # В utils/visit_calculator.py, после метода extract_base_data, вставляем:


    """Рассчитывает план на каждый этап, день проекта. Возвращает план на дату"""
    def calculate_plan_on_date_full(self, base_data, google_df, array_df, calc_params):
        """Рассчитывает 'План на дату, шт.' по полной логике."""
        
        result = base_data.copy()
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
        result['Факт проекта, шт.'] = 0  # ← Новая колонка
        result['Факт на дату, шт.'] = 0
        
        start_date = calc_params['start_date']
        end_date = calc_params['end_date']
        surrogate_date = pd.Timestamp('1900-01-01')
        
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            
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
                        for visit_date in stage_visits['Дата визита']:
                            if start_date <= visit_date.date() <= end_date:
                                result.at[idx, 'Факт на дату, шт.'] += 1
                        
                        day_pointer = stage_end + timedelta(days=1)
        
        return result
    
# Глобальный экземпляр
visit_calculator = VisitCalculator()






