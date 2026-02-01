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
            
            # Удаляем дубликаты и возвращаем
            base = base.drop_duplicates(subset=['Код проекта', 'Название проекта'], keep='first')
            return base
            
        except Exception:
            return pd.DataFrame()

    def _get_project_dates(self, project_code, project_name, google_df):
        """Находит даты проекта в Google или возвращает None"""
        try:
            google_mask = (
                (google_df['Код проекта RU00.000.00.01SVZ24'] == project_code) &
                (google_df['Название волны на Чекере/ином ПО'] == project_name)
            )
            
            if google_mask.any():
                start_date = pd.to_datetime(google_df.loc[google_mask, 'Дата старта'].iloc[0])
                end_date = pd.to_datetime(google_df.loc[google_mask, 'Дата финиша с продлением'].iloc[0])
                return start_date, end_date
            return None, None
        except:
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

    def calculate_plan_on_date_full(self, base_data, google_df, array_df, calc_params):
        """Рассчитывает 'План на дату, шт.' для всех проектов"""
        
        result = base_data.copy()
        result['План проекта, шт.'] = 0
        result['План на дату, шт.'] = 0.0
        
        start_period = calc_params['start_date']
        end_period = calc_params['end_date']
        coeffs = calc_params['coefficients']
        
        # Считаем план для каждого проекта
        for idx, row in result.iterrows():
            project_code = row['Код проекта']
            project_name = row['Название проекта']
            
            # 1. Ищем даты проекта в Google
            start_date, end_date = self._get_project_dates(project_code, project_name, google_df)
            
            if start_date is None or end_date is None:
                # Проекта нет в Google - пропускаем
                continue
            
            # 2. Считаем общий план проекта
            total_plan = 0
            
            # Смотрим в массиве (для проектов на Чеккере или не определено ПО)
            if row['ПО'] in ['Чеккер', 'не определено']:
                project_rows = array_df[
                    (array_df['Код анкеты'] == project_code) & 
                    (array_df['Название проекта'] == project_name)
                ]
                total_plan += len(project_rows)
            
            # TODO: Здесь можно добавить логику для CXWAY
            # if row['ПО'] in ['CXWAY', 'не определено']:
            #     cxway_plan = len(cxway_rows)  # Нужен доступ к данным CXWAY
            #     total_plan += cxway_plan
            
            if total_plan == 0:
                continue
                
            result.at[idx, 'План проекта, шт.'] = total_plan
            
            # 3. Считаем длительность проекта
            duration_days = (end_date - start_date).days + 1
            
            # 4. Распределяем план по этапам
            stages_plan, stages_days = self._calculate_stages_plan(total_plan, duration_days, coeffs)
            
            # 5. Считаем план на дату (распределение по дням)
            plan_on_date = 0.0
            current_date = start_date
            
            for stage_idx in range(4):
                stage_plan = stages_plan[stage_idx]
                stage_days = stages_days[stage_idx]
                
                if stage_plan > 0 and stage_days > 0:
                    daily_plan = stage_plan / stage_days
                    
                    # Для каждого дня этапа
                    for day_offset in range(stage_days):
                        current_day = current_date + timedelta(days=day_offset)
                        
                        # Если день в периоде расчета
                        if start_period <= current_day.date() <= end_period:
                            plan_on_date += daily_plan
                
                current_date += timedelta(days=stage_days)
            
            result.at[idx, 'План на дату, шт.'] = round(plan_on_date, 1)
        
        return result
    
    def calculate_plan_on_date_full(self, base_data, google_df, array_df, cxway_df, calc_params):
    """Рассчитывает 'План на дату, шт.' для всех проектов (Массив + CXWAY)"""
    
    result = base_data.copy()
    result['План проекта, шт.'] = 0
    result['План на дату, шт.'] = 0.0
    
    start_period = calc_params['start_date']
    end_period = calc_params['end_date']
    coeffs = calc_params['coefficients']
    
    # Считаем план для каждого проекта
    for idx, row in result.iterrows():
        project_code = row['Код проекта']
        project_name = row['Название проекта']
        project_po = row['ПО']
        
        # 1. Ищем даты проекта в Google
        start_date, end_date = self._get_project_dates(project_code, project_name, google_df)
        
        if start_date is None or end_date is None:
            # Проекта нет в Google - пропускаем
            continue
        
        # 2. Считаем общий план проекта из всех источников
        total_plan = 0
        
        # ПЛАН из МАССИВА (только для проектов на Чеккере или не определено ПО)
        if project_po in ['Чеккер', 'не определено']:
            project_rows_array = array_df[
                (array_df['Код анкеты'] == project_code) & 
                (array_df['Название проекта'] == project_name)
            ]
            total_plan += len(project_rows_array)
        
        # ПЛАН из CXWAY (только для проектов на CXWAY или не определено ПО)
        if project_po in ['CXWAY', 'не определено'] and cxway_df is not None:
            project_rows_cxway = cxway_df[
                (cxway_df['Код проекта'] == project_code) &
                (cxway_df['Название проекта'] == project_name)
            ]
            total_plan += len(project_rows_cxway)
        
        if total_plan == 0:
            continue
            
        result.at[idx, 'План проекта, шт.'] = total_plan
        
        # 3. Считаем длительность проекта
        duration_days = (end_date - start_date).days + 1
        
        # 4. Распределяем план по этапам
        stages_plan, stages_days = self._calculate_stages_plan(total_plan, duration_days, coeffs)
        
        # 5. Считаем план на дату (распределение по дням)
        plan_on_date = 0.0
        current_date = start_date
        
        for stage_idx in range(4):
            stage_plan = stages_plan[stage_idx]
            stage_days = stages_days[stage_idx]
            
            if stage_plan > 0 and stage_days > 0:
                daily_plan = stage_plan / stage_days
                
                # Для каждого дня этапа
                for day_offset in range(stage_days):
                    current_day = current_date + timedelta(days=day_offset)
                    
                    # Если день в периоде расчета
                    if start_period <= current_day.date() <= end_period:
                        plan_on_date += daily_plan
            
            current_date += timedelta(days=stage_days)
        
        result.at[idx, 'План на дату, шт.'] = round(plan_on_date, 1)
    
    return result
 
# Глобальный экземпляр
visit_calculator = VisitCalculator()

