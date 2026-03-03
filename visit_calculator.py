# utils/visit_calculator.py
# draft 4.0 - simplified
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

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
    
    
    def extract_hierarchical_data(self, visits_df, google_df=None):
        """
        Создаёт полную иерархию Проект→Клиент→Волна→Регион→DSM→ASM→RS
        с базовой информацией о проекте
        """
        
        try:
            # 1. Определяем колонку региона
            region_col = 'Регион short'
            
            # Создаём иерархию из visits_df (уникальные цепочки)
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
            
            # ТОЛЬКО ПОЛЕВЫЕ ПРОЕКТЫ
            hierarchy = hierarchy[hierarchy['Полевой'] == 1]
            hierarchy = hierarchy.drop('Полевой', axis=1)
            
            # Удаляем дубликаты
            hierarchy = hierarchy.drop_duplicates().reset_index(drop=True)
            
            # Даты - по умолчанию пустые
            hierarchy['Дата старта'] = pd.NaT
            hierarchy['Дата финиша'] = pd.NaT
            
            # Обогащаем датами из google_df
            if google_df is not None and not google_df.empty:
                try:
                    # Маппинги для дат
                    start_mapping = {}
                    finish_mapping = {}
                    
                    for idx, row in google_df.iterrows():
                        code = str(row.get('Код проекта RU00.000.00.01SVZ24', '')).strip()
                        if code and code not in ['nan', '']:
                            start_date = row.get('Дата старта')
                            finish_date = row.get('Дата финиша с продлением')
                            
                            if pd.notna(start_date):
                                start_mapping[code] = start_date
                            if pd.notna(finish_date):
                                finish_mapping[code] = finish_date
                    
                    hierarchy['Дата старта'] = hierarchy['Проект'].map(start_mapping)
                    hierarchy['Дата финиша'] = hierarchy['Проект'].map(finish_mapping)
                    
                except Exception as e:
                    pass
            
            # Добавляем информацию о квоте для Easymerch Мултон
            if 'План_квота' in visits_df.columns:
                # Создаем маппинг квот по уникальным комбинациям
                kvota_mapping = {}
                for idx, row in visits_df.iterrows():
                    if row['ПО'] == 'Easymerch' and row['Имя клиента'] == 'Мултон':
                        key = (row['Код анкеты'], row['Название проекта'], row['Регион short'])
                        kvota = row.get('План_квота', 0)
                        if kvota > 0:
                            kvota_mapping[key] = kvota
                
                # Добавляем колонку с квотой в иерархию
                hierarchy['План_квота'] = 0
                for idx, row in hierarchy.iterrows():
                    key = (row['Проект'], row['Волна'], row['Регион'])
                    if key in kvota_mapping:
                        hierarchy.at[idx, 'План_квота'] = kvota_mapping[key]
            else:
                hierarchy['План_квота'] = 0
            
            # Рассчитываем длительность
            hierarchy['Длительность'] = 0
            mask_valid_dates = hierarchy['Дата старта'].notna() & hierarchy['Дата финиша'].notna()
            
            if mask_valid_dates.any():
                hierarchy.loc[mask_valid_dates, 'Длительность'] = (
                    hierarchy.loc[mask_valid_dates, 'Дата финиша'] - 
                    hierarchy.loc[mask_valid_dates, 'Дата старта']
                ).dt.days + 1
            
            # Сортируем
            hierarchy = hierarchy.sort_values(['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM', 'RS'])
            hierarchy = hierarchy[hierarchy['RS'] != 'Итого']
            
            return hierarchy
            
        except KeyError as e:
            return pd.DataFrame()
            
        except Exception as e:
            return pd.DataFrame()
    
    def calculate_hierarchical_plan_on_date(self, hierarchy_df, visits_df, calc_params):
        """
        РАССЧИТЫВАЕТ ПЛАН ТОЛЬКО ДЛЯ УРОВНЯ RS
        """
        try:
            if hierarchy_df.empty or visits_df.empty:
                return pd.DataFrame()
            
            start_period = calc_params['start_date']
            end_period = calc_params['end_date']
            coefficients = calc_params['coefficients']
            
            # Планы проектов+волн+регионов (для обычных проектов)
            project_wave_region_plans = visits_df.groupby([
                'Код анкеты', 
                'Название проекта',
                'Регион short'
            ]).size()
            
            results = []
            
            for _, row in hierarchy_df.iterrows():
                region = row['Регион']
                project_code = row['Проект']
                wave_name = row['Волна']
                po = row['ПО']
                client = row['Клиент']
                
                # ОСОБЫЙ СЛУЧАЙ: Easymerch Мултон - план из квоты
                if po == 'Easymerch' and client == 'Мултон' and 'План_квота' in hierarchy_df.columns:
                    total_plan = row['План_квота']
                    if total_plan <= 0:
                        continue
                else:
                    # Обычный случай - план из количества визитов
                    plan_key = (project_code, wave_name, region)
                    if plan_key not in project_wave_region_plans.index:
                        continue
                    total_plan = project_wave_region_plans.loc[plan_key]
                
                # Проверка дат
                start_date = row['Дата старта']
                finish_date = row['Дата финиша']
                duration = row['Длительность']
                
                if pd.isna(start_date) or pd.isna(finish_date) or duration <= 0:
                    continue
                
                # Проверка пересечения с периодом
                if end_period < start_date.date() or start_period > finish_date.date():
                    continue
                
                # ДНИ ПРОЕКТА, ПОПАДАЮЩИЕ В ПЕРИОД
                days_in_period = 0
                current_date = start_date
                for day in range(duration):
                    check_date = current_date + timedelta(days=day)
                    if start_period <= check_date.date() <= end_period:
                        days_in_period += 1
                
                if days_in_period == 0:
                    continue
                
                # ДНЕВНОЙ ПЛАН ВОЛНЫ (равномерное распределение)
                daily_plan_wave = total_plan / duration
                
                # ДОЛИ RS
                rs_weights = self._calculate_rs_weights(visits_df, project_code, wave_name, region)
                rs_name = row['RS']
                
                if rs_name not in rs_weights or rs_weights[rs_name] <= 0:
                    continue
                
                rs_weight = rs_weights[rs_name]
                
                # Дневной план RS = дневной план волны × доля RS
                rs_daily_plan = daily_plan_wave * rs_weight
                
                # План RS на дату = дневной план × дни в периоде
                rs_plan_on_date = rs_daily_plan * days_in_period
                
                # Запись результата
                results.append({
                    'Проект': project_code,
                    'Клиент': row['Клиент'],
                    'Волна': wave_name,
                    'Регион': row['Регион'],
                    'DSM': row['DSM'],
                    'ASM': row['ASM'],
                    'RS': rs_name,
                    'ПО': po,
                    'Уровень': 'RS',
                    'План проекта, шт.': total_plan,
                    'План на дату, шт.': round(rs_plan_on_date, 1),
                    'Длительность': int(duration),
                    'Дата старта': start_date,
                    'Дата финиша': finish_date,
                    'Дней в периоде': days_in_period,
                    'Дневной план RS, шт.': round(rs_daily_plan, 2)
                })
                            
            if not results:
                return pd.DataFrame()
            return pd.DataFrame(results)
            
        except Exception as e:
            return pd.DataFrame()
        
    def calculate_hierarchical_fact_on_date(self, plan_df, visits_df, calc_params):
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
            
            # ФИЛЬТРЫ
            completed_mask = visits_df[status_col] == 'Выполнено'
            start_date = pd.Timestamp(calc_params['start_date'])
            end_date = pd.Timestamp(calc_params['end_date'])
            period_mask = (
                (visits_df['Дата визита'] >= start_date) &
                (visits_df['Дата визита'] <= end_date)
            )
            
            # СЧИТАЕМ ФАКТЫ
            completed_df = visits_df[completed_mask]
            rs_facts_total = completed_df.groupby([
                'Код анкеты',
                'Название проекта',
                region_col,
                rs_col
            ]).size().to_dict()
            
            completed_in_period = visits_df[completed_mask & period_mask]
            rs_facts_period = completed_in_period.groupby([
                'Код анкеты',
                'Название проекта',
                region_col,
                rs_col
            ]).size().to_dict()
            
            # ✅ СОЗДАЁМ КОЛОНКИ
            result_df['Факт проекта, шт.'] = 0
            result_df['Факт на дату, шт.'] = 0
            
            # ✅ ЗАПОЛНЯЕМ
            for idx in result_df[result_df['Уровень'] == 'RS'].index:
                row = result_df.loc[idx]
                project = str(row['Проект']).strip()
                wave = str(row['Волна']).strip()
                region = str(row['Регион']).strip()
                rs = str(row['RS']).strip()
                
                key = (project, wave, region, rs)
                result_df.at[idx, 'Факт проекта, шт.'] = rs_facts_total.get(key, 0)
                result_df.at[idx, 'Факт на дату, шт.'] = rs_facts_period.get(key, 0)
    
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
        if plan_df is not None and 'План на дату, шт.' in plan_df.columns:
            df['План на дату, шт.'] = plan_df['План на дату, шт.']
        elif 'План на дату, шт.' not in df.columns:
            df['План на дату, шт.'] = 0
        
        # 2. ФАКТ
        if 'Факт на дату, шт.' not in df.columns:
            df['Факт на дату, шт.'] = 0
        
        # 3. БАЗОВЫЕ МЕТРИКИ
        mask = df['План на дату, шт.'] > 0
        
        # % выполнения плана
        df['%ПФ на дату'] = 0.0
        if mask.any():
            df.loc[mask, '%ПФ на дату'] = (
                df.loc[mask, 'Факт на дату, шт.'] / 
                df.loc[mask, 'План на дату, шт.'] * 100
            ).round(1)
        
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
            
            # Средний план на день
            df['Ср. план на день для 100% плана'] = 0.0
            if mask_duration.any() and mask.any():
                remaining_plan = df.loc[mask & mask_duration, 'План на дату, шт.'] - df.loc[mask & mask_duration, 'Факт на дату, шт.']
                days_left = df.loc[mask & mask_duration, 'Дней до конца проекта'].replace(0, 1)
                df.loc[mask & mask_duration, 'Ср. план на день для 100% плана'] = (remaining_plan / days_left).round(2)
        
        return df
        

# Глобальный экземпляр
visit_calculator = VisitCalculator()





























































