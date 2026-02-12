# utils/visit_calculator.py
# draft 2.0 
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

class VisitCalculator:
    
    def _calculate_rs_weights(self, array_df, project_code, wave_name, region):
        """
        Доли RS = визиты RS в проекте+волне+регионе / все визиты проекта+волны+региона
        """
        try:
            # Ищем колонку RS
            rs_col = None
            for col in array_df.columns:
                if any(name in str(col).lower() for name in ['эм', 'rs']):
                    rs_col = col
                    break
            
            if not rs_col:
                return {}
            
            
            # Все визиты проекта+волны+региона
            project_wave_region_mask = (
                (array_df['Код анкеты'] == project_code) &
                (array_df['Название проекта'] == wave_name) &
                (array_df['Регион'] == region)
            )
            filtered_visits = array_df[project_wave_region_mask]
            
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
    
    
    def extract_hierarchical_data(self, array_df, google_df=None):
        """
        Создаёт полную иерархию Проект→Клиент→Волна→Регион→DSM→ASM→RS
        с базовой информацией о проекте
        """
        
        try:
            # 1. Определяем колонку региона
            region_col = 'Регион short' if 'Регион short' in array_df.columns else 'Регион'
            
            # Создаём иерархию из array_df (уникальные цепочки)
            hierarchy = pd.DataFrame({
                'Проект': array_df['Код анкеты'].fillna('Не указано'),
                'Клиент': array_df['Имя клиента'].fillna('Не указано'),
                'Волна': array_df['Название проекта'].fillna('Не указано'),
                'Регион': array_df[region_col].fillna('Не указано'),
                'DSM': array_df['ЗОД'].fillna('Не указано'),
                'ASM': array_df['АСС'].fillna('Не указано'),
                'RS': array_df['ЭМ'].fillna('Не указано'),
                'ПО': array_df['ПО'].fillna('не определено'),           # ✅ Берем ПО из массива
                'Полевой': array_df['Полевой']                          # Для фильтрации
            })
            
            # 🔴 ТОЛЬКО ПОЛЕВЫЕ ПРОЕКТЫ
            hierarchy = hierarchy[hierarchy['Полевой'] == 1]
            hierarchy = hierarchy.drop('Полевой', axis=1)
            
            # Удаляем дубликаты
            hierarchy = hierarchy.drop_duplicates().reset_index(drop=True)
            
            # Даты - по умолчанию пустые
            hierarchy['Дата старта'] = pd.NaT
            hierarchy['Дата финиша'] = pd.NaT
            
            # 3. Обогащаем ТОЛЬКО датами из google_df
            if google_df is not None and not google_df.empty:
                try:
                    # Маппинги ТОЛЬКО для дат
                    start_mapping = {}
                    finish_mapping = {}
                    
                    for idx, row in google_df.iterrows():
                        code = str(row.get('Код проекта RU00.000.00.01SVZ24', '')).strip()
                        if code and code not in ['nan', '']:
                            # Даты
                            start_date = row.get('Дата старта')
                            finish_date = row.get('Дата финиша с продлением')
                            
                            if pd.notna(start_date):
                                start_mapping[code] = start_date
                            if pd.notna(finish_date):
                                finish_mapping[code] = finish_date
                    
                    # Применяем маппинги ТОЛЬКО для дат
                    hierarchy['Дата старта'] = hierarchy['Проект'].map(start_mapping)
                    hierarchy['Дата финиша'] = hierarchy['Проект'].map(finish_mapping)
                    
                except Exception as e:
                    st.warning(f"⚠️ Не удалось обогатить датами из гугл таблицы: {str(e)[:100]}")
            
            # 4. Рассчитываем длительность
            hierarchy['Длительность'] = 0
            mask_valid_dates = hierarchy['Дата старта'].notna() & hierarchy['Дата финиша'].notna()
            
            if mask_valid_dates.any():
                hierarchy.loc[mask_valid_dates, 'Длительность'] = (
                    hierarchy.loc[mask_valid_dates, 'Дата финиша'] - 
                    hierarchy.loc[mask_valid_dates, 'Дата старта']
                ).dt.days + 1
            
            # 5. Сортируем
            hierarchy = hierarchy.sort_values(['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM', 'RS'])
            hierarchy = hierarchy[hierarchy['RS'] != 'Итого']
            
            return hierarchy
            
        except KeyError as e:
            missing_col = str(e).replace("'", "")
            st.error(f"❌ В массиве отсутствует колонка: '{missing_col}'")
            st.write("📋 **Какие колонки есть в массиве:**")
            cols_list = ", ".join(array_df.columns)
            st.write(f"`{cols_list}`")
            return pd.DataFrame()
            
        except Exception as e:
            st.error(f"❌ Ошибка создания иерархии: {str(e)[:200]}")
            return pd.DataFrame()
    
    def calculate_hierarchical_plan_on_date(self, hierarchy_df, array_df, calc_params):
        """
        РАССЧИТЫВАЕТ ПЛАН ТОЛЬКО ДЛЯ УРОВНЯ RS
        Ключевое исправление: распределение ДНЕВНОГО плана по долям RS
        """
        try:
            if hierarchy_df.empty or array_df.empty:
                return pd.DataFrame()
            
            start_period = calc_params['start_date']
            end_period = calc_params['end_date']
            coefficients = calc_params['coefficients']
            
            # Планы проектов+волн+регионов
            project_wave_region_plans = array_df.groupby([
                'Код анкеты', 
                'Название проекта',
                'Регион'
            ]).size()
            
            results = []
            
            for _, row in hierarchy_df.iterrows():
                region = row['Регион']
                project_code = row['Проект']
                wave_name = row['Волна']
                
                # План проекта+волны+регион
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
                rs_weights = self._calculate_rs_weights(array_df, project_code, wave_name, region)
                rs_name = row['RS']
                
                if rs_name not in rs_weights or rs_weights[rs_name] <= 0:
                    continue
                
                rs_weight = rs_weights[rs_name]
                
                # ✅ ПРАВИЛЬНО: дневной план RS = дневной план волны × доля RS
                rs_daily_plan = daily_plan_wave * rs_weight
                
                # ✅ ПРАВИЛЬНО: план RS на дату = дневной план × дни в периоде
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
                    'ПО': row.get('ПО', 'не определено'),
                    'Уровень': 'RS',
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
            print(f"❌ Ошибка: {e}")
            import traceback
            print(traceback.format_exc())
            return pd.DataFrame()
        
    def calculate_hierarchical_fact_on_date(self, plan_df, array_df, calc_params):
        """
        ОПТИМИЗИРОВАННЫЙ: считает факт за один проход с учетом волн
        """
        try:
            if plan_df.empty or array_df.empty:
                return pd.DataFrame()
            
            result_df = plan_df.copy()
            region_col = 'Регион short' if 'Регион short' in array_df.columns else 'Регион'
            
            # Ищем колонки
            status_col = ' Статус' if ' Статус' in array_df.columns else 'Статус'
            
            rs_col = None
            for col in array_df.columns:
                if any(name in str(col).lower() for name in ['эм', 'rs']):
                    rs_col = col
                    break
            
            if not rs_col:
                result_df['Факт проекта, шт.'] = 0
                result_df['Факт на дату, шт.'] = 0
                return result_df
            
            # ФИЛЬТРЫ
            completed_mask = array_df[status_col] == 'Выполнено'
            start_date = pd.Timestamp(calc_params['start_date'])
            end_date = pd.Timestamp(calc_params['end_date'])
            period_mask = (
                (array_df['Дата визита'] >= start_date) &
                (array_df['Дата визита'] <= end_date)
            )
            
            # СЧИТАЕМ ФАКТЫ С УЧЕТОМ ВОЛН
            completed_df = array_df[completed_mask]
            rs_facts_total = completed_df.groupby([
                'Код анкеты',          # Проект
                'Название проекта',    # Волна
                region_col,            # Регион
                rs_col                 # RS
            ]).size().to_dict()
            
            completed_in_period = array_df[completed_mask & period_mask]
            rs_facts_period = completed_in_period.groupby([
                'Код анкеты',
                'Название проекта', 
                region_col,         
                rs_col
            ]).size().to_dict()
            
            # ДОБАВЛЯЕМ ФАКТЫ К RS УРОВНЮ
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
            print(f"❌ Ошибка: {e}")
            return pd.DataFrame()


    # 6. РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ ПОКАЗАТЕЛЕЙ                
    def _calculate_metrics(self, df, calc_params=None):
        """
        Упрощённый расчёт метрик (как в исходном коде)
        """
        df = df.copy()
        
        # 1. Базовые метрики
        df['%ПФ на дату'] = 0.0
        mask = df['План на дату, шт.'] > 0
        df.loc[mask, '%ПФ на дату'] = (df.loc[mask, 'Факт на дату, шт.'] / 
                                       df.loc[mask, 'План на дату, шт.'] * 100).round(1)
        
        df['△План/Факт на дату, шт.'] = (df['План на дату, шт.'] - 
                                         df['Факт на дату, шт.']).round(1)
        
        # df['%ПФ проекта'] = 0.0
        # mask_proj = df['План проекта, шт.'] > 0
        # df.loc[mask_proj, '%ПФ проекта'] = (df.loc[mask_proj, 'Факт проекта, шт.'] / 
        #                                    df.loc[mask_proj, 'План проекта, шт.'] * 100).round(1)
        
        # df['△План/Факт проекта, шт.'] = (df['План проекта, шт.'] - 
        #                                  df['Факт проекта, шт.']).round(1)
        
        # 2. Метрики по дням (только с calc_params)
        if calc_params and 'Дата старта' in df.columns:
            end_period = calc_params['end_date']
            
            df['Дней потрачено'] = 0
            df['Дней до конца проекта'] = 0
            df['Ср. план на день для 100% плана'] = 0.0
            
            for idx, row in df.iterrows():
                start_date = row.get('Дата старта')
                finish_date = row.get('Дата финиша')
                duration = row.get('Длительность', 0)
                
                if pd.notna(start_date) and pd.notna(finish_date) and duration > 0:
                    # Дней потрачено
                    days_spent = (end_period - start_date.date()).days + 1
                    df.at[idx, 'Дней потрачено'] = max(0, min(days_spent, duration))
                    
                    # Дней до конца проекта
                    days_left = (finish_date.date() - end_period).days
                    df.at[idx, 'Дней до конца проекта'] = max(0, days_left)
                    
                    # # Средний план на день
                    # plan_project = row.get('План проекта, шт.', 0)
                    # if plan_project > 0:
                    #     df.at[idx, 'Ср. план на день для 100% плана'] = round(plan_project / duration, 1)
            
            # Утилизация тайминга
            df['Утилизация тайминга, %'] = 0.0
            mask_duration = df['Длительность'] > 0
            df.loc[mask_duration, 'Утилизация тайминга, %'] = (
                df.loc[mask_duration, 'Дней потрачено'] / 
                df.loc[mask_duration, 'Длительность'] * 100
            ).round(1)
            
            # Фокус
            # df['Фокус'] = 'Нет'
            # mask_focus = (
            #     (df['%ПФ проекта'] < 80) & 
            #     (df['Утилизация тайминга, %'] > 80) & 
            #     (df['Утилизация тайминга, %'] < 100)
            # )
            # df.loc[mask_focus, 'Фокус'] = 'Да'
        
        # 3. Исполнение проекта = %ПФ на дату
        df['Исполнение Проекта,%'] = df['%ПФ на дату']
        
        return df
        

# Глобальный экземпляр
visit_calculator = VisitCalculator()


































