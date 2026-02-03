# utils/visit_calculator.py
# draft 2.0 
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

class VisitCalculator:
    
    def extract_hierarchical_data(self, array_df, google_df=None):
        """
        Создаёт полную иерархию Проект→Клиент→Волна→Регион→DSM→ASM→RS
        с базовой информацией о проекте
        """
        try:
            # 1. Создаём иерархию из array_df (уникальные цепочки)
            hierarchy = pd.DataFrame({
                'Проект': array_df['Код анкеты'].fillna('Не указано'),
                'Клиент': array_df['Имя клиента'].fillna('Не указано'),
                'Волна': array_df['Название проекта'].fillna('Не указано'),
                'Регион': array_df['Регион'].fillna('Не указано'),
                'DSM': array_df['ЗОД'].fillna('Не указано'),
                'ASM': array_df['АСС'].fillna('Не указано'),
                'RS': array_df['ЭМ рег'].fillna('Не указано')
            })
            
            # Удаляем полные дубликаты
            hierarchy = hierarchy.drop_duplicates().reset_index(drop=True)
            
            # 2. Добавляем базовую информацию
            # ПО - по умолчанию
            hierarchy['ПО'] = 'не определено'
            
            # Даты - по умолчанию пустые
            hierarchy['Дата старта'] = pd.NaT
            hierarchy['Дата финиша'] = pd.NaT
            
            # 3. Обогащаем данными из google_df если есть
            if google_df is not None and not google_df.empty:
                try:
                    # Создаём маппинги
                    portal_mapping = {}
                    start_mapping = {}
                    finish_mapping = {}
                    
                    # Проходим по гугл таблице
                    for idx, row in google_df.iterrows():
                        code = str(row.get('Код проекта RU00.000.00.01SVZ24', '')).strip()
                        if code and code not in ['nan', '']:
                            # ПО
                            portal = str(row.get('Портал на котором идет проект (для работы полевой команды)', '')).strip()
                            if portal:
                                portal_mapping[code] = portal
                            
                            # Даты
                            start_date = row.get('Дата старта')
                            finish_date = row.get('Дата финиша с продлением')
                            
                            if pd.notna(start_date):
                                start_mapping[code] = start_date
                            if pd.notna(finish_date):
                                finish_mapping[code] = finish_date
                    
                    # Применяем маппинги
                    hierarchy['ПО'] = hierarchy['Проект'].map(portal_mapping).fillna('не определено')
                    hierarchy['Дата старта'] = hierarchy['Проект'].map(start_mapping)
                    hierarchy['Дата финиша'] = hierarchy['Проект'].map(finish_mapping)
                    
                except Exception as e:
                    st.warning(f"⚠️ Не удалось обогатить данными из гугл таблицы: {str(e)[:100]}")
            
            # 4. Рассчитываем длительность (в днях)
            hierarchy['Длительность'] = 0
            mask_valid_dates = hierarchy['Дата старта'].notna() & hierarchy['Дата финиша'].notna()
            
            if mask_valid_dates.any():
                hierarchy.loc[mask_valid_dates, 'Длительность'] = (
                    hierarchy.loc[mask_valid_dates, 'Дата финиша'] - 
                    hierarchy.loc[mask_valid_dates, 'Дата старта']
                ).dt.days + 1
            
            # 5. Сортируем для удобства
            hierarchy = hierarchy.sort_values(['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM', 'RS'])
            
            return hierarchy
            
        except KeyError as e:
            # Если нет какой-то колонки в array_df
            missing_col = str(e).replace("'", "")
            st.error(f"❌ В массиве отсутствует колонка: '{missing_col}'")
            # Возвращаем пустой DataFrame
            return pd.DataFrame()
            
        except Exception as e:
            st.error(f"❌ Ошибка создания иерархии: {str(e)[:200]}")
            return pd.DataFrame()

def calculate_hierarchical_plan_on_date(self, hierarchy_df, array_df, calc_params):
    """
    Рассчитывает план на дату для всей иерархии
    """
    try:
        if hierarchy_df.empty or array_df.empty:
            return pd.DataFrame()
        
        # Параметры
        start_period = calc_params['start_date']
        end_period = calc_params['end_date']
        coefficients = calc_params['coefficients']
        total_coeff = sum(coefficients)
        norm_coeff = [c/total_coeff for c in coefficients]
        
        # Планы проектов
        project_plans = array_df.groupby('Код анкеты').size()
        
        results = []
        
        # Для каждой уникальной RS
        for _, row in hierarchy_df.iterrows():
            project_code = row['Проект']
            
            # План проекта
            if project_code in project_plans.index:
                total_plan = project_plans[project_code]
            else:
                continue
            
            # Даты
            start_date = row['Дата старта']
            finish_date = row['Дата финиша']
            duration = row['Длительность']
            
            if pd.isna(start_date) or pd.isna(finish_date) or duration <= 0:
                continue
            
            # Распределение по этапам
            stage_days = [duration // 4] * 3
            stage_days.append(duration - sum(stage_days))
            
            stage_plans = [total_plan * coeff for coeff in norm_coeff[:3]]
            stage_plans.append(total_plan - sum(stage_plans))
            
            # План на дату
            plan_on_date = 0.0
            current_date = start_date
            
            for i in range(4):
                if stage_plans[i] > 0 and stage_days[i] > 0:
                    daily_plan = stage_plans[i] / stage_days[i]
                    
                    for day in range(stage_days[i]):
                        check_date = current_date + timedelta(days=day)
                        if start_period <= check_date.date() <= end_period:
                            plan_on_date += daily_plan
                
                current_date += timedelta(days=stage_days[i])
            
            # Запись
            results.append({
                'Проект': row['Проект'],
                'Клиент': row['Клиент'],
                'Волна': row['Волна'],
                'Регион': row['Регион'],
                'DSM': row['DSM'],
                'ASM': row['ASM'],
                'RS': row['RS'],
                'ПО': row.get('ПО', 'не определено'),
                'Уровень': 'RS',
                'План проекта, шт.': float(total_plan),
                'План на дату, шт.': round(plan_on_date, 1),
                'Длительность': int(duration),
                'Дата старта': start_date,
                'Дата финиша': finish_date
            })
        
        if not results:
            return pd.DataFrame()
        
        # Создаём DataFrame
        plan_df = pd.DataFrame(results)
        
        # Автоагрегация вверх
        levels = [
            ('ASM', ['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM']),
            ('DSM', ['Проект', 'Клиент', 'Волна', 'Регион', 'DSM']),
            ('Регион', ['Проект', 'Клиент', 'Волна', 'Регион']),
            ('Волна', ['Проект', 'Клиент', 'Волна']),
            ('Клиент', ['Проект', 'Клиент']),
            ('Проект', ['Проект'])
        ]
        
        all_results = plan_df.to_dict('records')
        
        for level_name, group_cols in levels:
            # Группировка
            grouped = plan_df.groupby(group_cols, as_index=False).agg({
                'План проекта, шт.': 'sum',
                'План на дату, шт.': 'sum',
                'Длительность': 'first',
                'Дата старта': 'first',
                'Дата финиша': 'first'
            })
            
            # Округление
            grouped['План на дату, шт.'] = grouped['План на дату, шт.'].round(1)
            grouped['План проекта, шт.'] = grouped['План проекта, шт.'].round(1)
            
            # Заполняем остальные колонки
            for col in ['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM', 'RS', 'ПО']:
                if col not in group_cols:
                    if col == 'ПО':
                        # Самое частое ПО в группе
                        po_mode = plan_df[plan_df['ПО'] != 'не определено']['ПО'].mode()
                        grouped['ПО'] = po_mode.iloc[0] if not po_mode.empty else 'не определено'
                    else:
                        grouped[col] = 'Итого'
            
            grouped['Уровень'] = level_name
            all_results.extend(grouped.to_dict('records'))
        
        # Финальный DataFrame
        final_df = pd.DataFrame(all_results)
        
        # Проверка
        if not final_df.empty:
            rs_sum = final_df[final_df['Уровень'] == 'RS']['План на дату, шт.'].sum()
            project_sum = final_df[final_df['Уровень'] == 'Проект']['План на дату, шт.'].sum()
            
            if abs(rs_sum - project_sum) > 0.01:
                st.warning(f"⚠️ Расхождение: RS={rs_sum:.1f}, Проекты={project_sum:.1f}")
        
        return final_df
        
    except Exception as e:
        st.error(f"❌ Ошибка в calculate_hierarchical_plan_on_date: {str(e)[:200]}")
        return pd.DataFrame()
    
    def calculate_fact_on_date_full(self, base_data, array_df, cxway_df, calc_params):
        """Рассчитывает 'Факт на дату, шт.' и 'Факт проекта' (Массив + CXWAY)."""
        
        result = base_data.copy()
        result['Факт проекта, шт.'] = 0
        result['Факт на дату, шт.'] = 0
        
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
            if project_po in ['Чеккер', 'CXWAY', 'не определено']:
                project_visits_array = array_df[
                    (array_df['Код анкеты'] == project_code) &
                    (array_df['Название проекта'] == project_name) &
                    (array_df[' Статус'] == 'Выполнено')
                ]
                fact_total = len(project_visits_array)
    
            
            if fact_total == 0:
                continue
                
            result.at[idx, 'Факт проекта, шт.'] = fact_total
            
            # 2. ФАКТ НА ДАТУ: визиты от 1-го числа месяца до вчера
            if project_po in ['Чеккер', 'CXWAY', 'не определено'] and project_visits_array is not None:
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



















