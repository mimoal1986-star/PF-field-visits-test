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
        
    def calculate_hierarchical_fact_on_date(self, plan_df, array_df, calc_params):
        """
        Рассчитывает факт на дату для всей иерархии
        Возвращает DataFrame с колонками: 'Факт на дату, шт.', 'Факт проекта, шт.'
        """
        try:
            if plan_df.empty or array_df.empty:
                return pd.DataFrame()
            
            result_df = plan_df.copy()
            
            # 1. Определяем колонку статуса
            status_col = 'Статус' if 'Статус' in array_df.columns else ' Статус'
            
            # 2. Фильтр: выполненные визиты в периоде
            completed_mask = (
                (array_df[status_col] == 'Выполнено') &
                (array_df['Дата визита'] >= pd.Timestamp(calc_params['start_date'])) &
                (array_df['Дата визита'] <= pd.Timestamp(calc_params['end_date']))
            )
            
            completed_df = array_df[completed_mask]
            
            # 3. Если нет выполненных - все факты = 0
            if completed_df.empty:
                result_df['Факт на дату, шт.'] = 0
                result_df['Факт проекта, шт.'] = 0
                return result_df
            
            # 4. Факт проекта (ВСЕ выполненные визиты проекта)
            project_facts = {}
            for project in result_df['Проект'].unique():
                if project != 'Итого':
                    project_mask = (
                        (array_df['Код анкеты'] == project) &
                        (array_df[status_col] == 'Выполнено')
                    )
                    project_facts[project] = array_df[project_mask].shape[0]
            
            # 5. Группируем выполненные визиты для RS
            fact_counts = completed_df.groupby([
                'Код анкеты', 'Имя клиента', 'Название проекта',
                'Регион', 'ЗОД', 'АСС', 'ЭМ рег'
            ]).size().reset_index(name='Факт_RS')
            
            # 6. Сопоставляем факты с RS строками
            rs_mask = result_df['Уровень'] == 'RS'
            
            for idx in result_df[rs_mask].index:
                row = result_df.loc[idx]
                
                # Ищем совпадение
                match = fact_counts[
                    (fact_counts['Код анкеты'] == row['Проект']) &
                    (fact_counts['Имя клиента'] == row['Клиент']) &
                    (fact_counts['Название проекта'] == row['Волна']) &
                    (fact_counts['Регион'] == row['Регион']) &
                    (fact_counts['ЗОД'] == row['DSM']) &
                    (fact_counts['АСС'] == row['ASM']) &
                    (fact_counts['ЭМ рег'] == row['RS'])
                ]
                
                # Факт на дату для RS
                result_df.at[idx, 'Факт на дату, шт.'] = match['Факт_RS'].iloc[0] if not match.empty else 0
                
                # Факт проекта для RS
                project_code = row['Проект']
                if project_code != 'Итого':
                    result_df.at[idx, 'Факт проекта, шт.'] = project_facts.get(project_code, 0)
            
            # 7. Автоагрегация фактов вверх по уровням
            levels = [
                ('ASM', ['Проект', 'Клиент', 'Волна', 'Регион', 'DSM', 'ASM']),
                ('DSM', ['Проект', 'Клиент', 'Волна', 'Регион', 'DSM']),
                ('Регион', ['Проект', 'Клиент', 'Волна', 'Регион']),
                ('Волна', ['Проект', 'Клиент', 'Волна']),
                ('Клиент', ['Проект', 'Клиент']),
                ('Проект', ['Проект'])
            ]
            
            for level_name, group_cols in levels:
                level_mask = result_df['Уровень'] == level_name
                
                for idx in result_df[level_mask].index:
                    # Находим дочерние RS строки
                    child_mask = (result_df['Уровень'] == 'RS')
                    row_values = result_df.loc[idx]
                    
                    for col in group_cols:
                        child_mask = child_mask & (result_df[col] == row_values[col])
                    
                    # Суммируем факты дочерних
                    if child_mask.any():
                        fact_sum = result_df.loc[child_mask, 'Факт на дату, шт.'].sum()
                        result_df.at[idx, 'Факт на дату, шт.'] = fact_sum
                        
                        # Факт проекта для агрегированных уровней
                        if 'Проект' in group_cols:
                            project_code = row_values['Проект']
                            if project_code != 'Итого':
                                result_df.at[idx, 'Факт проекта, шт.'] = project_facts.get(project_code, 0)
            
            # 8. Заполняем пропуски нулями
            result_df['Факт на дату, шт.'] = result_df['Факт на дату, шт.'].fillna(0)
            result_df['Факт проекта, шт.'] = result_df['Факт проекта, шт.'].fillna(0)
            
            return result_df
            
        except Exception as e:
            st.error(f"❌ Ошибка в calculate_hierarchical_fact_on_date: {str(e)[:200]}")
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
        
        df['%ПФ проекта'] = 0.0
        mask_proj = df['План проекта, шт.'] > 0
        df.loc[mask_proj, '%ПФ проекта'] = (df.loc[mask_proj, 'Факт проекта, шт.'] / 
                                           df.loc[mask_proj, 'План проекта, шт.'] * 100).round(1)
        
        df['△План/Факт проекта, шт.'] = (df['План проекта, шт.'] - 
                                         df['Факт проекта, шт.']).round(1)
        
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
                    
                    # Средний план на день
                    plan_project = row.get('План проекта, шт.', 0)
                    if plan_project > 0:
                        df.at[idx, 'Ср. план на день для 100% плана'] = round(plan_project / duration, 1)
            
            # Утилизация тайминга
            df['Утилизация тайминга, %'] = 0.0
            mask_duration = df['Длительность'] > 0
            df.loc[mask_duration, 'Утилизация тайминга, %'] = (
                df.loc[mask_duration, 'Дней потрачено'] / 
                df.loc[mask_duration, 'Длительность'] * 100
            ).round(1)
            
            # Фокус
            df['Фокус'] = 'Нет'
            mask_focus = (
                (df['%ПФ проекта'] < 80) & 
                (df['Утилизация тайминга, %'] > 80) & 
                (df['Утилизация тайминга, %'] < 100)
            )
            df.loc[mask_focus, 'Фокус'] = 'Да'
        
        # 3. Исполнение проекта = %ПФ на дату
        df['Исполнение Проекта,%'] = df['%ПФ на дату']
        
        return df
        

# Глобальный экземпляр
visit_calculator = VisitCalculator()




















