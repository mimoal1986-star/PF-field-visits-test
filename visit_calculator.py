# utils/visit_calculator.py
# draft 2.0 
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
                'ПО': visits_df['ПО'].fillna('не определено'),           # ✅ Берем ПО из массива
                'Полевой': visits_df['Полевой']                          # Для фильтрации
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
            cols_list = ", ".join(visits_df.columns)
            st.write(f"`{cols_list}`")
            return pd.DataFrame()
            
        except Exception as e:
            st.error(f"❌ Ошибка создания иерархии: {str(e)[:200]}")
            return pd.DataFrame()
    
    def calculate_hierarchical_plan_on_date(self, hierarchy_df, visits_df, calc_params):
        """
        РАССЧИТЫВАЕТ ПЛАН ТОЛЬКО ДЛЯ УРОВНЯ RS
        Ключевое исправление: распределение ДНЕВНОГО плана по долям RS
        """
        try:
            if hierarchy_df.empty or visits_df.empty:
                return pd.DataFrame()
            
            start_period = calc_params['start_date']
            end_period = calc_params['end_date']
            coefficients = calc_params['coefficients']
            
            # Планы проектов+волн+регионов
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
            
                # 🔴 ВСТАВЬ ДИАГНОСТИКУ СЮДА - ПОСЛЕ РАСЧЕТА days_in_period
                if _ == 0:
                    st.write(f"**7️⃣ Проверка дат и периода:**")
                    st.write(f"   Дата старта: {start_date.date()}")
                    st.write(f"   Дата финиша: {finish_date.date()}")
                    st.write(f"   Период расчета: {start_period} - {end_period}")
                    st.write(f"   Дней в периоде: {days_in_period}")
    
                if days_in_period == 0:
                    continue
                
                # ДНЕВНОЙ ПЛАН ВОЛНЫ (равномерное распределение)
                daily_plan_wave = total_plan / duration
                
                # ДОЛИ RS
                rs_weights = self._calculate_rs_weights(visits_df, project_code, wave_name, region)
                rs_name = row['RS']
                # 🔴 ДИАГНОСТИКА ДЛЯ ПЕРВОЙ СТРОКИ
                if _ == 0:  # только для первой итерации
                    st.write(f"**6️⃣ Доли RS для первого проекта:**")
                    st.write(f"   Проект: {project_code}")
                    st.write(f"   Волна: {wave_name}")
                    st.write(f"   Регион: {region}")
                    st.write(f"   RS из иерархии: '{rs_name}'")
                    st.write(f"   Найденные RS: {list(rs_weights.keys()) if rs_weights else 'НЕТ'}")
                    st.write(f"   Вес для этого RS: {rs_weights.get(rs_name, 0)}")
        
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
                    'План проекта, шт.': total_plan,
                    'План на дату, шт.': round(rs_plan_on_date, 1),
                    'Длительность': int(duration),
                    'Дата старта': start_date,
                    'Дата финиша': finish_date,
                    'Дней в периоде': days_in_period,
                    'Дневной план RS, шт.': round(rs_daily_plan, 2)
                })

                # 🔴 ДИАГНОСТИКА - СРАЗУ ПОСЛЕ APPEND
                if _ == 0:
                    st.write(f"**8️⃣ Результат для первого проекта:**")
                    st.write(f"   Добавлен в results: ✅ ДА")
                    st.write(f"   План на дату: {round(rs_plan_on_date, 1)}")
                    st.write(f"   Текущий размер results: {len(results)}")
                            
            if not results:
                return pd.DataFrame()
            return pd.DataFrame(results)
            
        except Exception as e:
            print(f"❌ Ошибка: {e}")
            import traceback
            print(traceback.format_exc())
            return pd.DataFrame()
        
    def calculate_hierarchical_fact_on_date(self, plan_df, visits_df, calc_params):
        try:
            if plan_df.empty or visits_df.empty:
                return pd.DataFrame()
            
            result_df = plan_df.copy()
            region_col = 'Регион short'
            
            # Ищем колонку статуса
            st.write("📋 **Все колонки в массиве:**")
            st.write(list(visits_df.columns))
            
            status_col = None
            for col in visits_df.columns:
                st.write(f"Проверяем: '{col}' -> очищенная: '{col.strip()}'")
                col_clean = col.strip()
                if col_clean == 'Статус':
                    status_col = col
                    break
            
            if not status_col:
                st.error(f"❌ Не найдена колонка 'Статус'")
                st.write(f"Искали: 'Статус'")
                st.write(f"Доступные колонки: {list(visits_df.columns)}")
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
                visits_df['Дата визита'] = pd.to_datetime(visits_df['Дата визита'], errors='coerce')
            
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
            
            # ПРОВЕРКА что колонка создалась
            if 'Факт на дату, шт.' not in result_df.columns:
                st.error("❌ Колонка 'Факт на дату, шт.' НЕ СОЗДАЛАСЬ!")
                result_df['Факт на дату, шт.'] = 0
    
            return result_df  # ← КЛЮЧЕВОЕ: возвращаем df с колонками!
            
        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
            return pd.DataFrame()


    # 6. РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ ПОКАЗАТЕЛЕЙ
    def _calculate_metrics(self, fact_df, calc_params=None, plan_df=None):
        df = fact_df.copy()
        
        # Берем план из plan_df
        if plan_df is not None and 'План на дату, шт.' in plan_df.columns:
            df['План на дату, шт.'] = plan_df['План на дату, шт.']
        
        # ✅ ПРОВЕРКА: есть ли колонка 'Факт на дату, шт.'
        if 'Факт на дату, шт.' not in df.columns:
            st.error("❌ В fact_df нет колонки 'Факт на дату, шт.'")
            df['Факт на дату, шт.'] = 0
        
        # 1. Базовые метрики
        df['%ПФ на дату'] = 0.0
        
        mask = df['План на дату, шт.'] > 0
        if mask.any():
            df.loc[mask, '%ПФ на дату'] = (df.loc[mask, 'Факт на дату, шт.'] / 
                                           df.loc[mask, 'План на дату, шт.'] * 100).round(1)
        
        df['△План/Факт на дату, шт.'] = (df['План на дату, шт.'] - 
                                         df['Факт на дату, шт.']).round(1)
        
        return df
    
    # def _calculate_metrics(self, fact_df, calc_params=None, plan_df=None):
    #     """Упрощённый расчёт метрик (как в исходном коде)"""
    #     df = fact_df.copy()
        
    #     # Берем план из plan_df, если он передан
    #     if plan_df is not None and 'План на дату, шт.' in plan_df.columns:
    #         df['План на дату, шт.'] = plan_df['План на дату, шт.']
        
    #     # 1. Базовые метрики
    #     df['%ПФ на дату'] = 0.0
    #     mask = df['План на дату, шт.'] > 0
    #     if mask.any():  # Добавляем проверку
    #         df.loc[mask, '%ПФ на дату'] = (df.loc[mask, 'Факт на дату, шт.'] / 
    #                                        df.loc[mask, 'План на дату, шт.'] * 100).round(1)
        
    #     df['△План/Факт на дату, шт.'] = (df['План на дату, шт.'] - 
    #                                      df['Факт на дату, шт.']).round(1)
        
    #     # 2. Метрики по дням (только с calc_params)
    #     if calc_params and 'Дата старта' in df.columns:
    #         end_period = calc_params['end_date']
            
    #         df['Дней потрачено'] = 0
    #         df['Дней до конца проекта'] = 0
    #         df['Ср. план на день для 100% плана'] = 0.0
            
    #         for idx, row in df.iterrows():
    #             start_date = row.get('Дата старта')
    #             finish_date = row.get('Дата финиша')
    #             duration = row.get('Длительность', 0)
                
    #             if pd.notna(start_date) and pd.notna(finish_date) and duration > 0:
    #                 # Дней потрачено
    #                 days_spent = (end_period - start_date.date()).days + 1
    #                 df.at[idx, 'Дней потрачено'] = max(0, min(days_spent, duration))
                    
    #                 # Дней до конца проекта
    #                 days_left = (finish_date.date() - end_period).days
    #                 df.at[idx, 'Дней до конца проекта'] = max(0, days_left)
            
    #         # Утилизация тайминга, %
    #         df['Утилизация тайминга, %'] = 0.0
    #         mask_duration = df['Длительность'] > 0
    #         if mask_duration.any():
    #             df.loc[mask_duration, 'Утилизация тайминга, %'] = (
    #                 df.loc[mask_duration, 'Дней потрачено'] / 
    #                 df.loc[mask_duration, 'Длительность'] * 100
    #             ).round(1)
        
    #     # 3. Исполнение Проекта = %ПФ на дату
    #     df['Исполнение Проекта,%'] = df['%ПФ на дату']
        
    #     return df
        

# Глобальный экземпляр
visit_calculator = VisitCalculator()


















































