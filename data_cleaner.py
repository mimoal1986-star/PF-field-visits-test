# utils/data_cleaner.py
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st
import io


class DataCleaner:

    def _find_column(self, df, possible_names):
        """Находит колонку по возможным названиям"""
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    def clean_google(self, df):
        """
        Шаги 1-7: Очистка Гугл таблицы (Проекты Сервизория)
        """
        if df is None or df.empty:
            st.warning("⚠️ Гугл таблица пустая или не загружена")
            return None
        
        df_clean = df.copy()
        original_rows = len(df_clean)
        original_cols = len(df_clean.columns)
        
        st.info(f"🧹 Начинаю очистку Гугл таблицы: {original_rows} строк × {original_cols} колонок")
        
        # === ШАГ 1: Удалить дубликаты записей ===
        st.write("**1️⃣ Удаляю дубликаты записей...**")
        
        # Ищем поля с учетом реальных названий
        code_field = self._find_column(df_clean, [
            'Код проекта RU00.000.00.01SVZ24',  # Основное название
        ])
        
        start_date_field = self._find_column(df_clean, [
            'Дата старта', # Основное название
        ])
        
        end_date_field = self._find_column(df_clean, [
            'Дата финиша с продлением',  # Основное название
        ])
        
        # Собираем найденные поля
        existing_fields = []
        field_display_names = []
        
        if code_field:
            existing_fields.append(code_field)
            field_display_names.append('Код проекта')
            
        if start_date_field:
            existing_fields.append(start_date_field)
            field_display_names.append('Дата старта')
            
        if end_date_field:
            existing_fields.append(end_date_field)
            field_display_names.append('Дата финиша')
        
        # Проверяем сколько полей нашлось
        if len(existing_fields) == 3:
            before = len(df_clean)
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first')
            after = len(df_clean)
            removed = before - after
            
            if removed > 0:
                st.success(f"   ✅ Удалено {removed} дубликатов")
                st.info(f"   По полям: {', '.join(field_display_names)}")
                st.info(f"   Фактические имена: {', '.join(existing_fields)}")
            else:
                st.info("   ℹ️ Дубликатов не найдено")
                
        elif len(existing_fields) >= 1:
            st.warning(f"   ⚠️ Найдено только {len(existing_fields)} из 3 полей: {', '.join(field_display_names)}")
            
            # Все равно пытаемся удалить дубли по найденным полям
            before = len(df_clean)
            df_clean = df_clean.drop_duplicates(subset=existing_fields, keep='first')
            after = len(df_clean)
            removed = before - after
            
            if removed > 0:
                st.success(f"   ✅ Удалено {removed} дубликатов (по найденным полям)")
            else:
                st.info("   ℹ️ Дубликатов не найдено по найденным полям")
                
        else:
            st.warning("   ⚠️ Не найдено ни одного ключевого поля для проверки дубликатов")
            st.info("   Проверьте названия колонок в файле")
            
            # Показываем какие колонки есть
            st.info("   **Найденные колонки:**")
            for i, col in enumerate(df_clean.columns[:10]):  # Первые 10 колонок
                st.info(f"   {i+1}. {col}")
        
        # === ШАГ 2: Сжать пробелы в кодах проектов ===
        st.write("**2️⃣ Чищу пробелы в кодах проектов...**")
        
        code_col = self._find_column(df_clean, ['Код проекта', 'Код', 'Project Code', 'КодПроекта'])
        
        if code_col:
            # Сохраняем оригинальные значения
            original_codes = df_clean[code_col].copy()
            
            # Приводим к строке
            df_clean[code_col] = df_clean[code_col].astype(str)
            
            # ТОЛЬКО удаляем пробелы в начале и конце (по инструкции)
            df_clean[code_col] = df_clean[code_col].str.strip()
            
            # Считаем изменения
            changed = (original_codes.fillna('') != df_clean[code_col].fillna('')).sum()
            if changed > 0:
                st.success(f"   ✅ Исправлено {changed} кодов проектов (удалены пробелы в начале/конце)")
            else:
                st.info("   ℹ️ Пробелы в кодах не найдены")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")
        
        # === ШАГ 3: Заполнить пустые коды проектов ===
        st.write("**3️⃣ Заполняю пустые коды проектов...**")
        
        if code_col:
            name_col = self._find_column(df_clean, ['Имя проекта', 'Название проекта', 'Проект', 'Project Name'])
            
            if name_col:
                # Определяем пустые коды
                empty_mask = (
                    df_clean[code_col].isna() | 
                    (df_clean[code_col].astype(str).str.strip() == '') |
                    (df_clean[code_col].astype(str).str.strip() == 'nan') |
                    (df_clean[code_col].astype(str).str.strip() == 'None')
                )
                
                empty_count = empty_mask.sum()
                
                if empty_count > 0:
                    # Базовая логика: Код проекта = Имя проекта
                    df_clean.loc[empty_mask, code_col] = df_clean.loc[empty_mask, name_col]
                    st.success(f"   ✅ Заполнено {empty_count} пустых кодов (временное решение)")
                    st.info("   ⚠️ Полная логика требует объединения с массивом")
                else:
                    st.info("   ℹ️ Пустых кодов не найдено")
            else:
                st.warning("   ⚠️ Колонка с именем проекта не найдена")
        
        # === ШАГ 4: Форматировать Пилоты/Семплы/Мультикоды ===
        st.write("**4️⃣ Форматирую Пилоты/Семплы/Мультикоды...**")
        
        # Используем code_col из шага 2 если есть
        if 'code_col' in locals() and code_col:
            target_col = code_col
        else:
            target_col = self._find_column(df_clean, [
                'Код проекта RU00.000.00.01SVZ24',
                'Код проекта',
                'Код'
            ])
        
        if target_col:
            changes_count = 0
            target_values = ['пилот', 'семпл', 'мультикод']
            
            for idx, value in df_clean[target_col].items():
                if pd.isna(value):
                    continue
                    
                str_value = str(value).strip()
                lower_value = str_value.lower()
                
                for target in target_values:
                    if target in lower_value:
                        formatted_value = str_value.capitalize() if str_value else str_value
                        
                        if formatted_value != str_value:
                            df_clean.at[idx, target_col] = formatted_value
                            changes_count += 1
                            break
            
            if changes_count > 0:
                st.success(f"   ✅ Отформатировано {changes_count} значений")
                st.info("   Пример: 'пиЛот' → 'Пилот', 'СЕМПЛ' → 'Семпл'")
            else:
                st.info("   ℹ️ Значения уже отформатированы")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")
        
        # === ШАГ 5: Заполнить пустые даты ===
        st.write("**5️⃣ Заполняю пустые даты...**")
        
        date_patterns = ['дата', 'date', 'срок', 'time', 'начал', 'старт', 'финиш', 'конец', 'заверш']
        date_cols = []
        
        for col in df_clean.columns:
            col_lower = str(col).lower()
            if any(pattern in col_lower for pattern in date_patterns):
                date_cols.append(col)
        
        if date_cols:
            date_fixes = 0
            
            for col in date_cols:
                try:
                    df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    empty_dates = df_clean[col].isna().sum()
                    
                    if empty_dates > 0:
                        col_lower = str(col).lower()
                        is_start_date = any(word in col_lower for word in ['старт', 'начал', 'start'])
                        is_end_date = any(word in col_lower for word in ['финиш', 'конец', 'end', 'заверш'])
                        current_date = pd.Timestamp.now()
                        
                        for idx in df_clean[df_clean[col].isna()].index:
                            if is_start_date:
                                df_clean.at[idx, col] = current_date.replace(day=1)
                            elif is_end_date:
                                next_month = current_date.replace(day=28) + timedelta(days=4)
                                df_clean.at[idx, col] = next_month - timedelta(days=next_month.day)
                            else:
                                df_clean.at[idx, col] = current_date
                        
                        date_fixes += empty_dates
                        st.info(f"   Заполнено {empty_dates} пустых дат в '{col}'")
                except Exception as e:
                    st.warning(f"   Ошибка в колонке '{col}': {str(e)[:100]}")
            
            if date_fixes > 0:
                st.success(f"   ✅ Заполнено {date_fixes} пустых дат")
            else:
                st.info("   ℹ️ Пустых дат не найдено")
        else:
            st.warning("   ⚠️ Колонки с датами не найдены")
        
        # === ШАГ 6: Исправить даты по бизнес-правилам ===
        st.write("**6️⃣ Применяю бизнес-правила для дат...**")
        
        date_rules_applied = 0
        today = pd.Timestamp.now()
        first_day_current_month = today.replace(day=1, hour=0, minute=0, second=0)
        next_month = today.replace(day=28) + timedelta(days=4)
        last_day_current_month = next_month - timedelta(days=next_month.day)
        
        for col in date_cols:
            if col not in df_clean.columns:
                continue
                
            col_lower = str(col).lower()
            
            if any(word in col_lower for word in ['старт', 'начал', 'start']):
                try:
                    if df_clean[col].dtype != 'datetime64[ns]':
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    mask = df_clean[col] < first_day_current_month
                    
                    if mask.any():
                        df_clean.loc[mask, col] = first_day_current_month
                        date_rules_applied += mask.sum()
                        st.info(f"   Исправлено {mask.sum()} дат старта")
                except Exception as e:
                    st.warning(f"   Не удалось обработать даты старта в '{col}': {str(e)[:100]}")
            
            elif any(word in col_lower for word in ['финиш', 'конец', 'end']):
                try:
                    if df_clean[col].dtype != 'datetime64[ns]':
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    
                    mask = df_clean[col] > last_day_current_month
                    
                    if mask.any():
                        df_clean.loc[mask, col] = last_day_current_month
                        date_rules_applied += mask.sum()
                        st.info(f"   Исправлено {mask.sum()} дат финиша")
                except Exception as e:
                    st.warning(f"   Не удалось обработать даты финиша в '{col}': {str(e)[:100]}")
        
        if date_rules_applied > 0:
            st.success(f"   ✅ Применено {date_rules_applied} бизнес-правил для дат")
        else:
            st.info("   ℹ️ Бизнес-правила для дат не потребовались")
        
        # === ШАГ 7: Добавить признак 'Полевой' ===
        st.write("**7️⃣ Добавляю признак 'Полевой'...**")
        
        if 'Полевой' not in df_clean.columns:
            df_clean['Полевой'] = 1
            st.success("   ✅ Добавлен признак 'Полевой' = 1 для всех записей")
        else:
            empty_field = df_clean['Полевой'].isna().sum()
            if empty_field > 0:
                df_clean['Полевой'] = df_clean['Полевой'].fillna(1)
                st.success(f"   ✅ Заполнено {empty_field} пустых значений")
            else:
                st.info("   ℹ️ Признак 'Полевой' уже заполнен")
        
        # === ИТОГИ ОЧИСТКИ ===
        st.markdown("---")
        final_rows = len(df_clean)
        final_cols = len(df_clean.columns)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Строк до очистки", original_rows, 
                     delta=f"{final_rows - original_rows}")
        
        with col2:
            st.metric("Строк после", final_rows)
        
        with col3:
            removed_pct = ((original_rows - final_rows) / original_rows * 100) if original_rows > 0 else 0
            st.metric("Удалено", f"{removed_pct:.1f}%")
        
        st.success(f"✅ Гугл таблица успешно очищена!")
        
        return df_clean
    
    def export_to_excel(self, original_df, cleaned_df, filename="очищенные_данные"):
        """
        Создает Excel файл с тремя вкладками для сравнения
        """
        try:
            if original_df is None or cleaned_df is None:
                return None
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Вкладка 1: Оригинальные данные
                original_df.to_excel(writer, sheet_name='ОРИГИНАЛ', index=False)
                
                # Вкладка 2: Очищенные данные
                cleaned_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ', index=False)
                
                # Вкладка 3: Сравнение изменений
                comparison_data = []
                
                comparison_data.append({
                    'Параметр': 'Количество строк',
                    'Оригинал': len(original_df),
                    'Очищено': len(cleaned_df),
                    'Изменение': len(cleaned_df) - len(original_df)
                })
                
                comparison_data.append({
                    'Параметр': 'Количество колонок',
                    'Оригинал': len(original_df.columns),
                    'Очищено': len(cleaned_df.columns),
                    'Изменение': len(cleaned_df.columns) - len(original_df.columns)
                })
                
                added_cols = set(cleaned_df.columns) - set(original_df.columns)
                removed_cols = set(original_df.columns) - set(cleaned_df.columns)
                
                if added_cols:
                    comparison_data.append({
                        'Параметр': 'Добавленные колонки',
                        'Оригинал': '-',
                        'Очищено': ', '.join(added_cols),
                        'Изменение': f'+{len(added_cols)}'
                    })
                
                if removed_cols:
                    comparison_data.append({
                        'Параметр': 'Удаленные колонки',
                        'Оригинал': ', '.join(removed_cols),
                        'Очищено': '-',
                        'Изменение': f'-{len(removed_cols)}'
                    })
                
                comparison_df = pd.DataFrame(comparison_data)
                comparison_df.to_excel(writer, sheet_name='СРАВНЕНИЕ', index=False)
            
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"Ошибка при создании Excel: {e}")
            return None


# Глобальный экземпляр
data_cleaner = DataCleaner()
