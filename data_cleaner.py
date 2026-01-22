# utils/data_cleaner.py
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st
import io


class DataCleaner:
    """
    Очистка данных по инструкции для ИУ Аудиты
    """
    
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
            
            # НЕ меняем внутренние пробелы!
            # df_clean[code_col] = df_clean[code_col].str.replace(r'\s+', ' ', regex=True)  # УБРАТЬ!
            
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
        
        # === ШАГ 4: Проверить капитализацию ===
        st.write("**4️⃣ Проверяю капитализацию категориальных полей...**")
        
        categorical_fields = ['Пилот', 'Семпл', 'Тип проекта', 'Статус', 'Тип', 'Статус проекта']
        existing_cat_fields = [col for col in categorical_fields if col in df_clean.columns]
        
        if existing_cat_fields:
            changes_count = 0
            
            for col in existing_cat_fields:
                original_values = df_clean[col].copy()
                
                # Приводим к строке
                df_clean[col] = df_clean[col].astype(str)
                
                # ИСПРАВЛЕНИЕ: Всегда приводим к формату "Пилот"
                mask = df_clean[col].str.strip() != ''
                df_clean.loc[mask, col] = df_clean.loc[mask, col].apply(
                    lambda x: x.strip().capitalize() if pd.notna(x) else x
                )
                
                # Считаем изменения
                changed = (original_values.fillna('') != df_clean[col].fillna('')).sum()
                changes_count += changed
            
            if changes_count > 0:
                st.success(f"   ✅ Исправлено {changes_count} значений (приведено к формату 'Пилот')")
            else:
                st.info("   ℹ️ Значения уже в правильном формате")
        else:
            st.info("   ℹ️ Категориальные поля не найдены")
        
        # === ШАГ 5: Заполнить пустые даты ===
        st.write("**5️⃣ Заполняю пустые даты...**")
        
        date_cols = [col for col in df_clean.columns 
                    if any(keyword in col.lower() for keyword in ['дата', 'date', 'срок', 'time'])]
        
        if date_cols:
            date_fixes = 0
            
            for col in date_cols:
                # Конвертируем в datetime
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                
                # Считаем пустые даты
                empty_dates = df_clean[col].isna().sum()
                
                if empty_dates > 0:
                    is_start_date = any(word in col.lower() for word in ['старт', 'начал', 'start', 'начала'])
                    is_end_date = any(word in col.lower() for word in ['финиш', 'конец', 'end', 'заверш'])
                    
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
            
            if date_fixes > 0:
                st.success(f"   ✅ Заполнено {date_fixes} пустых дат")
            else:
                st.info("   ℹ️ Пустых дат не найдено")
        else:
            st.warning("   ⚠️ Колонки с датами не найдены")
        
        # === ШАГ 6: Исправить даты по бизнес-правилам ===
        st.write("**6️⃣ Применяю бизнес-правила для дат...**")
        
        date_rules_applied = self._apply_date_business_rules(df_clean)
        
        if date_rules_applied > 0:
            st.success(f"   ✅ Применено {date_rules_applied} бизнес-правил для дат")
        else:
            st.info("   ℹ️ Бизнес-правила для дат не потребовались")
        
        # === ШАГ 7: Добавить признак Полевой/Неполевой ===
        st.write("**7️⃣ Добавляю признак 'Полевой'...**")
        
        df_clean = self._add_field_type_flag(df_clean)
        
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
    
    def _find_column(self, df, possible_names):
        """
        Находит колонку по возможным названиям
        """
        # Сначала проверяем точное совпадение
        for name in possible_names:
            if name in df.columns:
                return name
        
        # Если нет точного, проверяем частичное совпадение
        for name in possible_names:
            for col in df.columns:
                # Приводим к строке и сравниваем без учета регистра и пробелов
                col_str = str(col).strip().lower().replace(' ', '')
                name_str = str(name).strip().lower().replace(' ', '')
                
                if col_str == name_str:
                    return col
                
                # Проверяем частичное вхождение
                if name_str in col_str or col_str in name_str:
                    return col
        
        return None
    
    def _apply_date_business_rules(self, df):
        """Шаг 6: Бизнес-правила для дат"""
        rules_applied = 0
        
        current_date = pd.Timestamp.now()
        current_month = current_date.month
        current_year = current_date.year
        
        # Правило 1: Дата начала = прошлый месяц → 1 число текущего
        start_cols = [col for col in df.columns if any(word in col.lower() for word in ['начал', 'старт'])]
        
        for col in start_cols:
            if col in df.columns and pd.api.types.is_datetime64_any_dtype(df[col]):
                mask = (
                    (df[col].dt.month == current_month - 1) & 
                    (df[col].dt.year == current_year)
                )
                
                if current_month == 1:
                    mask = mask | (
                        (df[col].dt.month == 12) & 
                        (df[col].dt.year == current_year - 1)
                    )
                
                if mask.any():
                    df.loc[mask, col] = df.loc[mask, col].apply(
                        lambda x: x.replace(month=current_month, year=current_year, day=1) 
                        if pd.notna(x) else x
                    )
                    rules_applied += mask.sum()
        
        # Правило 2: Дата конца < 5 числа → 5 число месяца
        end_cols = [col for col in df.columns if any(word in col.lower() for word in ['конец', 'финиш'])]
        
        for col in end_cols:
            if col in df.columns and pd.api.types.is_datetime64_any_dtype(df[col]):
                mask = (df[col].dt.day < 5) & df[col].notna()
                
                if mask.any():
                    df.loc[mask, col] = df.loc[mask, col].apply(
                        lambda x: x.replace(day=5) if pd.notna(x) else x
                    )
                    rules_applied += mask.sum()
        
        return rules_applied
    
    def _add_field_type_flag(self, df):
        """Шаг 7: Добавляет признак Полевой/Неполевой"""
        if 'Полевой' not in df.columns:
            df['Полевой'] = 1
        return df
    
    def export_to_excel(self, original_df, cleaned_df, filename="очищенные_данны"):
        """
        Создает Excel файл с двумя вкладками: оригинал и очищенный
        для сверки изменений
        """
        if original_df is None or cleaned_df is None:
            return None
        
        # Создаем буфер для Excel
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Вкладка 1: Оригинальные данные
            original_df.to_excel(writer, sheet_name='ОРИГИНАЛ', index=False)
            
            # Вкладка 2: Очищенные данные
            cleaned_df.to_excel(writer, sheet_name='ОЧИЩЕННЫЙ', index=False)
            
            # Вкладка 3: Сравнение изменений
            comparison = self._create_comparison_sheet(original_df, cleaned_df)
            comparison.to_excel(writer, sheet_name='СРАВНЕНИЕ', index=False)
        
        output.seek(0)
        
        return output
    
    def _create_comparison_sheet(self, original_df, cleaned_df):
        """Создает лист сравнения изменений"""
        comparison_data = []
        
        # Сравниваем размеры
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
        
        # Ищем добавленные/удаленные колонки
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
        
        return pd.DataFrame(comparison_data)


# Глобальный экземпляр
data_cleaner = DataCleaner()



