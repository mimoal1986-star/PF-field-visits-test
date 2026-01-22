# utils/data_cleaner.py
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import streamlit as st


class DataCleaner:
    """
    Очистка данных по инструкции для ИУ Аудиты
    """
    
    def clean_google(self, df):
        """
        Шаги 1-7: Очистка Гугл таблицы (Проекты Сервизория)
        
        Шаг 1: Удалить дубликаты записей
        Шаг 2: Сжать пробелы по полю Код проекта
        Шаг 3: Заполнить Код проекта, если Пусто
        Шаг 4: Проверить Пилоты, Семплы и т.п. - с заглавной буквы
        Шаг 5: Заполнить пустые даты
        Шаг 6: Исправить дату начала, дату конца
        Шаг 7: Добавить признак Полевой/Неполевой
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
        
        # Находим колонки для проверки дубликатов
        possible_key_cols = ['Код проекта', 'Дата старта', 'Дата финиша', 
                           'Дата начала', 'Дата конца', 'Имя проекта', 'Проект']
        
        existing_key_cols = [col for col in possible_key_cols if col in df_clean.columns]
        
        if existing_key_cols:
            before = len(df_clean)
            df_clean = df_clean.drop_duplicates(subset=existing_key_cols, keep='first')
            after = len(df_clean)
            removed = before - after
            
            if removed > 0:
                st.success(f"   ✅ Удалено {removed} дубликатов")
            else:
                st.info("   ℹ️ Дубликатов не найдено")
        else:
            st.warning("   ⚠️ Не найдены ключевые колонки для проверки дубликатов")
        
        # === ШАГ 2: Сжать пробелы в кодах проектов ===
        st.write("**2️⃣ Чищу пробелы в кодах проектов...**")
        
        # Находим колонку с кодом проекта
        code_cols = ['Код проекта', 'Код', 'Project Code', 'КодПроекта']
        code_col = None
        
        for col in code_cols:
            if col in df_clean.columns:
                code_col = col
                break
        
        if code_col:
            # Сохраняем оригинальные значения для отчета
            original_codes = df_clean[code_col].copy()
            
            # Приводим к строке
            df_clean[code_col] = df_clean[code_col].astype(str)
            
            # Удаляем пробелы в начале и конце
            df_clean[code_col] = df_clean[code_col].str.strip()
            
            # Заменяем множественные пробелы на один
            df_clean[code_col] = df_clean[code_col].str.replace(r'\s+', ' ', regex=True)
            
            # Заменяем нестандартные пробелы
            df_clean[code_col] = df_clean[code_col].str.replace(r'[　⠀  ]', ' ', regex=True)
            
            # Считаем изменения
            changed = (original_codes.fillna('') != df_clean[code_col].fillna('')).sum()
            if changed > 0:
                st.success(f"   ✅ Исправлено {changed} кодов проектов")
            else:
                st.info("   ℹ️ Пробелы в кодах не найдены")
        else:
            st.warning("   ⚠️ Колонка с кодом проекта не найдена")
        
        # === ШАГ 3: Заполнить пустые коды проектов ===
        st.write("**3️⃣ Заполняю пустые коды проектов...**")
        
        if code_col:
            # Находим колонку с именем проекта
            name_cols = ['Имя проекта', 'Название проекта', 'Проект', 'Project Name']
            name_col = None
            
            for col in name_cols:
                if col in df_clean.columns:
                    name_col = col
                    break
            
            if name_col:
                # Определяем пустые коды
                empty_mask = (
                    df_clean[code_col].isna() | 
                    (df_clean[code_col].astype(str).str.strip() == '') |
                    (df_clean[code_col].astype(str).str.strip() == 'nan')
                )
                
                empty_count = empty_mask.sum()
                
                if empty_count > 0:
                    # Заполняем пустые коды именами проектов
                    df_clean.loc[empty_mask, code_col] = df_clean.loc[empty_mask, name_col]
                    st.success(f"   ✅ Заполнено {empty_count} пустых кодов")
                else:
                    st.info("   ℹ️ Пустых кодов не найдено")
            else:
                st.warning("   ⚠️ Колонка с именем проекта не найдена")
        
        # === ШАГ 4: Проверить капитализацию ===
        st.write("**4️⃣ Проверяю капитализацию категориальных полей...**")
        
        # Поля которые должны быть с заглавной буквы
        categorical_fields = [
            'Пилот', 'Семпл', 'Тип проекта', 'Статус', 'Тип', 'Статус проекта',
            'Вид', 'Категория', 'Тип визита'
        ]
        
        existing_cat_fields = [col for col in categorical_fields if col in df_clean.columns]
        
        if existing_cat_fields:
            changes_count = 0
            
            for col in existing_cat_fields:
                original_values = df_clean[col].copy()
                
                # Приводим к строке
                df_clean[col] = df_clean[col].astype(str)
                
                # Капитализируем (только если вся строка в нижнем регистре)
                mask = df_clean[col].str.islower() & (df_clean[col].str.strip() != '')
                df_clean.loc[mask, col] = df_clean.loc[mask, col].str.capitalize()
                
                # Считаем изменения
                changed = (original_values.fillna('') != df_clean[col].fillna('')).sum()
                changes_count += changed
            
            if changes_count > 0:
                st.success(f"   ✅ Исправлено {changes_count} значений")
            else:
                st.info("   ℹ️ Значения уже в правильном регистре")
        else:
            st.info("   ℹ️ Категориальные поля не найдены")
        
        # === ШАГ 5: Заполнить пустые даты ===
        st.write("**5️⃣ Заполняю пустые даты...**")
        
        # Находим все колонки с датами
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
                    # Определяем тип даты: начало или конец
                    is_start_date = any(word in col.lower() for word in ['старт', 'начал', 'start', 'начала'])
                    is_end_date = any(word in col.lower() for word in ['финиш', 'конец', 'end', 'заверш'])
                    
                    current_date = pd.Timestamp.now()
                    
                    for idx in df_clean[df_clean[col].isna()].index:
                        if is_start_date:
                            # Первое число текущего месяца
                            df_clean.at[idx, col] = current_date.replace(day=1)
                        elif is_end_date:
                            # Последнее число текущего месяца
                            next_month = current_date.replace(day=28) + timedelta(days=4)
                            df_clean.at[idx, col] = next_month - timedelta(days=next_month.day)
                        else:
                            # По умолчанию - текущая дата
                            df_clean.at[idx, col] = current_date
                    
                    date_fixes += empty_dates
                    st.info(f"   Заполнено {empty_dates} дат в '{col}'")
            
            if date_fixes > 0:
                st.success(f"   ✅ Всего заполнено {date_fixes} пустых дат")
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
        
        # Статистика очистки
        final_rows = len(df_clean)
        final_cols = len(df_clean.columns)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Строк до очистки", original_rows, 
                     delta=f"{final_rows - original_rows}")
        
        with col2:
            st.metric("Колонок", final_cols)
        
        with col3:
            removed_pct = ((original_rows - final_rows) / original_rows * 100) if original_rows > 0 else 0
            st.metric("Удалено строк", f"{removed_pct:.1f}%")
        
        st.success(f"✅ Гугл таблица успешно очищена!")
        
        # Показываем предпросмотр очищенных данных
        with st.expander("👀 Просмотр очищенных данных (первые 10 строк)"):
            st.dataframe(df_clean.head(10), use_container_width=True)
        
        return df_clean
    
    def _apply_date_business_rules(self, df):
        """
        Применяет бизнес-правила для дат:
        1. Если дата начала = прошлый месяц → 1 число текущего месяца
        2. Если дата конца < 5 числа → 5 число месяца
        """
        rules_applied = 0
        
        current_date = pd.Timestamp.now()
        current_month = current_date.month
        current_year = current_date.year
        
        # Правило 1: Дата начала = прошлый месяц → 1 число текущего
        start_cols = [col for col in df.columns if any(word in col.lower() for word in ['начал', 'старт'])]
        
        for col in start_cols:
            if col in df.columns and pd.api.types.is_datetime64_any_dtype(df[col]):
                # Находим даты из прошлого месяца
                mask = (
                    (df[col].dt.month == current_month - 1) & 
                    (df[col].dt.year == current_year)
                )
                
                # Если январь, проверяем декабрь прошлого года
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
        """
        Добавляет признак Полевой/Неполевой
        Временно: все проекты считаем полевыми
        TODO: Интеграция со справочником "Разметка проекта"
        """
        if 'Полевой' not in df.columns:
            df['Полевой'] = 1  # 1 = полевой, 0 = неполевой
            st.info("   ✅ Добавлена колонка 'Полевой' (все = 1)")
        
        # TODO: Реализовать VLOOKUP с Разметкой проекта
        # TODO: Удалить записи с Полевой = 0
        
        return df
    
    def get_cleaning_report(self, original_df, cleaned_df):
        """
        Генерирует отчет об очистке
        """
        if original_df is None or cleaned_df is None:
            return None
        
        report = {
            'original_rows': len(original_df),
            'cleaned_rows': len(cleaned_df),
            'rows_removed': len(original_df) - len(cleaned_df),
            'original_cols': len(original_df.columns),
            'cleaned_cols': len(cleaned_df.columns),
            'columns_added': set(cleaned_df.columns) - set(original_df.columns),
            'columns_removed': set(original_df.columns) - set(cleaned_df.columns),
        }
        
        return report


# Глобальный экземпляр
data_cleaner = DataCleaner()