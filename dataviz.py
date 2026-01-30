import pandas as pd
import streamlit as st

class DataVisualizer:
    
    def create_planfact_tab(self, data, hierarchy_df=None):
        """Создает вкладку ПланФакт на дату"""
        if data is None or data.empty:
            st.warning("Нет данных для отчета")
            return
        
        # Копия для модификаций
        df = data.copy()
        
        # 1. Добавляем недостающие колонки
        df['△План/Факт,%'] = df.apply(
            lambda x: 0 if x['План на дату, шт.'] == 0 
            else (x['△План/Факт на дату, шт.'] / x['План на дату, шт.']) * 100, 
            axis=1
        )
        
        df['Прогноз месяц (свод),%'] = df.apply(
            lambda x: 0 if x['План проекта, шт.'] == 0 
            else (x['Прогноз на месяц, шт.'] / x['План проекта, шт.']) * 100,
            axis=1
        )
        
        # Переименовываем существующие колонки
        df = df.rename(columns={
            'Утилизация тайминга, %': 'Потребление времени, %',
            'Фокус': 'Важно/Срочно',
            'Кол-во визитов до 100% плана, шт.': 'Кол-во визитов до 100% плана, шт. ',
            'Дней до конца проекта': 'Дней до конца проекта ',
            'Ср. план на день для 100% плана': 'Ср. план на день для 100% плана, шт. '
        })
        
        # 2. Выбираем нужные колонки в правильном порядке
        display_columns = [
            'Код проекта', 'Имя клиента', 'Название проекта', 'ПО',
            'ЗОД', 'АСС', 'Регион', 'Регион short',
            'Исполнение Проекта,%', 'Потребление времени, %', 'Важно/Срочно',
            'План на дату, шт. ', 'Факт на дату, шт. ', '△План/Факт на дату, шт.', '△План/Факт,%',
            'План проекта, шт. ', 'Прогноз на месяц, шт. ', 'Прогноз месяц (свод),%',
            'Кол-во визитов до 100% плана, шт. ', 'Поручено', 'Доля Поручено, %',
            'Дней до конца проекта ', 'Ср. план на день для 100% плана, шт. '
        ]
        
        # Оставляем только существующие
        existing_cols = [col for col in display_columns if col in df.columns]
        df_display = df[existing_cols]
        
        # 3. Фильтры
        st.sidebar.header("Фильтры")
        
        # Связанные фильтры ЗОД → АСС
        if hierarchy_df is not None and 'ЗОД' in hierarchy_df.columns and 'АСС' in hierarchy_df.columns:
            all_zod = df_display['ЗОД'].dropna().unique()
            selected_zod = st.sidebar.multiselect("ЗОД", all_zod)
            
            if selected_zod:
                # Фильтруем АСС по выбранным ЗОД
                filtered_ass = hierarchy_df[hierarchy_df['ЗОД'].isin(selected_zod)]['АСС'].unique()
                df_display = df_display[df_display['ЗОД'].isin(selected_zod)]
            else:
                filtered_ass = df_display['АСС'].dropna().unique()
        else:
            selected_zod = []
            filtered_ass = df_display['АСС'].dropna().unique()
        
        selected_ass = st.sidebar.multiselect("АСС", filtered_ass)
        if selected_ass:
            df_display = df_display[df_display['АСС'].isin(selected_ass)]
        
        # Остальные фильтры
        all_clients = df_display['Имя клиента'].dropna().unique()
        selected_clients = st.sidebar.multiselect("Имя клиента", all_clients)
        if selected_clients:
            df_display = df_display[df_display['Имя клиента'].isin(selected_clients)]
        
        all_regions = df_display['Регион'].dropna().unique()
        selected_regions = st.sidebar.multiselect("Регион", all_regions)
        if selected_regions:
            df_display = df_display[df_display['Регион'].isin(selected_regions)]
        
        # 4. KPI сверху
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = df_display['План на дату, шт. '].sum() if 'План на дату, шт. ' in df_display.columns else 0
            st.metric("📊 План на дату", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = df_display['Факт на дату, шт. '].sum() if 'Факт на дату, шт. ' in df_display.columns else 0
            st.metric("✅ Факт на дату", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт %", f"{pf_percent:.1f}%")
        
        # 5. Добавляем строку Итого
        if not df_display.empty:
            total_row = self._calculate_totals(df_display)
            df_with_totals = pd.concat([df_display, total_row], ignore_index=True)
            
            # 6. Отображаем таблицу
            st.dataframe(df_with_totals, use_container_width=True, height=400)
    
    def _calculate_totals(self, df):
        """Создает строку Итого"""
        total_row = {}
        
        # Колонки для суммирования
        sum_columns = [
            'План на дату, шт. ', 'Факт на дату, шт. ', '△План/Факт на дату, шт.',
            'План проекта, шт. ', 'Прогноз на месяц, шт. ',
            'Кол-во визитов до 100% плана, шт. ', 'Поручено'
        ]
        
        for col in df.columns:
            if col in sum_columns:
                total_row[col] = df[col].sum()
            elif col == '△План/Факт,%':
                plan = df['План на дату, шт. '].sum()
                delta = df['△План/Факт на дату, шт.'].sum()
                total_row[col] = (delta / plan * 100) if plan != 0 else 0
            elif col == 'Прогноз месяц (свод),%':
                plan_total = df['План проекта, шт. '].sum()
                forecast = df['Прогноз на месяц, шт. '].sum()
                total_row[col] = (forecast / plan_total * 100) if plan_total != 0 else 0
            elif col == 'Доля Поручено, %':
                need = df['Кол-во визитов до 100% плана, шт. '].sum()
                assigned = df['Поручено'].sum()
                total_row[col] = (assigned / need * 100) if need != 0 else 0
            else:
                total_row[col] = ''
        
        # Заполняем заголовки
        total_row['Код проекта'] = 'Итого'
        total_row['Имя клиента'] = ''
        total_row['Название проекта'] = ''
        
        return pd.DataFrame([total_row])

# Глобальный экземпляр
dataviz = DataVisualizer()