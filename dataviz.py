# utils/dataviz.py
# draft 3.0 
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

class DataVisualizer:

    def create_project_summary(self, df):
        """
        Агрегация данных по проектам
        Одна строка = один проект
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        # Проверяем наличие колонки 'Проект'
        project_col = 'Проект'
        if project_col not in df.columns:
            st.error(f"❌ В данных нет колонки '{project_col}'")
            return pd.DataFrame()
        
        # Список колонок для агрегации
        agg_columns = {
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Длительность': 'first',
            'Дата старта': 'first',
            'Дата финиша': 'first',
            'Клиент': 'first',
            'ПО': 'first',
            'Дней до конца проекта': 'first',
            'Утилизация тайминга, %': 'first',
            'Ср. план на день для 100% плана': 'sum'
        }
        
        # Только существующие колонки
        existing_agg = {k: v for k, v in agg_columns.items() if k in df.columns}
        
        # Группируем по проекту
        project_agg = df.groupby(project_col).agg(existing_agg).reset_index()
        
        # 1. План/Факт на дату,% (было Исполнение проекта,%)
        project_agg['План/Факт на дату,%'] = 0.0
        mask_plan = project_agg['План на дату, шт.'] > 0
        if mask_plan.any():
            project_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                project_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                project_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        # 2. △План/Факт на дату, шт (исправленная формула: Факт - План)
        project_agg['△План/Факт на дату, шт'] = (
            project_agg['Факт на дату, шт.'] - project_agg['План на дату, шт.']
        ).round(1)
        
        # 3. △План/Факт на дату, % (исправленная формула)
        project_agg['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            project_agg.loc[mask_plan, '△План/Факт на дату, %'] = (
                (project_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                 project_agg.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # 4. Исполнение проекта,% (было %ПФ проекта)
        project_agg['Исполнение проекта,%'] = 0.0
        mask_project_plan = project_agg['План проекта, шт.'] > 0
        if mask_project_plan.any():
            project_agg.loc[mask_project_plan, 'Исполнение проекта,%'] = (
                project_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                project_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        # 5. Прогноз на месяц, шт.
        if 'plan_calc_params' in st.session_state:
            days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                            st.session_state['plan_calc_params']['start_date']).days + 1
        else:
            days_in_period = 12
            
        project_agg['Прогноз на месяц, шт.'] = (
            project_agg['Факт на дату, шт.'] / days_in_period * 28
        ).round(1)
        
        # 6. Фокус
        project_agg['Фокус'] = 'Нет'
        if all(col in project_agg.columns for col in ['Исполнение проекта,%', 'Утилизация тайминга, %']):
            mask_focus = (
                (project_agg['Исполнение проекта,%'] < 80) & 
                (project_agg['Утилизация тайминга, %'] > 80) & 
                (project_agg['Утилизация тайминга, %'] < 100)
            )
            project_agg.loc[mask_focus, 'Фокус'] = 'Да'
        
        # Сортируем по План/Факт на дату,%
        project_agg = project_agg.sort_values('План/Факт на дату,%', ascending=True)
        
        return project_agg
    
    def create_planfact_tab(self, data, hierarchy_df=None):
        """Создает вкладку ПланФакт на дату с фильтрами и разверткой"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по проектам")
        
        # 🔍 ФИЛЬТРЫ (каскадные)
        with st.expander("🔍 Фильтры", expanded=True):
            # Определяем, какую колонку региона использовать
            region_col = 'Регион'  # по умолчанию короткий
            if 'Регион' in data.columns and 'Регион short' in data.columns:
                # Если есть обе, используем длинный для фильтрации
                region_col = 'Регион'
            elif 'Регион' in data.columns:
                region_col = 'Регион'
            elif 'Регион short' in data.columns:
                region_col = 'Регион short'
            
            # Получаем уникальные значения для фильтров
            all_zod = data['DSM'].dropna().unique() if 'DSM' in data.columns else []
            all_asm = data['ASM'].dropna().unique() if 'ASM' in data.columns else []
            all_regions = data[region_col].dropna().unique() if region_col in data.columns else []
            all_clients = data['Клиент'].dropna().unique() if 'Клиент' in data.columns else []
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                selected_zod = st.multiselect('ЗОД (DSM)', all_zod, key='filter_zod')
            with col2:
                # Если выбран ЗОД, фильтруем АСС по нему
                asm_options = all_asm
                if selected_zod and 'DSM' in data.columns and 'ASM' in data.columns:
                    asm_options = data[data['DSM'].isin(selected_zod)]['ASM'].dropna().unique()
                selected_asm = st.multiselect('АСС (ASM)', asm_options, key='filter_asm')
            with col3:
                # Фильтруем регионы по выбранным ЗОД и АСС
                region_options = all_regions
                filtered_for_region = data.copy()
                if selected_zod and 'DSM' in filtered_for_region.columns:
                    filtered_for_region = filtered_for_region[filtered_for_region['DSM'].isin(selected_zod)]
                if selected_asm and 'ASM' in filtered_for_region.columns:
                    filtered_for_region = filtered_for_region[filtered_for_region['ASM'].isin(selected_asm)]
                if region_col in filtered_for_region.columns:
                    region_options = filtered_for_region[region_col].dropna().unique()
                selected_region = st.multiselect('Регион', region_options, key='filter_region')
            with col4:
                # Фильтруем клиентов по всем выбранным фильтрам
                client_options = all_clients
                filtered_for_client = data.copy()
                if selected_zod and 'DSM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['DSM'].isin(selected_zod)]
                if selected_asm and 'ASM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['ASM'].isin(selected_asm)]
                if selected_region and region_col in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client[region_col].isin(selected_region)]
                if 'Клиент' in filtered_for_client.columns:
                    client_options = filtered_for_client['Клиент'].dropna().unique()
                selected_client = st.multiselect('Клиент', client_options, key='filter_client')
        
        # Применяем фильтры к данным
        filtered_data = data.copy()  # ← ВОТ ЗДЕСЬ!
        
        if selected_zod and 'DSM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['DSM'].isin(selected_zod)]
        if selected_asm and 'ASM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if selected_region and region_col in filtered_data.columns:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        if selected_client and 'Клиент' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        
        
        # 📊 РАЗВЕРТКА (ЧЕК-БОКСЫ)
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            show_regions = st.checkbox("Показать регионы", key='show_regions')
        with col2:
            show_dsm = st.checkbox("Показать ЗОД", key='show_dsm')
        with col3:
            show_asm = st.checkbox("Показать АСС", key='show_asm')
        with col4:
            show_rs = st.checkbox("Показать RS", key='show_rs')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Проект', 'Клиент', 'ПО']
        
        if show_regions and 'Регион' in filtered_data.columns:
            group_cols.append('Регион')
        if show_dsm and 'DSM' in filtered_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        
        # Агрегируем данные с учетом развертки
        if len(group_cols) > 3:  # если есть дополнительные уровни
            agg_columns = {
                'План проекта, шт.': 'sum',
                'План на дату, шт.': 'sum',
                'Факт проекта, шт.': 'sum',
                'Факт на дату, шт.': 'sum',
                'Длительность': 'first',
                'Дата старта': 'first',
                'Дата финиша': 'first',
                'Дней до конца проекта': 'first',
                'Утилизация тайминга, %': 'first',
                'Ср. план на день для 100% плана': 'sum'
            }
            
            # Только существующие колонки
            existing_agg = {k: v for k, v in agg_columns.items() if k in filtered_data.columns}
            
            # Группируем с учетом развертки
            detailed_data = filtered_data.groupby(group_cols).agg(existing_agg).reset_index()
            
            # Пересчитываем метрики для детальных данных
            detailed_data['План/Факт на дату,%'] = 0.0
            mask_plan = detailed_data['План на дату, шт.'] > 0
            if mask_plan.any():
                detailed_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                    detailed_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                    detailed_data.loc[mask_plan, 'План на дату, шт.'] * 100
                ).round(1)
            
            detailed_data['△План/Факт на дату, шт'] = (
                detailed_data['Факт на дату, шт.'] - detailed_data['План на дату, шт.']
            ).round(1)
            
            detailed_data['△План/Факт на дату, %'] = 0.0
            if mask_plan.any():
                detailed_data.loc[mask_plan, '△План/Факт на дату, %'] = (
                    (detailed_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                     detailed_data.loc[mask_plan, 'План на дату, шт.']) - 1
                ).round(3) * 100
            
            detailed_data['Исполнение проекта,%'] = 0.0
            mask_project_plan = detailed_data['План проекта, шт.'] > 0
            if mask_project_plan.any():
                detailed_data.loc[mask_project_plan, 'Исполнение проекта,%'] = (
                    detailed_data.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                    detailed_data.loc[mask_project_plan, 'План проекта, шт.'] * 100
                ).round(1)
            
            # Прогноз на месяц
            if 'plan_calc_params' in st.session_state:
                days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                                st.session_state['plan_calc_params']['start_date']).days + 1
            else:
                days_in_period = 12
                
            detailed_data['Прогноз на месяц, шт.'] = (
                detailed_data['Факт на дату, шт.'] / days_in_period * 28
            ).round(1)
            
            # Фокус
            detailed_data['Фокус'] = 'Нет'
            if all(col in detailed_data.columns for col in ['Исполнение проекта,%', 'Утилизация тайминга, %']):
                mask_focus = (
                    (detailed_data['Исполнение проекта,%'] < 80) & 
                    (detailed_data['Утилизация тайминга, %'] > 80) & 
                    (detailed_data['Утилизация тайминга, %'] < 100)
                )
                detailed_data.loc[mask_focus, 'Фокус'] = 'Да'
            
            project_data = detailed_data
        else:
            # Если развертка не выбрана - используем стандартную агрегацию по проектам
            project_data = self.create_project_summary(filtered_data)
        
        # Показываем количество отфильтрованных записей
        st.caption(f"📌 Отображается записей: {len(project_data)}")
        
        if project_data.empty:
            st.warning("⚠️ Нет данных после фильтрации")
            return
        
        # Колонки для отображения
        display_columns = [
            'Проект',
            'Клиент',
            'ПО',
            'Длительность',
            'План проекта, шт.',
            'Факт проекта, шт.',
            'Исполнение проекта,%',
            'Фокус',
            'План на дату, шт.',
            'Факт на дату, шт.',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'План/Факт на дату,%',
            'Прогноз на месяц, шт.',
            'Дней до конца проекта',
            'Утилизация тайминга, %',
            'Ср. план на день для 100% плана, шт.'
        ]
        
        # Добавляем колонки развертки, если они есть
        if show_regions and 'Регион' in project_data.columns:
            display_columns.insert(3, 'Регион')
        if show_dsm and 'DSM' in project_data.columns:
            display_columns.insert(4, 'DSM')
        if show_asm and 'ASM' in project_data.columns:
            display_columns.insert(5, 'ASM')
        if show_rs and 'RS' in project_data.columns:
            display_columns.insert(6, 'RS')
        
        # Только существующие колонки
        existing_cols = [col for col in display_columns if col in project_data.columns]
        df_display = project_data[existing_cols].copy()
        
        # Форматирование
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        
        if 'Исполнение проекта,%' in df_display.columns:
            df_display['Исполнение проекта,%'] = df_display['Исполнение проекта,%'].map(lambda x: f"{x:.1f}%")
        
        if 'Утилизация тайминга, %' in df_display.columns:
            df_display['Утилизация тайминга, %'] = df_display['Утилизация тайминга, %'].map(lambda x: f"{x:.1f}%")
        
        # KPI
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            plan_total = df_display['План на дату, шт.'].sum() if 'План на дату, шт.' in df_display.columns else 0
            st.metric("📊 План на дату", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = df_display['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in df_display.columns else 0
            st.metric("✅ Факт на дату", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_percent:.1f}%")
        
        with col4:
            forecast_total = df_display['Прогноз на месяц, шт.'].sum() if 'Прогноз на месяц, шт.' in df_display.columns else 0
            st.metric("📈 Прогноз на месяц", f"{forecast_total:,.0f} шт")
        
        # Таблица
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, sheet_name='План_факт_проекты', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"план_факт_проекты_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# Глобальный экземпляр
dataviz = DataVisualizer()





