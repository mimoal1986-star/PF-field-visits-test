# utils/dataviz.py
# draft 3.1 
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
        
        # Переименовываем колонки для отображения
        rename_cols = {
            'ЗОД': 'DSM',
            'АСС': 'ASM',
            'ЭМ': 'RS'
        }
        data = data.rename(columns=rename_cols)
        
        # 🔍 ФИЛЬТРЫ (каскадные)
        with st.expander("🔍 Фильтры", expanded=True):
            # Определяем, какую колонку региона использовать
            region_col = 'Регион'
            if 'Регион short' in data.columns and 'Регион' not in data.columns:
                region_col = 'Регион short'
            
            # Получаем уникальные значения для фильтров
            all_dsm = data['DSM'].dropna().unique() if 'DSM' in data.columns else []
            all_asm = data['ASM'].dropna().unique() if 'ASM' in data.columns else []
            all_regions = data[region_col].dropna().unique() if region_col in data.columns else []
            all_clients = data['Клиент'].dropna().unique() if 'Клиент' in data.columns else []
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                selected_dsm = st.multiselect('DSM', all_dsm, key='filter_dsm')
            with col2:
                asm_options = all_asm
                if selected_dsm and 'DSM' in data.columns and 'ASM' in data.columns:
                    asm_options = data[data['DSM'].isin(selected_dsm)]['ASM'].dropna().unique()
                selected_asm = st.multiselect('ASM', asm_options, key='filter_asm')
            with col3:
                region_options = all_regions
                filtered_for_region = data.copy()
                if selected_dsm and 'DSM' in filtered_for_region.columns:
                    filtered_for_region = filtered_for_region[filtered_for_region['DSM'].isin(selected_dsm)]
                if selected_asm and 'ASM' in filtered_for_region.columns:
                    filtered_for_region = filtered_for_region[filtered_for_region['ASM'].isin(selected_asm)]
                if region_col in filtered_for_region.columns:
                    region_options = filtered_for_region[region_col].dropna().unique()
                selected_region = st.multiselect('Регион', region_options, key='filter_region')
            with col4:
                client_options = all_clients
                filtered_for_client = data.copy()
                if selected_dsm and 'DSM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['DSM'].isin(selected_dsm)]
                if selected_asm and 'ASM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['ASM'].isin(selected_asm)]
                if selected_region and region_col in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client[region_col].isin(selected_region)]
                if 'Клиент' in filtered_for_client.columns:
                    client_options = filtered_for_client['Клиент'].dropna().unique()
                selected_client = st.multiselect('Клиент', client_options, key='filter_client')
        
        # Применяем фильтры к данным
        filtered_data = data.copy()
        
        if selected_dsm and 'DSM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['DSM'].isin(selected_dsm)]
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
            show_regions = st.checkbox("Регионы", key='show_regions')
        with col2:
            show_dsm = st.checkbox("DSM", key='show_dsm')
        with col3:
            show_asm = st.checkbox("ASM", key='show_asm')
        with col4:
            show_rs = st.checkbox("RS", key='show_rs')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Проект', 'Клиент', 'ПО']
        
        if show_regions and region_col in filtered_data.columns:
            group_cols.append(region_col)
        if show_dsm and 'DSM' in filtered_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        
        # Агрегируем данные с учетом развертки
        if len(group_cols) > 3:
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
            
            existing_agg = {k: v for k, v in agg_columns.items() if k in filtered_data.columns}
            detailed_data = filtered_data.groupby(group_cols).agg(existing_agg).reset_index()
            
            # Пересчитываем метрики для детальных данных
            detailed_data['План/Факт на дату,%'] = 0.0
            mask_plan = detailed_data['План на дату, шт.'] > 0
            if mask_plan.any():
                detailed_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                    detailed_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                    detailed_data.loc[mask_plan, 'План на дату, шт.'] * 100
                ).round(1)
            
            detailed_data['План/Факт проекта,%'] = 0.0
            mask_project_plan = detailed_data['План проекта, шт.'] > 0
            if mask_project_plan.any():
                detailed_data.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                    detailed_data.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                    detailed_data.loc[mask_project_plan, 'План проекта, шт.'] * 100
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
            
            detailed_data['Исполнение проекта,%'] = detailed_data['План/Факт проекта,%']
            
            if 'plan_calc_params' in st.session_state:
                days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                                st.session_state['plan_calc_params']['start_date']).days + 1
            else:
                days_in_period = 12
                
            detailed_data['Прогноз на месяц, шт.'] = (
                detailed_data['Факт на дату, шт.'] / days_in_period * 28
            ).round(1)
            
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
            project_data = self.create_project_summary(filtered_data)
            # Добавляем План/Факт проекта если его нет
            if 'План/Факт проекта,%' not in project_data.columns and 'План проекта, шт.' in project_data.columns:
                project_data['План/Факт проекта,%'] = 0.0
                mask = project_data['План проекта, шт.'] > 0
                if mask.any():
                    project_data.loc[mask, 'План/Факт проекта,%'] = (
                        project_data.loc[mask, 'Факт проекта, шт.'] / 
                        project_data.loc[mask, 'План проекта, шт.'] * 100
                    ).round(1)
        
        st.caption(f"📌 Отображается записей: {len(project_data)}")
        
        if project_data.empty:
            st.warning("⚠️ Нет данных после фильтрации")
            return
        
        # KPI - 6 метрик в два ряда по 3
        st.markdown("### 📊 Ключевые показатели")
        
        # Первый ряд: План проекта, Факт проекта, План/Факт проекта
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_project_total = project_data['План проекта, шт.'].sum() if 'План проекта, шт.' in project_data.columns else 0
            st.metric("📊 План проекта", f"{plan_project_total:,.0f} шт")
        
        with col2:
            fact_project_total = project_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in project_data.columns else 0
            st.metric("✅ Факт проекта", f"{fact_project_total:,.0f} шт")
        
        with col3:
            pf_project_percent = (fact_project_total / plan_project_total * 100) if plan_project_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_project_percent:.1f}%")
        
        # Второй ряд: План на дату, Факт на дату, План/Факт на дату
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = project_data['План на дату, шт.'].sum() if 'План на дату, шт.' in project_data.columns else 0
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        
        with col5:
            fact_date_total = project_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in project_data.columns else 0
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
        
        # Колонки для отображения
        display_columns = [
            'Проект',
            'Клиент',
            'ПО',
            'Длительность',
            'План проекта, шт.',
            'Факт проекта, шт.',
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Прогноз на месяц, шт.',
            'Фокус',
            'Дней до конца проекта',
            'Утилизация тайминга, %',
            'Ср. план на день для 100% плана'
        ]
        
        # Добавляем колонки развертки, если они есть
        if show_regions and region_col in project_data.columns:
            display_columns.insert(3, region_col)
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
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        
        if 'План/Факт проекта,%' in df_display.columns:
            df_display['План/Факт проекта,%'] = df_display['План/Факт проекта,%'].map(lambda x: f"{x:.1f}%")
        
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        if 'Утилизация тайминга, %' in df_display.columns:
            df_display['Утилизация тайминга, %'] = df_display['Утилизация тайминга, %'].map(lambda x: f"{x:.1f}%")
        
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







