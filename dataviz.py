# utils/dataviz.py
# draft 4.1 
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

class DataVisualizer:

    def __init__(self):
        self.region_mapping = {
            'AD': 'Республика Адыгея',
            'AL': 'Алтайский край',
            'AM': 'Амурская область',
            'AR': 'Архангельская область',
            'AS': 'Астраханская область',
            'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область',
            'BR': 'Брянская область',
            'BU': 'Республика Бурятия',
            'CL': 'Челябинская область',
            'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика',
            'DA': 'Республика Дагестан',
            'IN': 'Республика Ингушетия',
            'IR': 'Иркутская область',
            'IV': 'Ивановская область',
            'KA': 'Камчатский край',
            'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика',
            'KD': 'Краснодарский край',
            'KE': 'Кемеровская область',
            'KG': 'Калужская область',
            'KH': 'Хабаровский край',
            'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия',
            'KL': 'Республика Калмыкия',
            'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область',
            'KO': 'Республика Коми',
            'KS': 'Курская область',
            'KT': 'Костромская область',
            'KU': 'Курганская область',
            'KV': 'Кировская область',
            'KY': 'Красноярский край',
            'LN': 'Ленинградская область',
            'LP': 'Липецкая область',
            'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область',
            'MM': 'Мурманская область',
            'MR': 'Республика Мордовия',
            'MS': 'Московская область',
            'NG': 'Новгородская область',
            'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия',
            'NS': 'Новосибирская область',
            'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область',
            'OL': 'Орловская область',
            'OM': 'Омская область',
            'PE': 'Пермский край',
            'PR': 'Приморский край',
            'PS': 'Псковская область',
            'PZ': 'Пензенская область',
            'RK': 'Республика Крым',
            'RO': 'Ростовская область',
            'RZ': 'Рязанская область',
            'SA': 'Самарская область',
            'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область',
            'SM': 'Смоленская область',
            'SR': 'Саратовская область',
            'ST': 'Ставропольский край',
            'SV': 'Свердловская область',
            'TB': 'Тамбовская область',
            'TL': 'Тульская область',
            'TO': 'Томская область',
            'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва',
            'TV': 'Тверская область',
            'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика',
            'UL': 'Ульяновская область',
            'VG': 'Волгоградская область',
            'VL': 'Владимирская область',
            'VO': 'Вологодская область',
            'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ',
            'YS': 'Ярославская область',
            'YV': 'Еврейская автономная область',
            'ZK': 'Забайкальский край'
        }
    
    def _get_long_region(self, short_code):
        """Преобразует короткий код региона в длинное название"""
        if pd.isna(short_code) or short_code == '':
            return short_code
        return self.region_mapping.get(str(short_code).strip().upper(), short_code)

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
            'ЭМ': 'RS',
            'Имя клиента': 'Клиент'
        }
        data = data.rename(columns=rename_cols)
        
        # Определяем колонку региона
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # 🔍 ФИЛЬТРЫ (с возможностью исключения)
        with st.expander("🔍 Фильтры", expanded=True):
            
            # Получаем уникальные значения
            all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
            all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
            
            # Регионы
            if region_col in data.columns:
                unique_codes = data[region_col].dropna().unique()
                self.region_display_map = {}
                all_regions_display = []
                for code in unique_codes:
                    long_name = self._get_long_region(code)
                    self.region_display_map[long_name] = code
                    all_regions_display.append(long_name)
                all_regions_display.sort()
            else:
                all_regions_display = []
                self.region_display_map = {}
            
            all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
            
            # Переменные для хранения выбранных значений
            selected_dsm, excluded_dsm = [], []
            selected_asm, excluded_asm = [], []
            selected_region, excluded_region = [], []
            selected_client, excluded_client = [], []
            
            col1, col2, col3, col4 = st.columns(4)
            
            # === DSM ===
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_mode", horizontal=True)
                if dsm_mode == "Включить":
                    selected_dsm = st.multiselect("Выбрать", all_dsm, key="dsm_include")
                else:
                    excluded_dsm = st.multiselect("Исключить", all_dsm, key="dsm_exclude")
            
            # === ASM (с учетом выбранных DSM) ===
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio("Режим", ["Включить", "Исключить"], key="asm_mode", horizontal=True)
                
                # Ограничиваем ASM выбранными DSM (для обоих режимов)
                if selected_dsm and 'DSM' in data.columns:
                    asm_options = sorted(data[data['DSM'].isin(selected_dsm)]['ASM'].dropna().unique())
                else:
                    asm_options = all_asm
                
                if asm_mode == "Включить":
                    selected_asm = st.multiselect("Выбрать", asm_options, key="asm_include")
                else:
                    excluded_asm = st.multiselect("Исключить", asm_options, key="asm_exclude")
            
            # === Регион ===
            with col3:
                st.markdown("**Регион**")
                region_mode = st.radio("Режим", ["Включить", "Исключить"], key="region_mode", horizontal=True)
                
                if region_mode == "Включить":
                    selected_region_display = st.multiselect("Выбрать", all_regions_display, key="region_include")
                    selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
                else:
                    excluded_region_display = st.multiselect("Исключить", all_regions_display, key="region_exclude")
                    excluded_region = [self.region_display_map.get(name, name) for name in excluded_region_display]
            
            # === Клиент (с учетом выбранных DSM, ASM, региона) ===
            with col4:
                st.markdown("**Клиент**")
                client_mode = st.radio("Режим", ["Включить", "Исключить"], key="client_mode", horizontal=True)
                
                # Ограничиваем клиентов выбранными фильтрами
                client_filtered = data.copy()
                if selected_dsm and 'DSM' in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered['DSM'].isin(selected_dsm)]
                if selected_asm and 'ASM' in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered['ASM'].isin(selected_asm)]
                if selected_region and region_col in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered[region_col].isin(selected_region)]
                client_options = sorted(client_filtered['Клиент'].dropna().unique()) if 'Клиент' in client_filtered.columns else all_clients
                
                if client_mode == "Включить":
                    selected_client = st.multiselect("Выбрать", client_options, key="client_include")
                else:
                    excluded_client = st.multiselect("Исключить", client_options, key="client_exclude")
        
        # === ПРИМЕНЯЕМ ФИЛЬТРЫ ===
        filtered_data = data.copy()
        
        # DSM
        if selected_dsm:
            filtered_data = filtered_data[filtered_data['DSM'].isin(selected_dsm)]
        if excluded_dsm:
            filtered_data = filtered_data[~filtered_data['DSM'].isin(excluded_dsm)]
        
        # ASM
        if selected_asm:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if excluded_asm:
            filtered_data = filtered_data[~filtered_data['ASM'].isin(excluded_asm)]
        
        # Регион
        if selected_region:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        if excluded_region:
            filtered_data = filtered_data[~filtered_data[region_col].isin(excluded_region)]
        
        # Клиент
        if selected_client:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        if excluded_client:
            filtered_data = filtered_data[~filtered_data['Клиент'].isin(excluded_client)]
        
        # 📊 РАЗВЕРТКА (ЧЕК-БОКСЫ)
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            show_project = st.checkbox("Код проекта", key='show_project_code')
        with col2:
            show_regions = st.checkbox("Регионы", key='show_regions')
        with col3:
            show_dsm = st.checkbox("DSM", key='show_dsm')
        with col4:
            show_asm = st.checkbox("ASM", key='show_asm')
        with col5:
            show_rs = st.checkbox("RS", key='show_rs')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Клиент', 'ПО']
        
        if show_project and 'Проект' in filtered_data.columns:
            group_cols.append('Проект')
        if show_regions and region_col in filtered_data.columns:
            group_cols.append(region_col)
        if show_dsm and 'DSM' in filtered_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        
        # Агрегируем данные с учетом развертки
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
        project_data = filtered_data.groupby(group_cols).agg(existing_agg).reset_index()
        
        # Пересчитываем метрики
        project_data['План/Факт на дату,%'] = 0.0
        mask_plan = project_data['План на дату, шт.'] > 0
        if mask_plan.any():
            project_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                project_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                project_data.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        project_data['План/Факт проекта,%'] = 0.0
        mask_project_plan = project_data['План проекта, шт.'] > 0
        if mask_project_plan.any():
            project_data.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                project_data.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                project_data.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        project_data['△План/Факт на дату, шт'] = (
            project_data['Факт на дату, шт.'] - project_data['План на дату, шт.']
        ).round(1)
        
        project_data['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            project_data.loc[mask_plan, '△План/Факт на дату, %'] = (
                (project_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                 project_data.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        project_data['Исполнение проекта,%'] = project_data['План/Факт проекта,%']
        
        if 'plan_calc_params' in st.session_state:
            days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                            st.session_state['plan_calc_params']['start_date']).days + 1
        else:
            days_in_period = 12
            
        project_data['Прогноз на месяц, шт.'] = (
            project_data['Факт на дату, шт.'] / days_in_period * 28
        ).round(1)
        
        project_data['Фокус'] = 'Нет'
        if all(col in project_data.columns for col in ['Исполнение проекта,%', 'Утилизация тайминга, %']):
            mask_focus = (
                (project_data['Исполнение проекта,%'] < 80) & 
                (project_data['Утилизация тайминга, %'] > 80) & 
                (project_data['Утилизация тайминга, %'] < 100)
            )
            project_data.loc[mask_focus, 'Фокус'] = 'Да'
        
        st.caption(f"📌 Отображается записей: {len(project_data)}")
        
        if project_data.empty:
            st.warning("⚠️ Нет данных после фильтрации")
            return
        
        # KPI - 6 метрик в два ряда по 3
        st.markdown("### 📊 Ключевые показатели")
        
        # Чек-бокс Продата (справа от KPI)
        col_kpi1, col_kpi2, col_kpi3, col_checkbox = st.columns([1, 1, 1, 0.5])
        with col_checkbox:
            include_prodata = st.checkbox("📊 Продата", key="include_prodata")
        
        # Получаем данные ПроДата
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        prodata_plan_total = 0
        prodata_fact_total = 0
        prodata_plan_date_total = 0
        prodata_fact_date_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            prodata_plan_total = prodata_df['План проекта, шт.'].sum() if 'План проекта, шт.' in prodata_df.columns else 0
            prodata_fact_total = prodata_df['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in prodata_df.columns else 0
            prodata_plan_date_total = prodata_df['План на дату, шт.'].sum() if 'План на дату, шт.' in prodata_df.columns else 0
            prodata_fact_date_total = prodata_df['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in prodata_df.columns else 0
        
        # Первый ряд: План проекта, Факт проекта, План/Факт проекта
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_project_total = project_data['План проекта, шт.'].sum() if 'План проекта, шт.' in project_data.columns else 0
            if include_prodata:
                plan_project_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_project_total:,.0f} шт")
        
        with col2:
            fact_project_total = project_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in project_data.columns else 0
            if include_prodata:
                fact_project_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_project_total:,.0f} шт")
        
        with col3:
            pf_project_percent = (fact_project_total / plan_project_total * 100) if plan_project_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_project_percent:.1f}%")
        
        # Второй ряд: План на дату, Факт на дату, План/Факт на дату
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = project_data['План на дату, шт.'].sum() if 'План на дату, шт.' in project_data.columns else 0
            if include_prodata:
                plan_date_total += prodata_plan_date_total
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        
        with col5:
            fact_date_total = project_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in project_data.columns else 0
            if include_prodata:
                fact_date_total += prodata_fact_date_total
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
    

    def create_region_summary(self, df):
        """
        Агрегация данных по регионам
        Одна строка = один регион
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        # Определяем колонку региона
        region_col = 'Регион'
        if 'Регион short' in df.columns and 'Регион' not in df.columns:
            region_col = 'Регион short'
        
        if region_col not in df.columns:
            st.error(f"❌ В данных нет колонки региона")
            return pd.DataFrame()
        
        # Список колонок для агрегации
        agg_columns = {
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Длительность': 'mean',
            'Клиент': lambda x: ', '.join(x.dropna().unique()[:3]),
            'Проект': 'nunique',   # количество уникальных проектов
            'RS': 'nunique',        # количество сотрудников
            'ПО': lambda x: ', '.join(x.dropna().unique()[:3])  # уникальные ПО
        }
        
        # Только существующие колонки
        existing_agg = {}
        for k, v in agg_columns.items():
            if k in df.columns:
                existing_agg[k] = v
        
        # Группируем по региону
        region_agg = df.groupby(region_col).agg(existing_agg).reset_index()
        region_agg = region_agg.rename(columns={region_col: 'Регион'})
        
        # Переименовываем колонки для понятности
        rename_map = {
            'Проект': 'Кол-во проектов',
            'RS': 'Кол-во сотрудников'
        }
        region_agg = region_agg.rename(columns=rename_map)
        
        # 1. План/Факт на дату,%
        region_agg['План/Факт на дату,%'] = 0.0
        mask_plan = region_agg['План на дату, шт.'] > 0
        if mask_plan.any():
            region_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                region_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                region_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        # 2. План/Факт проекта,%
        region_agg['План/Факт проекта,%'] = 0.0
        mask_project_plan = region_agg['План проекта, шт.'] > 0
        if mask_project_plan.any():
            region_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                region_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                region_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        # 3. △План/Факт на дату, шт
        region_agg['△План/Факт на дату, шт'] = (
            region_agg['Факт на дату, шт.'] - region_agg['План на дату, шт.']
        ).round(1)
        
        # 4. △План/Факт на дату, %
        region_agg['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            region_agg.loc[mask_plan, '△План/Факт на дату, %'] = (
                (region_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                 region_agg.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # 5. Прогноз на месяц, шт.
        if 'plan_calc_params' in st.session_state:
            days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                            st.session_state['plan_calc_params']['start_date']).days + 1
        else:
            days_in_period = 12
            
        region_agg['Прогноз на месяц, шт.'] = (
            region_agg['Факт на дату, шт.'] / days_in_period * 28
        ).round(1)
        
        # Сортируем по региону
        region_agg = region_agg.sort_values('Регион')
        
        return region_agg
    
    def create_region_tab(self, data, hierarchy_df=None):
        """Создает вкладку Регионы с фильтрами и разверткой"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по регионам")
        
        # Переименовываем колонки для отображения
        rename_cols = {
            'ЗОД': 'DSM',
            'АСС': 'ASM',
            'ЭМ': 'RS'
        }
        data = data.rename(columns=rename_cols)
        
        # Определяем колонку региона
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # 🔍 ФИЛЬТРЫ (с возможностью исключения)
        with st.expander("🔍 Фильтры", expanded=True):
            
            # Получаем уникальные значения
            all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
            all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
            
            # Регионы
            if region_col in data.columns:
                unique_codes = data[region_col].dropna().unique()
                self.region_display_map = {}
                all_regions_display = []
                for code in unique_codes:
                    long_name = self._get_long_region(code)
                    self.region_display_map[long_name] = code
                    all_regions_display.append(long_name)
                all_regions_display.sort()
            else:
                all_regions_display = []
                self.region_display_map = {}
            
            all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
            
            # Переменные для хранения выбранных значений
            selected_dsm, excluded_dsm = [], []
            selected_asm, excluded_asm = [], []
            selected_region, excluded_region = [], []
            selected_client, excluded_client = [], []
            
            col1, col2, col3, col4 = st.columns(4)
            
            # === DSM ===
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio("Режим", ["Включить", "Исключить"], key="region_dsm_mode", horizontal=True)
                if dsm_mode == "Включить":
                    selected_dsm = st.multiselect("Выбрать", all_dsm, key="region_dsm_include")
                else:
                    excluded_dsm = st.multiselect("Исключить", all_dsm, key="region_dsm_exclude")
            
            # === ASM (с учетом выбранных DSM) ===
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio("Режим", ["Включить", "Исключить"], key="region_asm_mode", horizontal=True)
                
                if selected_dsm and 'DSM' in data.columns:
                    asm_options = sorted(data[data['DSM'].isin(selected_dsm)]['ASM'].dropna().unique())
                else:
                    asm_options = all_asm
                
                if asm_mode == "Включить":
                    selected_asm = st.multiselect("Выбрать", asm_options, key="region_asm_include")
                else:
                    excluded_asm = st.multiselect("Исключить", asm_options, key="region_asm_exclude")
            
            # === Регион ===
            with col3:
                st.markdown("**Регион**")
                region_mode = st.radio("Режим", ["Включить", "Исключить"], key="region_region_mode", horizontal=True)
                
                if region_mode == "Включить":
                    selected_region_display = st.multiselect("Выбрать", all_regions_display, key="region_region_include")
                    selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
                else:
                    excluded_region_display = st.multiselect("Исключить", all_regions_display, key="region_region_exclude")
                    excluded_region = [self.region_display_map.get(name, name) for name in excluded_region_display]
            
            # === Клиент (с учетом выбранных DSM, ASM, региона) ===
            with col4:
                st.markdown("**Клиент**")
                client_mode = st.radio("Режим", ["Включить", "Исключить"], key="region_client_mode", horizontal=True)
                
                client_filtered = data.copy()
                if selected_dsm and 'DSM' in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered['DSM'].isin(selected_dsm)]
                if selected_asm and 'ASM' in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered['ASM'].isin(selected_asm)]
                if selected_region and region_col in client_filtered.columns:
                    client_filtered = client_filtered[client_filtered[region_col].isin(selected_region)]
                client_options = sorted(client_filtered['Клиент'].dropna().unique()) if 'Клиент' in client_filtered.columns else all_clients
                
                if client_mode == "Включить":
                    selected_client = st.multiselect("Выбрать", client_options, key="region_client_include")
                else:
                    excluded_client = st.multiselect("Исключить", client_options, key="region_client_exclude")
        
        # === ПРИМЕНЯЕМ ФИЛЬТРЫ ===
        filtered_data = data.copy()
        
        # DSM
        if selected_dsm:
            filtered_data = filtered_data[filtered_data['DSM'].isin(selected_dsm)]
        if excluded_dsm:
            filtered_data = filtered_data[~filtered_data['DSM'].isin(excluded_dsm)]
        
        # ASM
        if selected_asm:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if excluded_asm:
            filtered_data = filtered_data[~filtered_data['ASM'].isin(excluded_asm)]
        
        # Регион
        if selected_region:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        if excluded_region:
            filtered_data = filtered_data[~filtered_data[region_col].isin(excluded_region)]
        
        # Клиент
        if selected_client:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        if excluded_client:
            filtered_data = filtered_data[~filtered_data['Клиент'].isin(excluded_client)]
        
        # 📊 РАЗВЕРТКА (ЧЕК-БОКСЫ)
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            show_project = st.checkbox("Проект", key='region_show_project')
        with col2:
            show_wave = st.checkbox("Волна", key='region_show_wave')
        with col3:
            show_dsm = st.checkbox("DSM", key='region_show_dsm')
        with col4:
            show_asm = st.checkbox("ASM", key='region_show_asm')
        with col5:
            show_rs = st.checkbox("RS", key='region_show_rs')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = [region_col]
        
        if show_project and 'Проект' in filtered_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in filtered_data.columns:
            group_cols.append('Волна')
        if show_dsm and 'DSM' in filtered_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        
        # Агрегируем данные с учетом развертки
        if len(group_cols) > 1:
            agg_columns = {
                'План проекта, шт.': 'sum',
                'План на дату, шт.': 'sum',
                'Факт проекта, шт.': 'sum',
                'Факт на дату, шт.': 'sum',
                'Длительность': 'mean',
                'ПО': lambda x: ', '.join(x.dropna().unique()[:3])
            }
            
            existing_agg = {k: v for k, v in agg_columns.items() if k in filtered_data.columns}
            detailed_data = filtered_data.groupby(group_cols).agg(existing_agg).reset_index()
            
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
            
            region_data = detailed_data
        else:
            region_data = self.create_region_summary(filtered_data)
        
        st.caption(f"📌 Отображается записей: {len(region_data)}")
        
        if region_data.empty:
            st.warning("⚠️ Нет данных после фильтрации")
            return
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_project_total = region_data['План проекта, шт.'].sum() if 'План проекта, шт.' in region_data.columns else 0
            st.metric("📊 План проекта", f"{plan_project_total:,.0f} шт")
        with col2:
            fact_project_total = region_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in region_data.columns else 0
            st.metric("✅ Факт проекта", f"{fact_project_total:,.0f} шт")
        with col3:
            pf_project_percent = (fact_project_total / plan_project_total * 100) if plan_project_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_project_percent:.1f}%")
        
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = region_data['План на дату, шт.'].sum() if 'План на дату, шт.' in region_data.columns else 0
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        with col5:
            fact_date_total = region_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in region_data.columns else 0
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
        
        # Колонки для отображения
        display_columns = [region_col]
        
        if show_project and 'Проект' in region_data.columns:
            display_columns.append('Проект')
        if show_wave and 'Волна' in region_data.columns:
            display_columns.append('Волна')
        if show_dsm and 'DSM' in region_data.columns:
            display_columns.append('DSM')
        if show_asm and 'ASM' in region_data.columns:
            display_columns.append('ASM')
        if show_rs and 'RS' in region_data.columns:
            display_columns.append('RS')
        
        metric_columns = [
            'План проекта, шт.',
            'Факт проекта, шт.',
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Прогноз на месяц, шт.'
        ]
        
        extra_columns = ['Кол-во клиентов', 'Кол-во проектов', 'Кол-во сотрудников', 'ПО']
        for col in extra_columns:
            if col in region_data.columns:
                metric_columns.append(col)
        
        display_columns.extend([col for col in metric_columns if col in region_data.columns])
        
        existing_cols = [col for col in display_columns if col in region_data.columns]
        df_display = region_data[existing_cols].copy()
        
        if region_col in df_display.columns:
            df_display[region_col] = df_display[region_col].apply(self._get_long_region)
        
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        if 'План/Факт проекта,%' in df_display.columns:
            df_display['План/Факт проекта,%'] = df_display['План/Факт проекта,%'].map(lambda x: f"{x:.1f}%")
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, sheet_name='Регионы', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"регионы_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    def create_dsm_summary(self, df):
        """
        Агрегация данных по DSM
        Одна строка = один DSM
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        if 'DSM' not in df.columns:
            st.error(f"❌ В данных нет колонки 'DSM'")
            return pd.DataFrame()
        
        # Список колонок для агрегации
        agg_columns = {
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Длительность': 'mean',
            'Клиент': lambda x: ', '.join(x.dropna().unique()[:3]),
            'Проект': 'nunique',
            'Регион': lambda x: ', '.join(x.dropna().unique()[:3]),
            'ASM': lambda x: ', '.join(x.dropna().unique()[:3]),
            'RS': 'nunique',
            'ПО': lambda x: ', '.join(x.dropna().unique()[:3])
        }
        
        # Только существующие колонки
        existing_agg = {}
        for k, v in agg_columns.items():
            if k in df.columns:
                existing_agg[k] = v
        
        # Группируем по DSM
        dsm_agg = df.groupby('DSM').agg(existing_agg).reset_index()
        
        # Переименовываем колонки для понятности
        rename_map = {
            'Проект': 'Кол-во проектов',
            'RS': 'Кол-во сотрудников'
        }
        dsm_agg = dsm_agg.rename(columns=rename_map)
        
        # 1. План/Факт на дату,%
        dsm_agg['План/Факт на дату,%'] = 0.0
        mask_plan = dsm_agg['План на дату, шт.'] > 0
        if mask_plan.any():
            dsm_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                dsm_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                dsm_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        # 2. План/Факт проекта,%
        dsm_agg['План/Факт проекта,%'] = 0.0
        mask_project_plan = dsm_agg['План проекта, шт.'] > 0
        if mask_project_plan.any():
            dsm_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                dsm_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                dsm_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        # 3. △План/Факт на дату, шт
        dsm_agg['△План/Факт на дату, шт'] = (
            dsm_agg['Факт на дату, шт.'] - dsm_agg['План на дату, шт.']
        ).round(1)
        
        # 4. △План/Факт на дату, %
        dsm_agg['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            dsm_agg.loc[mask_plan, '△План/Факт на дату, %'] = (
                (dsm_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                 dsm_agg.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # 5. Прогноз на месяц, шт.
        if 'plan_calc_params' in st.session_state:
            days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                            st.session_state['plan_calc_params']['start_date']).days + 1
        else:
            days_in_period = 12
            
        dsm_agg['Прогноз на месяц, шт.'] = (
            dsm_agg['Факт на дату, шт.'] / days_in_period * 28
        ).round(1)
        
        # Сортируем по DSM
        dsm_agg = dsm_agg.sort_values('DSM')
        
        return dsm_agg
    
    def create_dsm_tab(self, data, hierarchy_df=None):
        """Создает вкладку DSM с фильтрами и разверткой"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по DSM")
        
        # Переименовываем колонки для отображения
        rename_cols = {
            'ЗОД': 'DSM',
            'АСС': 'ASM',
            'ЭМ': 'RS'
        }
        data = data.rename(columns=rename_cols)
        
        # Определяем колонку региона
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # 🔍 ФИЛЬТРЫ (с возможностью исключения)
        with st.expander("🔍 Фильтры", expanded=True):
            
            # Получаем уникальные значения
            all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
            all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
            all_projects = sorted(data['Проект'].dropna().unique()) if 'Проект' in data.columns else []
            all_waves = sorted(data['Волна'].dropna().unique()) if 'Волна' in data.columns else []
            
            # Регионы
            if region_col in data.columns:
                unique_codes = data[region_col].dropna().unique()
                self.region_display_map = {}
                all_regions_display = []
                for code in unique_codes:
                    long_name = self._get_long_region(code)
                    self.region_display_map[long_name] = code
                    all_regions_display.append(long_name)
                all_regions_display.sort()
            else:
                all_regions_display = []
                self.region_display_map = {}
            
            # Переменные для хранения выбранных значений
            selected_asm, excluded_asm = [], []
            selected_client, excluded_client = [], []
            selected_project, excluded_project = [], []
            selected_wave, excluded_wave = [], []
            selected_region, excluded_region = [], []
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            # === ASM ===
            with col1:
                st.markdown("**ASM**")
                asm_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_asm_mode", horizontal=True)
                if asm_mode == "Включить":
                    selected_asm = st.multiselect("Выбрать", all_asm, key="dsm_asm_include")
                else:
                    excluded_asm = st.multiselect("Исключить", all_asm, key="dsm_asm_exclude")
            
            # === Клиент ===
            with col2:
                st.markdown("**Клиент**")
                client_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_client_mode", horizontal=True)
                if client_mode == "Включить":
                    selected_client = st.multiselect("Выбрать", all_clients, key="dsm_client_include")
                else:
                    excluded_client = st.multiselect("Исключить", all_clients, key="dsm_client_exclude")
            
            # === Проект ===
            with col3:
                st.markdown("**Код проекта**")
                project_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_project_mode", horizontal=True)
                if project_mode == "Включить":
                    selected_project = st.multiselect("Выбрать", all_projects, key="dsm_project_include")
                else:
                    excluded_project = st.multiselect("Исключить", all_projects, key="dsm_project_exclude")
            
            # === Волна ===
            with col4:
                st.markdown("**Волна**")
                wave_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_wave_mode", horizontal=True)
                if wave_mode == "Включить":
                    selected_wave = st.multiselect("Выбрать", all_waves, key="dsm_wave_include")
                else:
                    excluded_wave = st.multiselect("Исключить", all_waves, key="dsm_wave_exclude")
            
            # === Регион ===
            with col5:
                st.markdown("**Регион**")
                region_mode = st.radio("Режим", ["Включить", "Исключить"], key="dsm_region_mode", horizontal=True)
                if region_mode == "Включить":
                    selected_region_display = st.multiselect("Выбрать", all_regions_display, key="dsm_region_include")
                    selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
                else:
                    excluded_region_display = st.multiselect("Исключить", all_regions_display, key="dsm_region_exclude")
                    excluded_region = [self.region_display_map.get(name, name) for name in excluded_region_display]
        
        # === ПРИМЕНЯЕМ ФИЛЬТРЫ ===
        filtered_data = data.copy()
        
        if selected_asm:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if excluded_asm:
            filtered_data = filtered_data[~filtered_data['ASM'].isin(excluded_asm)]
        
        if selected_client:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        if excluded_client:
            filtered_data = filtered_data[~filtered_data['Клиент'].isin(excluded_client)]
        
        if selected_project:
            filtered_data = filtered_data[filtered_data['Проект'].isin(selected_project)]
        if excluded_project:
            filtered_data = filtered_data[~filtered_data['Проект'].isin(excluded_project)]
        
        if selected_wave:
            filtered_data = filtered_data[filtered_data['Волна'].isin(selected_wave)]
        if excluded_wave:
            filtered_data = filtered_data[~filtered_data['Волна'].isin(excluded_wave)]
        
        if selected_region:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        if excluded_region:
            filtered_data = filtered_data[~filtered_data[region_col].isin(excluded_region)]
        
        # 📊 РАЗВЕРТКА (ЧЕК-БОКСЫ)
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            show_asm = st.checkbox("ASM", key='dsm_show_asm')
        with col2:
            show_rs = st.checkbox("RS", key='dsm_show_rs')
        with col3:
            show_project = st.checkbox("Код проекта", key='dsm_show_project')
        with col4:
            show_wave = st.checkbox("Волна", key='dsm_show_wave')
        with col5:
            show_region = st.checkbox("Регион", key='dsm_show_region')
        
        # Формируем groupby
        group_cols = ['DSM']
        
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        if show_project and 'Проект' in filtered_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in filtered_data.columns:
            group_cols.append('Волна')
        if show_region and region_col in filtered_data.columns:
            group_cols.append(region_col)
        
        # Агрегируем
        if len(group_cols) > 1:
            agg_columns = {
                'План проекта, шт.': 'sum',
                'План на дату, шт.': 'sum',
                'Факт проекта, шт.': 'sum',
                'Факт на дату, шт.': 'sum',
                'Длительность': 'mean',
                'ПО': lambda x: ', '.join(x.dropna().unique()[:3])
            }
            
            existing_agg = {k: v for k, v in agg_columns.items() if k in filtered_data.columns}
            detailed_data = filtered_data.groupby(group_cols).agg(existing_agg).reset_index()
            
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
            
            dsm_data = detailed_data
        else:
            dsm_data = self.create_dsm_summary(filtered_data)
        
        st.caption(f"📌 Отображается записей: {len(dsm_data)}")
        
        if dsm_data.empty:
            st.warning("⚠️ Нет данных после фильтрации")
            return
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_project_total = dsm_data['План проекта, шт.'].sum() if 'План проекта, шт.' in dsm_data.columns else 0
            st.metric("📊 План проекта", f"{plan_project_total:,.0f} шт")
        with col2:
            fact_project_total = dsm_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in dsm_data.columns else 0
            st.metric("✅ Факт проекта", f"{fact_project_total:,.0f} шт")
        with col3:
            pf_project_percent = (fact_project_total / plan_project_total * 100) if plan_project_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_project_percent:.1f}%")
        
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = dsm_data['План на дату, шт.'].sum() if 'План на дату, шт.' in dsm_data.columns else 0
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        with col5:
            fact_date_total = dsm_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in dsm_data.columns else 0
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
        
        # Колонки для отображения
        display_columns = ['DSM']
        
        if show_asm and 'ASM' in dsm_data.columns:
            display_columns.append('ASM')
        if show_rs and 'RS' in dsm_data.columns:
            display_columns.append('RS')
        if show_project and 'Проект' in dsm_data.columns:
            display_columns.append('Проект')
        if show_wave and 'Волна' in dsm_data.columns:
            display_columns.append('Волна')
        if show_region and region_col in dsm_data.columns:
            display_columns.append(region_col)
        
        metric_columns = [
            'План проекта, шт.',
            'Факт проекта, шт.',
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Прогноз на месяц, шт.'
        ]
        
        extra_columns = ['Кол-во проектов', 'Кол-во сотрудников', 'ПО']
        for col in extra_columns:
            if col in dsm_data.columns:
                metric_columns.append(col)
        
        display_columns.extend([col for col in metric_columns if col in dsm_data.columns])
        
        existing_cols = [col for col in display_columns if col in dsm_data.columns]
        df_display = dsm_data[existing_cols].copy()
        
        if show_region and region_col in df_display.columns:
            df_display[region_col] = df_display[region_col].apply(self._get_long_region)
        
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        if 'План/Факт проекта,%' in df_display.columns:
            df_display['План/Факт проекта,%'] = df_display['План/Факт проекта,%'].map(lambda x: f"{x:.1f}%")
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, sheet_name='DSM', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"dsm_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
    def create_prodata_table(self, prodata_df):
        """
        Создает отдельную таблицу для ПроДата с возможностью развертки по Типу мониторинга
        Формат: Клиент | Тип мониторинга | Факт проекта
        """
        if prodata_df is None or prodata_df.empty:
            return
        
        st.markdown("---")
        st.subheader("📊 ПроДата (Мониторинги)")
        st.caption("Данные ПроДата не участвуют в основной таблице проектов")
        
        # Чек-бокс для развертки (по умолчанию выключен)
        show_detail = st.checkbox("📋 Показать детализацию по типам мониторинга", key="prodata_detail", value=False)
        
        if show_detail:
            # Развернутая таблица - показываем каждый тип мониторинга отдельно
            table_df = prodata_df[['Клиент', 'Тип мониторинга', 'Факт проекта, шт.']].copy()
            table_df = table_df.sort_values(['Клиент', 'Тип мониторинга'])
            
            # Форматирование
            table_df['Факт проекта, шт.'] = table_df['Факт проекта, шт.'].map(lambda x: f"{x:.1f}")
            
            # Отображаем таблицу
            st.dataframe(
                table_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'Клиент': 'Клиент',
                    'Тип мониторинга': 'Тип мониторинга',
                    'Факт проекта, шт.': st.column_config.TextColumn('Факт проекта, шт.')
                }
            )
            
            # Кнопка скачивания для развернутой таблицы
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                table_df.to_excel(writer, sheet_name='ПроДата_детально', index=False)
            
            st.download_button(
                label="⬇️ Скачать ПроДата (детально)",
                data=output.getvalue(),
                file_name=f"prodata_detailed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True
            )
        else:
            # Свернутая таблица - группируем по клиенту (суммируем все типы мониторинга)
            prodata_agg = prodata_df.groupby('Клиент')['Факт проекта, шт.'].sum().reset_index()
            prodata_agg = prodata_agg.sort_values('Клиент')
            
            # Форматирование
            prodata_agg['Факт проекта, шт.'] = prodata_agg['Факт проекта, шт.'].map(lambda x: f"{x:.1f}")
            
            # Отображаем таблицу
            st.dataframe(
                prodata_agg,
                use_container_width=True,
                hide_index=True,
                column_config={
                    'Клиент': 'Клиент',
                    'Факт проекта, шт.': st.column_config.TextColumn('Факт проекта, шт.')
                }
            )
            
            # Кнопка скачивания для свернутой таблицы
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                prodata_agg.to_excel(writer, sheet_name='ПроДата', index=False)
            
            st.download_button(
                label="⬇️ Скачать ПроДата",
                data=output.getvalue(),
                file_name=f"prodata_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True
            )

# Глобальный экземпляр
dataviz = DataVisualizer()
