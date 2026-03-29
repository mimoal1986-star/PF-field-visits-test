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

    def _compute_base_planfact_aggregations(self, data, region_col):
        """Вычисляет базовые агрегации для вкладки ПланФакт"""
        
        # Уникальные значения
        all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
        all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
        all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
        
        # Регионы
        if region_col in data.columns:
            unique_codes = data[region_col].dropna().unique()
            region_map = {}
            regions_display = []
            for code in unique_codes:
                long_name = self._get_long_region(code)
                region_map[long_name] = code
                regions_display.append(long_name)
            regions_display.sort()
        else:
            region_map = {}
            regions_display = []
        
        return {
            'raw_data': data.copy(),
            'all_dsm': all_dsm,
            'all_asm': all_asm,
            'all_clients': all_clients,
            'region_map': region_map,
            'regions_display': regions_display
        }

    def _apply_planfact_filters(self, data, dsm_selected, dsm_mode, asm_selected, asm_mode,region_selected, region_mode, client_selected, client_mode, region_col):
        """Применяет фильтры к данным"""
        filtered = data.copy()
        
        # DSM
        if dsm_selected:
            if dsm_mode == 'Включить':
                filtered = filtered[filtered['DSM'].isin(dsm_selected)]
            else:
                filtered = filtered[~filtered['DSM'].isin(dsm_selected)]
        
        # ASM
        if asm_selected:
            if asm_mode == 'Включить':
                filtered = filtered[filtered['ASM'].isin(asm_selected)]
            else:
                filtered = filtered[~filtered['ASM'].isin(asm_selected)]
        
        # Регион
        if region_selected:
            if region_mode == 'Включить':
                filtered = filtered[filtered[region_col].isin(region_selected)]
            else:
                filtered = filtered[~filtered[region_col].isin(region_selected)]
        
        # Клиент
        if client_selected:
            if client_mode == 'Включить':
                filtered = filtered[filtered['Клиент'].isin(client_selected)]
            else:
                filtered = filtered[~filtered['Клиент'].isin(client_selected)]
        
        return filtered

    def _compute_base_region_aggregations(self, data, region_col):
        """Вычисляет базовые агрегации для вкладки Регионы"""
        
        # Уникальные значения
        all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
        all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
        all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
        
        # Регионы
        if region_col in data.columns:
            unique_codes = data[region_col].dropna().unique()
            region_map = {}
            regions_display = []
            for code in unique_codes:
                long_name = self._get_long_region(code)
                region_map[long_name] = code
                regions_display.append(long_name)
            regions_display.sort()
        else:
            region_map = {}
            regions_display = []
        
        return {
            'raw_data': data.copy(),
            'all_dsm': all_dsm,
            'all_asm': all_asm,
            'all_clients': all_clients,
            'region_map': region_map,
            'regions_display': regions_display
        }

    def _compute_base_dsm_aggregations(self, data, region_col):
        """Вычисляет базовые агрегации для вкладки DSM"""
        
        # Уникальные значения
        all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
        all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
        all_projects = sorted(data['Проект'].dropna().unique()) if 'Проект' in data.columns else []
        all_waves = sorted(data['Волна'].dropna().unique()) if 'Волна' in data.columns else []
        
        # Регионы
        if region_col in data.columns:
            unique_codes = data[region_col].dropna().unique()
            region_map = {}
            regions_display = []
            for code in unique_codes:
                long_name = self._get_long_region(code)
                region_map[long_name] = code
                regions_display.append(long_name)
            regions_display.sort()
        else:
            region_map = {}
            regions_display = []
        
        return {
            'raw_data': data.copy(),
            'all_asm': all_asm,
            'all_clients': all_clients,
            'all_projects': all_projects,
            'all_waves': all_waves,
            'region_map': region_map,
            'regions_display': regions_display
        }

    def _apply_region_filters(self, data, dsm_selected, dsm_mode, asm_selected, asm_mode,
                              region_selected, region_mode, client_selected, client_mode, region_col):
        """Применяет фильтры к данным для вкладки Регионы"""
        filtered = data.copy()
        
        # DSM
        if dsm_selected:
            if dsm_mode == 'Включить':
                filtered = filtered[filtered['DSM'].isin(dsm_selected)]
            else:
                filtered = filtered[~filtered['DSM'].isin(dsm_selected)]
        
        # ASM
        if asm_selected:
            if asm_mode == 'Включить':
                filtered = filtered[filtered['ASM'].isin(asm_selected)]
            else:
                filtered = filtered[~filtered['ASM'].isin(asm_selected)]
        
        # Регион
        if region_selected:
            if region_mode == 'Включить':
                filtered = filtered[filtered[region_col].isin(region_selected)]
            else:
                filtered = filtered[~filtered[region_col].isin(region_selected)]
        
        # Клиент
        if client_selected:
            if client_mode == 'Включить':
                filtered = filtered[filtered['Клиент'].isin(client_selected)]
            else:
                filtered = filtered[~filtered['Клиент'].isin(client_selected)]
        
        return filtered
    
    def _apply_dsm_filters(self, data, asm_selected, asm_mode, client_selected, client_mode,
                           project_selected, project_mode, wave_selected, wave_mode,
                           region_selected, region_mode, region_col):
        """Применяет фильтры к данным для вкладки DSM"""
        filtered = data.copy()
        
        # ASM
        if asm_selected:
            if asm_mode == 'Включить':
                filtered = filtered[filtered['ASM'].isin(asm_selected)]
            else:
                filtered = filtered[~filtered['ASM'].isin(asm_selected)]
        
        # Клиент
        if client_selected:
            if client_mode == 'Включить':
                filtered = filtered[filtered['Клиент'].isin(client_selected)]
            else:
                filtered = filtered[~filtered['Клиент'].isin(client_selected)]
        
        # Проект
        if project_selected:
            if project_mode == 'Включить':
                filtered = filtered[filtered['Проект'].isin(project_selected)]
            else:
                filtered = filtered[~filtered['Проект'].isin(project_selected)]
        
        # Волна
        if wave_selected:
            if wave_mode == 'Включить':
                filtered = filtered[filtered['Волна'].isin(wave_selected)]
            else:
                filtered = filtered[~filtered['Волна'].isin(wave_selected)]
        
        # Регион
        if region_selected:
            if region_mode == 'Включить':
                filtered = filtered[filtered[region_col].isin(region_selected)]
            else:
                filtered = filtered[~filtered[region_col].isin(region_selected)]
        
        return filtered
    
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
        """Создает вкладку ПланФакт на дату с фильтрами в форме"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по проектам")
        
        # Переименовываем колонки
        rename_cols = {'ЗОД': 'DSM', 'АСС': 'ASM', 'ЭМ': 'RS'}
        data = data.rename(columns=rename_cols)
        
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # ============================================
        # 1. БАЗОВЫЕ АГРЕГАЦИИ (один раз за сессию)
        # ============================================
        
        # Проверяем, изменились ли данные
        data_hash = hash(data.values.tobytes()) if not data.empty else 0
        
        if 'planfact_base_data' not in st.session_state or st.session_state.get('planfact_data_hash') != data_hash:
            # Вычисляем базовые агрегации
            base_agg = self._compute_base_planfact_aggregations(data, region_col)
            st.session_state.planfact_base_data = base_agg
            st.session_state.planfact_data_hash = data_hash
            st.session_state.planfact_filtered_data = None  # сбрасываем фильтры
        
        base_data = st.session_state.planfact_base_data
        
        # ============================================
        # 2. ФИЛЬТРЫ В ФОРМЕ (без rerun при выборе)
        # ============================================
        
        with st.form("planfact_filters_form"):
            
            # Получаем уникальные значения из базовых данных
            all_dsm = base_data['all_dsm']
            all_asm = base_data['all_asm']
            all_clients = base_data['all_clients']
            all_regions_display = base_data['regions_display']
            region_map = base_data['region_map']
            
            st.markdown("### 🔍 Фильтры")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="planfact_dsm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('planfact_dsm_mode', 'Включить') == 'Включить' else 1
                )
                # Читаем значение из session_state (не присваиваем)
                dsm_mode = st.session_state.planfact_dsm_mode
                
                dsm_selected = st.multiselect(
                    "Выбрать DSM",
                    all_dsm,
                    key="planfact_dsm_values",
                    default=[]
                )
            
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="planfact_asm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('planfact_asm_mode', 'Включить') == 'Включить' else 1
                )
                asm_mode = st.session_state.planfact_asm_mode
                
                if dsm_selected:
                    asm_options = [a for a in all_asm if any(a in d for d in dsm_selected)]
                else:
                    asm_options = all_asm
                
                asm_selected = st.multiselect(
                    "Выбрать ASM",
                    asm_options,
                    key="planfact_asm_values",
                    default=[]
                )
            
            with col3:
                st.markdown("**Регион**")
                region_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="planfact_region_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('planfact_region_mode', 'Включить') == 'Включить' else 1
                )
                region_mode = st.session_state.planfact_region_mode
                
                region_selected_display = st.multiselect(
                    "Выбрать регион",
                    all_regions_display,
                    key="planfact_region_values",
                    default=[]
                )
                region_selected = [region_map.get(name, name) for name in region_selected_display]
            
            with col4:
                st.markdown("**Клиент**")
                client_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="planfact_client_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('planfact_client_mode', 'Включить') == 'Включить' else 1
                )
                client_mode = st.session_state.planfact_client_mode
                
                # Ограничиваем клиентов выбранными фильтрами
                client_filtered = all_clients
                if dsm_selected:
                    client_filtered = [c for c in client_filtered if c in dsm_selected]
                
                client_selected = st.multiselect(
                    "Выбрать клиента",
                    client_filtered,
                    key="planfact_client_values",
                    default=[]
                )
            
            st.markdown("---")
            
            # КНОПКА ПРИМЕНЕНИЯ ФИЛЬТРОВ
            apply_filters = st.form_submit_button(
                "✅ Применить фильтры",
                type="primary",
                use_container_width=True,
                key="planfact_apply_filters_btn"
            )
        
        # ============================================
        # 3. ПРИМЕНЯЕМ ФИЛЬТРЫ (только по кнопке)
        # ============================================
        
        if apply_filters:
            filtered_data = self._apply_planfact_filters(
                base_data['raw_data'],
                dsm_selected, st.session_state.planfact_dsm_mode,
                asm_selected, st.session_state.planfact_asm_mode,
                region_selected, st.session_state.planfact_region_mode,
                client_selected, st.session_state.planfact_client_mode,
                region_col
            )
            st.session_state.planfact_filtered_data = filtered_data
        
        # Берем данные для отображения
        if st.session_state.planfact_filtered_data is not None:
            display_data = st.session_state.planfact_filtered_data
        else:
            display_data = base_data['raw_data']
        
        # ============================================
        # 4. РАЗВЕРТКА И ОТОБРАЖЕНИЕ
        # ============================================
        
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            show_project = st.checkbox("Код проекта", key='planfact_show_project')
        with col2:
            show_regions = st.checkbox("Регионы", key='planfact_show_regions')
        with col3:
            show_dsm = st.checkbox("DSM", key='planfact_show_dsm')
        with col4:
            show_asm = st.checkbox("ASM", key='planfact_show_asm')
        with col5:
            show_rs = st.checkbox("RS", key='planfact_show_rs')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Клиент', 'ПО']
        
        if show_project and 'Проект' in display_data.columns:
            group_cols.append('Проект')
        if show_regions and region_col in display_data.columns:
            group_cols.append(region_col)
        if show_dsm and 'DSM' in display_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in display_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in display_data.columns:
            group_cols.append('RS')
        
        # Агрегируем для отображения
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
        
        existing_agg = {k: v for k, v in agg_columns.items() if k in display_data.columns}
        
        if len(group_cols) > 1:
            project_data = display_data.groupby(group_cols).agg(existing_agg).reset_index()
        else:
            project_data = display_data
        
        # Добавляем вычисляемые метрики
        mask_plan = project_data['План на дату, шт.'] > 0
        project_data['План/Факт на дату,%'] = 0.0
        if mask_plan.any():
            project_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                project_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                project_data.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        mask_project_plan = project_data['План проекта, шт.'] > 0
        project_data['План/Факт проекта,%'] = 0.0
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
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        col_kpi1, col_kpi2, col_kpi3, col_checkbox = st.columns([1, 1, 1, 0.5])
        with col_checkbox:
            include_prodata = st.checkbox("📊 Продата", key="planfact_include_prodata")
        
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        prodata_plan_total = 0
        prodata_fact_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            prodata_plan_total = prodata_df['План проекта, шт.'].sum() if 'План проекта, шт.' in prodata_df.columns else 0
            prodata_fact_total = prodata_df['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in prodata_df.columns else 0
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = project_data['План проекта, шт.'].sum() if 'План проекта, шт.' in project_data.columns else 0
            if include_prodata:
                plan_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = project_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in project_data.columns else 0
            if include_prodata:
                fact_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_percent:.1f}%")
        
        # Отображаем таблицу
        display_cols = ['Клиент', 'ПО', 'План проекта, шт.', 'Факт проекта, шт.', 'План/Факт проекта,%']
        existing_display = [c for c in display_cols if c in project_data.columns]
        st.dataframe(project_data[existing_display], use_container_width=True, hide_index=True)
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            project_data.to_excel(writer, sheet_name='План_факт_проекты', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"план_факт_проекты_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    def create_region_summary(self, df):
        """
        Агрегация данных по регионам
        Одна строка = один регион
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        region_col = 'Регион'
        if 'Регион short' in df.columns and 'Регион' not in df.columns:
            region_col = 'Регион short'
        
        if region_col not in df.columns:
            st.error(f"❌ В данных нет колонки региона")
            return pd.DataFrame()
        
        agg_columns = {
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Длительность': 'mean',
            'Клиент': lambda x: ', '.join(x.dropna().unique()[:3]),
            'Проект': 'nunique',
            'RS': 'nunique',
            'ПО': lambda x: ', '.join(x.dropna().unique()[:3])
        }
        
        existing_agg = {}
        for k, v in agg_columns.items():
            if k in df.columns:
                existing_agg[k] = v
        
        region_agg = df.groupby(region_col).agg(existing_agg).reset_index()
        region_agg = region_agg.rename(columns={region_col: 'Регион'})
        
        rename_map = {'Проект': 'Кол-во проектов', 'RS': 'Кол-во сотрудников'}
        region_agg = region_agg.rename(columns=rename_map)
        
        region_agg['План/Факт на дату,%'] = 0.0
        mask_plan = region_agg['План на дату, шт.'] > 0
        if mask_plan.any():
            region_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                region_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                region_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        region_agg['План/Факт проекта,%'] = 0.0
        mask_project_plan = region_agg['План проекта, шт.'] > 0
        if mask_project_plan.any():
            region_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                region_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                region_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        region_agg['△План/Факт на дату, шт'] = (
            region_agg['Факт на дату, шт.'] - region_agg['План на дату, шт.']
        ).round(1)
        
        region_agg['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            region_agg.loc[mask_plan, '△План/Факт на дату, %'] = (
                (region_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                 region_agg.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        if 'plan_calc_params' in st.session_state:
            days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
                            st.session_state['plan_calc_params']['start_date']).days + 1
        else:
            days_in_period = 12
            
        region_agg['Прогноз на месяц, шт.'] = (
            region_agg['Факт на дату, шт.'] / days_in_period * 28
        ).round(1)
        
        region_agg = region_agg.sort_values('Регион')
        return region_agg

    def create_region_tab(self, data, hierarchy_df=None):
        """Создает вкладку Регионы с фильтрами в форме"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по регионам")
        
        # Переименовываем колонки
        rename_cols = {'ЗОД': 'DSM', 'АСС': 'ASM', 'ЭМ': 'RS'}
        data = data.rename(columns=rename_cols)
        
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # ============================================
        # 1. БАЗОВЫЕ АГРЕГАЦИИ (один раз за сессию)
        # ============================================
        
        data_hash = hash(data.values.tobytes()) if not data.empty else 0
        
        if 'region_base_data' not in st.session_state or st.session_state.get('region_data_hash') != data_hash:
            base_agg = self._compute_base_region_aggregations(data, region_col)
            st.session_state.region_base_data = base_agg
            st.session_state.region_data_hash = data_hash
            st.session_state.region_filtered_data = None
        
        base_data = st.session_state.region_base_data
        
        # ============================================
        # 2. ФИЛЬТРЫ В ФОРМЕ
        # ============================================
        
        with st.form("region_filters_form"):
            
            all_dsm = base_data['all_dsm']
            all_asm = base_data['all_asm']
            all_clients = base_data['all_clients']
            all_regions_display = base_data['regions_display']
            region_map = base_data['region_map']
            
            st.markdown("### 🔍 Фильтры")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="region_dsm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('region_dsm_mode', 'Включить') == 'Включить' else 1
                )
                dsm_mode = st.session_state.region_dsm_mode
                
                dsm_selected = st.multiselect(
                    "Выбрать DSM",
                    all_dsm,
                    key="region_dsm_values",
                    default=[]
                )
            
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="region_asm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('region_asm_mode', 'Включить') == 'Включить' else 1
                )
                asm_mode = st.session_state.region_asm_modee
                
                if dsm_selected:
                    asm_options = [a for a in all_asm if any(a in d for d in dsm_selected)]
                else:
                    asm_options = all_asm
                
                asm_selected = st.multiselect(
                    "Выбрать ASM",
                    asm_options,
                    key="region_asm_values",
                    default=[]
                )
            
            with col3:
                st.markdown("**Регион**")
                region_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="region_region_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('region_region_mode', 'Включить') == 'Включить' else 1
                )
                region_mode = st.session_state.region_region_mode
                
                region_selected_display = st.multiselect(
                    "Выбрать регион",
                    all_regions_display,
                    key="region_region_values",
                    default=[]
                )
                region_selected = [region_map.get(name, name) for name in region_selected_display]
            
            with col4:
                st.markdown("**Клиент**")
                client_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="region_client_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('region_client_mode', 'Включить') == 'Включить' else 1
                )
                client_mode = st.session_state.region_client_mode
                
                client_selected = st.multiselect(
                    "Выбрать клиента",
                    all_clients,
                    key="region_client_values",
                    default=[]
                )
            
            st.markdown("---")
            
            apply_filters = st.form_submit_button(
                "✅ Применить фильтры",
                type="primary",
                use_container_width=True,
                key="region_apply_filters_btn"
            )
        
        # ============================================
        # 3. ПРИМЕНЯЕМ ФИЛЬТРЫ
        # ============================================
        
        if apply_filters:
            filtered_data = self._apply_region_filters(
                base_data['raw_data'],
                dsm_selected, st.session_state.region_dsm_mode,
                asm_selected, st.session_state.region_asm_mode,
                region_selected, st.session_state.region_region_mode,
                client_selected, st.session_state.region_client_mode,
                region_col
            )
            st.session_state.region_filtered_data = filtered_data
        
        if st.session_state.region_filtered_data is not None:
            display_data = st.session_state.region_filtered_data
        else:
            display_data = base_data['raw_data']
        
        # ============================================
        # 4. АГРЕГАЦИЯ ПО РЕГИОНАМ
        # ============================================
        
        # Группируем по региону
        region_agg = display_data.groupby(region_col).agg({
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Проект': 'nunique',
            'RS': 'nunique'
        }).reset_index()
        
        region_agg = region_agg.rename(columns={
            'Проект': 'Кол-во проектов',
            'RS': 'Кол-во сотрудников'
        })
        
        # Добавляем метрики
        mask_plan = region_agg['План на дату, шт.'] > 0
        region_agg['План/Факт на дату,%'] = 0.0
        if mask_plan.any():
            region_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                region_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                region_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        mask_project_plan = region_agg['План проекта, шт.'] > 0
        region_agg['План/Факт проекта,%'] = 0.0
        if mask_project_plan.any():
            region_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                region_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                region_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        region_agg['△План/Факт на дату, шт'] = (
            region_agg['Факт на дату, шт.'] - region_agg['План на дату, шт.']
        ).round(1)
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        include_prodata = st.checkbox("📊 Продата", key="region_include_prodata")
        
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        prodata_plan_total = 0
        prodata_fact_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            prodata_plan_total = prodata_df['План проекта, шт.'].sum() if 'План проекта, шт.' in prodata_df.columns else 0
            prodata_fact_total = prodata_df['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in prodata_df.columns else 0
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = region_agg['План проекта, шт.'].sum() if 'План проекта, шт.' in region_agg.columns else 0
            if include_prodata:
                plan_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = region_agg['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in region_agg.columns else 0
            if include_prodata:
                fact_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_percent:.1f}%")
        
        # Отображаем таблицу
        display_cols = ['Регион', 'План проекта, шт.', 'Факт проекта, шт.', 'План/Факт проекта,%',
                        'Кол-во проектов', 'Кол-во сотрудников']
        existing_display = [c for c in display_cols if c in region_agg.columns]
        
        # Преобразуем коды регионов в названия
        if 'Регион' in region_agg.columns:
            region_agg['Регион'] = region_agg['Регион'].apply(self._get_long_region)
        
        st.dataframe(region_agg[existing_display], use_container_width=True, hide_index=True)
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            region_agg.to_excel(writer, sheet_name='Регионы', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"регионы_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    
    def create_dsm_tab(self, data, hierarchy_df=None):
        """Создает вкладку DSM с фильтрами в форме"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по DSM")
        
        # Переименовываем колонки
        rename_cols = {'ЗОД': 'DSM', 'АСС': 'ASM', 'ЭМ': 'RS'}
        data = data.rename(columns=rename_cols)
        
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # ============================================
        # 1. БАЗОВЫЕ АГРЕГАЦИИ (один раз за сессию)
        # ============================================
        
        data_hash = hash(data.values.tobytes()) if not data.empty else 0
        
        if 'dsm_base_data' not in st.session_state or st.session_state.get('dsm_data_hash') != data_hash:
            base_agg = self._compute_base_dsm_aggregations(data, region_col)
            st.session_state.dsm_base_data = base_agg
            st.session_state.dsm_data_hash = data_hash
            st.session_state.dsm_filtered_data = None
        
        base_data = st.session_state.dsm_base_data
        
        # ============================================
        # 2. ФИЛЬТРЫ В ФОРМЕ
        # ============================================
        
        with st.form("dsm_filters_form"):
            
            all_asm = base_data['all_asm']
            all_clients = base_data['all_clients']
            all_projects = base_data['all_projects']
            all_waves = base_data['all_waves']
            all_regions_display = base_data['regions_display']
            region_map = base_data['region_map']
            
            st.markdown("### 🔍 Фильтры")
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.markdown("**ASM**")
                asm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_asm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_asm_mode', 'Включить') == 'Включить' else 1
                )
                st.session_state.dsm_asm_mode = asm_mode
                
                asm_selected = st.multiselect(
                    "Выбрать ASM",
                    all_asm,
                    key="dsm_asm_values",
                    default=[]
                )
            
            with col2:
                st.markdown("**Клиент**")
                client_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_client_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_client_mode', 'Включить') == 'Включить' else 1
                )
                st.session_state.dsm_client_mode = client_mode
                
                client_selected = st.multiselect(
                    "Выбрать клиента",
                    all_clients,
                    key="dsm_client_values",
                    default=[]
                )
            
            with col3:
                st.markdown("**Код проекта**")
                project_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_project_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_project_mode', 'Включить') == 'Включить' else 1
                )
                st.session_state.dsm_project_mode = project_mode
                
                project_selected = st.multiselect(
                    "Выбрать проект",
                    all_projects,
                    key="dsm_project_values",
                    default=[]
                )
            
            with col4:
                st.markdown("**Волна**")
                wave_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_wave_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_wave_mode', 'Включить') == 'Включить' else 1
                )
                st.session_state.dsm_wave_mode = wave_mode
                
                wave_selected = st.multiselect(
                    "Выбрать волну",
                    all_waves,
                    key="dsm_wave_values",
                    default=[]
                )
            
            with col5:
                st.markdown("**Регион**")
                region_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_region_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_region_mode', 'Включить') == 'Включить' else 1
                )
                st.session_state.dsm_region_mode = region_mode
                
                region_selected_display = st.multiselect(
                    "Выбрать регион",
                    all_regions_display,
                    key="dsm_region_values",
                    default=[]
                )
                region_selected = [region_map.get(name, name) for name in region_selected_display]
            
            st.markdown("---")
            
            apply_filters = st.form_submit_button(
                "✅ Применить фильтры",
                type="primary",
                use_container_width=True,
                key="dsm_apply_filters_btn"
            )
        
        # ============================================
        # 3. ПРИМЕНЯЕМ ФИЛЬТРЫ
        # ============================================
        
        if apply_filters:
            filtered_data = self._apply_dsm_filters(
                base_data['raw_data'],
                asm_selected, st.session_state.dsm_asm_mode,
                client_selected, st.session_state.dsm_client_mode,
                project_selected, st.session_state.dsm_project_mode,
                wave_selected, st.session_state.dsm_wave_mode,
                region_selected, st.session_state.dsm_region_mode,
                region_col
            )
            st.session_state.dsm_filtered_data = filtered_data
        
        if st.session_state.dsm_filtered_data is not None:
            display_data = st.session_state.dsm_filtered_data
        else:
            display_data = base_data['raw_data']
        
        # ============================================
        # 4. АГРЕГАЦИЯ ПО DSM
        # ============================================
        
        # Группируем по DSM
        dsm_agg = display_data.groupby('DSM').agg({
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт проекта, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Проект': 'nunique',
            'RS': 'nunique'
        }).reset_index()
        
        dsm_agg = dsm_agg.rename(columns={
            'Проект': 'Кол-во проектов',
            'RS': 'Кол-во сотрудников'
        })
        
        # Добавляем метрики
        mask_plan = dsm_agg['План на дату, шт.'] > 0
        dsm_agg['План/Факт на дату,%'] = 0.0
        if mask_plan.any():
            dsm_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
                dsm_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                dsm_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        mask_project_plan = dsm_agg['План проекта, шт.'] > 0
        dsm_agg['План/Факт проекта,%'] = 0.0
        if mask_project_plan.any():
            dsm_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                dsm_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                dsm_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        dsm_agg['△План/Факт на дату, шт'] = (
            dsm_agg['Факт на дату, шт.'] - dsm_agg['План на дату, шт.']
        ).round(1)
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        include_prodata = st.checkbox("📊 Продата", key="dsm_include_prodata")
        
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        prodata_plan_total = 0
        prodata_fact_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            prodata_plan_total = prodata_df['План проекта, шт.'].sum() if 'План проекта, шт.' in prodata_df.columns else 0
            prodata_fact_total = prodata_df['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in prodata_df.columns else 0
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = dsm_agg['План проекта, шт.'].sum() if 'План проекта, шт.' in dsm_agg.columns else 0
            if include_prodata:
                plan_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = dsm_agg['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in dsm_agg.columns else 0
            if include_prodata:
                fact_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_percent:.1f}%")
        
        # Отображаем таблицу
        display_cols = ['DSM', 'План проекта, шт.', 'Факт проекта, шт.', 'План/Факт проекта,%',
                        'Кол-во проектов', 'Кол-во сотрудников']
        existing_display = [c for c in display_cols if c in dsm_agg.columns]
        
        st.dataframe(dsm_agg[existing_display], use_container_width=True, hide_index=True)
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dsm_agg.to_excel(writer, sheet_name='DSM', index=False)
        
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
