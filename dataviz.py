# utils/dataviz.py
# draft 4.1 
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO
from visit_calculator import visit_calculator

class DataVisualizer:

    def __init__(self):
        self.region_mapping = {
            'AD': 'Республика Адыгея', 'AL': 'Алтайский край', 'AM': 'Амурская область',
            'AR': 'Архангельская область', 'AS': 'Астраханская область', 'BK': 'Республика Башкортостан',
            'BL': 'Белгородская область', 'BR': 'Брянская область', 'BU': 'Республика Бурятия',
            'CK': 'Чукотский автономный округ', 'CL': 'Челябинская область', 'CN': 'Чеченская Республика',
            'CV': 'Чувашская Республика', 'DA': 'Республика Дагестан', 'DN': 'Донецкая Народная Республика',
            'GA': 'Республика Алтай', 'IN': 'Республика Ингушетия', 'IR': 'Иркутская область',
            'IV': 'Ивановская область', 'KA': 'Камчатский край', 'KB': 'Кабардино-Балкарская Республика',
            'KC': 'Карачаево-Черкесская Республика', 'KD': 'Краснодарский край', 'KE': 'Кемеровская область',
            'KG': 'Калужская область', 'KH': 'Хабаровский край', 'KI': 'Республика Карелия',
            'KK': 'Республика Хакасия', 'KL': 'Республика Калмыкия', 'KM': 'Ханты-Мансийский автономный округ',
            'KN': 'Калининградская область', 'KO': 'Республика Коми', 'KS': 'Курская область',
            'KT': 'Костромская область', 'KU': 'Курганская область', 'KV': 'Кировская область',
            'KY': 'Красноярский край', 'LG': 'Луганская Народная Республика', 'LN': 'Ленинградская область',
            'LP': 'Липецкая область', 'MC': 'Московская область', 'ME': 'Республика Марий Эл',
            'MG': 'Магаданская область', 'MM': 'Мурманская область', 'MR': 'Республика Мордовия',
            'MS': 'Московская область', 'NG': 'Новгородская область', 'NN': 'Ненецкий автономный округ',
            'NO': 'Республика Северная Осетия', 'NS': 'Новосибирская область', 'NZ': 'Нижегородская область',
            'OB': 'Оренбургская область', 'OL': 'Орловская область', 'OM': 'Омская область',
            'PE': 'Пермский край', 'PR': 'Приморский край', 'PS': 'Псковская область',
            'PZ': 'Пензенская область', 'RK': 'Республика Крым', 'RO': 'Ростовская область',
            'RZ': 'Рязанская область', 'SA': 'Самарская область', 'SK': 'Республика Саха (Якутия)',
            'SL': 'Сахалинская область', 'SM': 'Смоленская область', 'SR': 'Саратовская область',
            'ST': 'Ставропольский край', 'SV': 'Свердловская область', 'TB': 'Тамбовская область',
            'TL': 'Тульская область', 'TO': 'Томская область', 'TT': 'Республика Татарстан',
            'TU': 'Республика Тыва', 'TV': 'Тверская область', 'TY': 'Тюменская область',
            'UD': 'Удмуртская Республика', 'UL': 'Ульяновская область', 'VG': 'Волгоградская область',
            'VL': 'Владимирская область', 'VO': 'Вологодская область', 'VR': 'Воронежская область',
            'YN': 'Ямало-Ненецкий автономный округ', 'YS': 'Ярославская область', 'YV': 'Еврейская автономная область',
            'ZK': 'Забайкальский край', 'ZO': 'Запорожская область'
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
        all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
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
            'all_dsm': all_dsm,
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
    

    def _apply_dsm_filters(self, data, dsm_selected, dsm_mode, asm_selected, asm_mode,
                           client_selected, client_mode, region_selected, region_mode, region_col):
        """Применяет фильтры к данным для вкладки DSM"""
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
        
        # Клиент
        if client_selected:
            if client_mode == 'Включить':
                filtered = filtered[filtered['Клиент'].isin(client_selected)]
            else:
                filtered = filtered[~filtered['Клиент'].isin(client_selected)]
        
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

    # def create_project_summary(self, df):
    #     """
    #     Агрегация данных по проектам
    #     Одна строка = один проект
    #     """
    #     if df is None or df.empty:
    #         return pd.DataFrame()
        
    #     # Проверяем наличие колонки 'Проект'
    #     project_col = 'Проект'
    #     if project_col not in df.columns:
    #         st.error(f"❌ В данных нет колонки '{project_col}'")
    #         return pd.DataFrame()
        
    #     # Список колонок для агрегации
    #     agg_columns = {
    #         'План проекта, шт.': 'sum',
    #         'План на дату, шт.': 'sum',
    #         'Факт проекта, шт.': 'sum',
    #         'Факт на дату, шт.': 'sum',
    #         'Длительность': 'first',
    #         'Дата старта': 'first',
    #         'Дата финиша': 'first',
    #         'Клиент': 'first',
    #         'ПО': 'first',
    #         'Дней до конца проекта': 'first',
    #         'Утилизация тайминга, %': 'first',
    #         'Ср. план на день для 100% плана': 'sum'
    #     }
        
    #     # Только существующие колонки
    #     existing_agg = {k: v for k, v in agg_columns.items() if k in df.columns}
        
    #     # Группируем по проекту
    #     project_agg = df.groupby(project_col).agg(existing_agg).reset_index()
        
    #     # 1. План/Факт на дату,% (было Исполнение проекта,%)
    #     project_agg['План/Факт на дату,%'] = 0.0
    #     mask_plan = project_agg['План на дату, шт.'] > 0
    #     if mask_plan.any():
    #         project_agg.loc[mask_plan, 'План/Факт на дату,%'] = (
    #             project_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
    #             project_agg.loc[mask_plan, 'План на дату, шт.'] * 100
    #         ).round(1)
        
    #     # 2. △План/Факт на дату, шт (исправленная формула: Факт - План)
    #     project_agg['△План/Факт на дату, шт'] = (
    #         project_agg['Факт на дату, шт.'] - project_agg['План на дату, шт.']
    #     ).round(1)
        
    #     # 3. △План/Факт на дату, % (исправленная формула)
    #     project_agg['△План/Факт на дату, %'] = 0.0
    #     if mask_plan.any():
    #         project_agg.loc[mask_plan, '△План/Факт на дату, %'] = (
    #             (project_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
    #              project_agg.loc[mask_plan, 'План на дату, шт.']) - 1
    #         ).round(3) * 100
        
    #     # 4. Исполнение проекта,% (было %ПФ проекта)
    #     project_agg['План/Факт проекта,%'] = 0.0
    #     mask_project_plan = project_agg['План проекта, шт.'] > 0
    #     if mask_project_plan.any():
    #         project_agg.loc[mask_project_plan, 'План/Факт проекта,%'] = (
    #             project_agg.loc[mask_project_plan, 'Факт проекта, шт.'] / 
    #             project_agg.loc[mask_project_plan, 'План проекта, шт.'] * 100
    #         ).round(1)
        
    #     # 5. Прогноз на месяц, шт.
    #     if 'plan_calc_params' in st.session_state:
    #         days_in_period = (st.session_state['plan_calc_params']['end_date'] - 
    #                         st.session_state['plan_calc_params']['start_date']).days + 1
    #     else:
    #         days_in_period = 12
            
    #     project_agg['Прогноз на месяц, шт.'] = (
    #         project_agg['Факт на дату, шт.'] / days_in_period * 28
    #     ).round(1)
        
    #     # 6. Фокус
    #     project_agg['Фокус'] = 'Нет'
    #     if all(col in project_agg.columns for col in ['План/Факт проекта,%', 'Утилизация тайминга, %']):
    #         mask_focus = (
    #             (project_agg['План/Факт проекта,%'] < 80) & 
    #             (project_agg['Утилизация тайминга, %'] > 80) & 
    #             (project_agg['Утилизация тайминга, %'] < 100)
    #         )
    #         project_agg.loc[mask_focus, 'Фокус'] = 'Да'
        
    #     # Сортируем по План/Факт на дату,%
    #     project_agg = project_agg.sort_values('План/Факт на дату,%', ascending=True)
        
    #     return project_agg
    
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
                # if dsm_selected:
                #     client_filtered = [c for c in client_filtered if c in dsm_selected]
                
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
        
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        with col1:
            show_project = st.checkbox("Код проекта", key='planfact_show_project')
        with col2:
            show_wave = st.checkbox("Волна", key='planfact_show_wave')
        with col3:
            show_regions = st.checkbox("Регионы", key='planfact_show_regions')
        with col4:
            show_dsm = st.checkbox("DSM", key='planfact_show_dsm')
        with col5:
            show_asm = st.checkbox("ASM", key='planfact_show_asm')
        with col6:
            show_rs = st.checkbox("RS", key='planfact_show_rs')
        with col7:
            show_po = st.checkbox("ПО", key='planfact_show_po')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Клиент']
        
        if show_project and 'Проект' in display_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in display_data.columns:
            group_cols.append('Волна')
        if show_regions and region_col in display_data.columns:
            group_cols.append(region_col)
        if show_dsm and 'DSM' in display_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in display_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in display_data.columns:
            group_cols.append('RS')
        if show_po and 'ПО' in display_data.columns:
            group_cols.append('ПО')
        
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
        
        # ВСЕГДА группируем (даже если group_cols == ['Клиент'])
        project_data = display_data.groupby(group_cols).agg(existing_agg).reset_index()
        # Преобразуем коды регионов в длинные названия
        if 'Регион' in project_data.columns:
            project_data['Регион'] = project_data['Регион'].apply(self._get_long_region)
        elif region_col in project_data.columns and region_col != 'Регион':
            project_data[region_col] = project_data[region_col].apply(self._get_long_region)
        
        project_data['Фокус'] = 'Нет'
        if 'План/Факт проекта,%' in project_data.columns and 'Утилизация тайминга, %' in project_data.columns:
            mask_focus = (
                (project_data['План/Факт проекта,%'] < 80) & 
                (project_data['Утилизация тайминга, %'] > 80) & 
                (project_data['Утилизация тайминга, %'] < 100)
            )
            project_data.loc[mask_focus, 'Фокус'] = 'Да'
        
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
        prodata_plan_date_total = 0
        prodata_fact_date_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            if 'План проекта, шт.' in prodata_df.columns:
                prodata_plan_total = prodata_df['План проекта, шт.'].sum()
            if 'Факт проекта, шт.' in prodata_df.columns:
                prodata_fact_total = prodata_df['Факт проекта, шт.'].sum()
            if 'План на дату, шт.' in prodata_df.columns:
                prodata_plan_date_total = prodata_df['План на дату, шт.'].sum()
            if 'Факт на дату, шт.' in prodata_df.columns:
                prodata_fact_date_total = prodata_df['Факт на дату, шт.'].sum()
        
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
        
        # Отображаем таблицу
        display_cols = ['Клиент']
        
        # Добавляем колонки развертки в зависимости от чек-боксов
        if show_project and 'Проект' in project_data.columns:
            display_cols.append('Проект')
        if show_wave and 'Волна' in project_data.columns:
            display_cols.append('Волна')
        if show_regions and region_col in project_data.columns:
            display_cols.append(region_col)
        if show_dsm and 'DSM' in project_data.columns:
            display_cols.append('DSM')
        if show_asm and 'ASM' in project_data.columns:
            display_cols.append('ASM')
        if show_rs and 'RS' in project_data.columns:
            display_cols.append('RS')
        if show_po and 'ПО' in project_data.columns:
            display_cols.append('ПО')
        
        # Колонки, которые показываются всегда
        always_show = [
            'План проекта, шт.', 
            'Факт проекта, шт.', 
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Фокус',
            'Прогноз на месяц, шт.',
            'Длительность',
            'Дней до конца проекта',
            'Утилизация тайминга, %',
            'Ср. план на день для 100% плана'
        ]
        
        # Добавляем всегда показываемые колонки, если они есть в данных
        for col in always_show:
            if col in project_data.columns:
                display_cols.append(col)
        
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
                asm_mode = st.session_state.region_asm_mode
                
                
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
        # 4. РАЗВЕРТКА И ОТОБРАЖЕНИЕ
        # ============================================
        
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        with col1:
            show_client = st.checkbox("Клиент", key='region_show_client')
        with col2:
            show_project = st.checkbox("Код проекта", key='region_show_project')
        with col3:
            show_wave = st.checkbox("Волна", key='region_show_wave')
        with col4:
            show_dsm = st.checkbox("DSM", key='region_show_dsm')
        with col5:
            show_asm = st.checkbox("ASM", key='region_show_asm')
        with col6:
            show_rs = st.checkbox("RS", key='region_show_rs')
        with col7:
            show_po = st.checkbox("ПО", key='region_show_po')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['Регион']
        
        if show_client and 'Клиент' in display_data.columns:
            group_cols.append('Клиент')
        if show_project and 'Проект' in display_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in display_data.columns:
            group_cols.append('Волна')
        if show_dsm and 'DSM' in display_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in display_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in display_data.columns:
            group_cols.append('RS')
        if show_po and 'ПО' in display_data.columns:
            group_cols.append('ПО')
        
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
        
        # ВСЕГДА группируем
        region_data = display_data.groupby(group_cols).agg(existing_agg).reset_index()
        region_data['Фокус'] = 'Нет'
        if 'План/Факт проекта,%' in region_data.columns and 'Утилизация тайминга, %' in region_data.columns:
            mask_focus = (
                (region_data['План/Факт проекта,%'] < 80) & 
                (region_data['Утилизация тайминга, %'] > 80) & 
                (region_data['Утилизация тайминга, %'] < 100)
            )
            region_data.loc[mask_focus, 'Фокус'] = 'Да'
        
        # Преобразуем коды регионов в длинные названия
        if 'Регион' in region_data.columns:
            region_data['Регион'] = region_data['Регион'].apply(self._get_long_region)
    
        # Добавляем вычисляемые метрики
        mask_plan = region_data['План на дату, шт.'] > 0
        region_data['План/Факт на дату,%'] = 0.0
        if mask_plan.any():
            region_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                region_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                region_data.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        mask_project_plan = region_data['План проекта, шт.'] > 0
        region_data['План/Факт проекта,%'] = 0.0
        if mask_project_plan.any():
            region_data.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                region_data.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                region_data.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        region_data['△План/Факт на дату, шт'] = (
            region_data['Факт на дату, шт.'] - region_data['План на дату, шт.']
        ).round(1)
        
        region_data['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            region_data.loc[mask_plan, '△План/Факт на дату, %'] = (
                (region_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                 region_data.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        col_kpi1, col_kpi2, col_kpi3, col_checkbox = st.columns([1, 1, 1, 0.5])
        with col_checkbox:
            include_prodata = st.checkbox("📊 Продата", key="region_include_prodata")
        
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        
        prodata_plan_total = 0
        prodata_fact_total = 0
        prodata_plan_date_total = 0
        prodata_fact_date_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            if 'План проекта, шт.' in prodata_df.columns:
                prodata_plan_total = prodata_df['План проекта, шт.'].sum()
            if 'Факт проекта, шт.' in prodata_df.columns:
                prodata_fact_total = prodata_df['Факт проекта, шт.'].sum()
            if 'План на дату, шт.' in prodata_df.columns:
                prodata_plan_date_total = prodata_df['План на дату, шт.'].sum()
            if 'Факт на дату, шт.' in prodata_df.columns:
                prodata_fact_date_total = prodata_df['Факт на дату, шт.'].sum()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = region_data['План проекта, шт.'].sum() if 'План проекта, шт.' in region_data.columns else 0
            if include_prodata:
                plan_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = region_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in region_data.columns else 0
            if include_prodata:
                fact_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_percent:.1f}%")

        # Второй ряд: План на дату, Факт на дату, План/Факт на дату
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = region_data['План на дату, шт.'].sum() if 'План на дату, шт.' in region_data.columns else 0
            if include_prodata:
                plan_date_total += prodata_plan_date_total
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        
        with col5:
            fact_date_total = region_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in region_data.columns else 0
            if include_prodata:
                fact_date_total += prodata_fact_date_total
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
        
        # Отображаем таблицу
        display_cols = ['Регион']
        
        if show_client and 'Клиент' in region_data.columns:
            display_cols.append('Клиент')
        if show_project and 'Проект' in region_data.columns:
            display_cols.append('Проект')
        if show_wave and 'Волна' in region_data.columns:
            display_cols.append('Волна')
        if show_dsm and 'DSM' in region_data.columns:
            display_cols.append('DSM')
        if show_asm and 'ASM' in region_data.columns:
            display_cols.append('ASM')
        if show_rs and 'RS' in region_data.columns:
            display_cols.append('RS')
        if show_po and 'ПО' in region_data.columns:
            display_cols.append('ПО')
        
        # Колонки, которые показываются всегда
        always_show = [
            'План проекта, шт.', 
            'Факт проекта, шт.', 
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Фокус',
            'Прогноз на месяц, шт.',
            'Длительность',
            'Дней до конца проекта',
            'Утилизация тайминга, %',
            'Ср. план на день для 100% плана'
        ]
        
        for col in always_show:
            if col in region_data.columns:
                display_cols.append(col)
        
        existing_display = [c for c in display_cols if c in region_data.columns]
        
        # Преобразуем коды регионов в названия
        if 'Регион' in region_data.columns:
            region_data['Регион'] = region_data['Регион'].apply(self._get_long_region)
        
        st.dataframe(region_data[existing_display], use_container_width=True, hide_index=True)
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            region_data.to_excel(writer, sheet_name='Регионы', index=False)
        
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
        all_dsm = base_data['all_dsm']
        
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
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_dsm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_dsm_mode', 'Включить') == 'Включить' else 1
                )
                dsm_mode = st.session_state.dsm_dsm_mode
                
                dsm_selected = st.multiselect(
                    "Выбрать DSM",
                    all_dsm,
                    key="dsm_dsm_values",
                    default=[]
                )
            
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_asm_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_asm_mode', 'Включить') == 'Включить' else 1
                )
                asm_mode = st.session_state.dsm_asm_mode

                asm_options = all_asm
                asm_selected = st.multiselect(
                    "Выбрать ASM",
                    asm_options,
                    key="dsm_asm_values",
                    default=[]
                )
            
            with col3:
                st.markdown("**Клиент**")
                client_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_client_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_client_mode', 'Включить') == 'Включить' else 1
                )
                client_mode = st.session_state.dsm_client_mode
                
                client_filtered = all_clients
                if dsm_selected:
                    client_filtered = [c for c in client_filtered if c in dsm_selected]
                
                client_selected = st.multiselect(
                    "Выбрать клиента",
                    client_filtered,
                    key="dsm_client_values",
                    default=[]
                )
            
            with col4:
                st.markdown("**Регион**")
                region_mode = st.radio(
                    "Режим",
                    ["Включить", "Исключить"],
                    key="dsm_region_mode",
                    horizontal=True,
                    index=0 if st.session_state.get('dsm_region_mode', 'Включить') == 'Включить' else 1
                )
                region_mode = st.session_state.dsm_region_mode
                
                region_selected_display = st.multiselect(
                    "Выбрать регион",
                    all_regions_display,
                    key="dsm_region_values",
                    default=[]
                )
                region_selected = [region_map.get(name, name) for name in region_selected_display]
                
                region_mode = st.session_state.dsm_region_mode
            
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
                dsm_selected, st.session_state.dsm_dsm_mode,
                asm_selected, st.session_state.dsm_asm_mode,
                client_selected, st.session_state.dsm_client_mode,
                region_selected, st.session_state.dsm_region_mode,
                region_col
            )
            st.session_state.dsm_filtered_data = filtered_data
        
        if st.session_state.dsm_filtered_data is not None:
            display_data = st.session_state.dsm_filtered_data
        else:
            display_data = base_data['raw_data']
        
        # ============================================
        # 4. РАЗВЕРТКА И ОТОБРАЖЕНИЕ
        # ============================================
        
        st.subheader("📊 Детализация")
        
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        with col1:
            show_client = st.checkbox("Клиент", key='dsm_show_client')
        with col2:
            show_project = st.checkbox("Код проекта", key='dsm_show_project')
        with col3:
            show_wave = st.checkbox("Волна", key='dsm_show_wave')
        with col4:
            show_region = st.checkbox("Регионы", key='dsm_show_region')
        with col5:
            show_asm = st.checkbox("ASM", key='dsm_show_asm')
        with col6:
            show_rs = st.checkbox("RS", key='dsm_show_rs')
        with col7:
            show_po = st.checkbox("ПО", key='dsm_show_po')
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['DSM']
        
        if show_client and 'Клиент' in display_data.columns:
            group_cols.append('Клиент')
        if show_project and 'Проект' in display_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in display_data.columns:
            group_cols.append('Волна')
        if show_region and region_col in display_data.columns:
            group_cols.append(region_col)
        if show_asm and 'ASM' in display_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in display_data.columns:
            group_cols.append('RS')
        if show_po and 'ПО' in display_data.columns:
            group_cols.append('ПО')
        
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
        
        # ВСЕГДА группируем
        dsm_data = display_data.groupby(group_cols).agg(existing_agg).reset_index()

        dsm_data['Фокус'] = 'Нет'
        if 'План/Факт проекта,%' in dsm_data.columns and 'Утилизация тайминга, %' in dsm_data.columns:
            mask_focus = (
                (dsm_data['План/Факт проекта,%'] < 80) & 
                (dsm_data['Утилизация тайминга, %'] > 80) & 
                (dsm_data['Утилизация тайминга, %'] < 100)
            )
            dsm_data.loc[mask_focus, 'Фокус'] = 'Да'
        
        # Добавляем вычисляемые метрики
        mask_plan = dsm_data['План на дату, шт.'] > 0

        # Преобразуем коды регионов в длинные названия
        if 'Регион' in dsm_data.columns:
            dsm_data['Регион'] = dsm_data['Регион'].apply(self._get_long_region)
        
        # Добавляем вычисляемые метрики
        mask_plan = dsm_data['План на дату, шт.'] > 0
        dsm_data['План/Факт на дату,%'] = 0.0
        if mask_plan.any():
            dsm_data.loc[mask_plan, 'План/Факт на дату,%'] = (
                dsm_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                dsm_data.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        mask_project_plan = dsm_data['План проекта, шт.'] > 0
        dsm_data['План/Факт проекта,%'] = 0.0
        if mask_project_plan.any():
            dsm_data.loc[mask_project_plan, 'План/Факт проекта,%'] = (
                dsm_data.loc[mask_project_plan, 'Факт проекта, шт.'] / 
                dsm_data.loc[mask_project_plan, 'План проекта, шт.'] * 100
            ).round(1)
        
        dsm_data['△План/Факт на дату, шт'] = (
            dsm_data['Факт на дату, шт.'] - dsm_data['План на дату, шт.']
        ).round(1)
        
        dsm_data['△План/Факт на дату, %'] = 0.0
        if mask_plan.any():
            dsm_data.loc[mask_plan, '△План/Факт на дату, %'] = (
                (dsm_data.loc[mask_plan, 'Факт на дату, шт.'] / 
                 dsm_data.loc[mask_plan, 'План на дату, шт.']) - 1
            ).round(3) * 100
        
        # KPI
        st.markdown("### 📊 Ключевые показатели")
        
        col_kpi1, col_kpi2, col_kpi3, col_checkbox = st.columns([1, 1, 1, 0.5])
        with col_checkbox:
            include_prodata = st.checkbox("📊 Продата", key="dsm_include_prodata")
        
        prodata_df = st.session_state.cleaned_data.get('prodata_processed', None)
        
        prodata_plan_total = 0
        prodata_fact_total = 0
        prodata_plan_date_total = 0
        prodata_fact_date_total = 0
        
        if include_prodata and prodata_df is not None and not prodata_df.empty:
            if 'План проекта, шт.' in prodata_df.columns:
                prodata_plan_total = prodata_df['План проекта, шт.'].sum()
            if 'Факт проекта, шт.' in prodata_df.columns:
                prodata_fact_total = prodata_df['Факт проекта, шт.'].sum()
            if 'План на дату, шт.' in prodata_df.columns:
                prodata_plan_date_total = prodata_df['План на дату, шт.'].sum()
            if 'Факт на дату, шт.' in prodata_df.columns:
                prodata_fact_date_total = prodata_df['Факт на дату, шт.'].sum()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = dsm_data['План проекта, шт.'].sum() if 'План проекта, шт.' in dsm_data.columns else 0
            if include_prodata:
                plan_total += prodata_plan_total
            st.metric("📊 План проекта", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = dsm_data['Факт проекта, шт.'].sum() if 'Факт проекта, шт.' in dsm_data.columns else 0
            if include_prodata:
                fact_total += prodata_fact_total
            st.metric("✅ Факт проекта", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт проекта", f"{pf_percent:.1f}%")

        # Второй ряд: План на дату, Факт на дату, План/Факт на дату
        col4, col5, col6 = st.columns(3)
        with col4:
            plan_date_total = dsm_data['План на дату, шт.'].sum() if 'План на дату, шт.' in dsm_data.columns else 0
            if include_prodata:
                plan_date_total += prodata_plan_date_total
            st.metric("📊 План на дату", f"{plan_date_total:,.0f} шт")
        
        with col5:
            fact_date_total = dsm_data['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in dsm_data.columns else 0
            if include_prodata:
                fact_date_total += prodata_fact_date_total
            st.metric("✅ Факт на дату", f"{fact_date_total:,.0f} шт")
        
        with col6:
            pf_date_percent = (fact_date_total / plan_date_total * 100) if plan_date_total > 0 else 0
            st.metric("🎯 План/Факт на дату", f"{pf_date_percent:.1f}%")
        
        # Отображаем таблицу
        display_cols = ['DSM']
        
        if show_client and 'Клиент' in dsm_data.columns:
            display_cols.append('Клиент')
        if show_project and 'Проект' in dsm_data.columns:
            display_cols.append('Проект')
        if show_wave and 'Волна' in dsm_data.columns:
            display_cols.append('Волна')
        if show_region and region_col in dsm_data.columns:
            display_cols.append(region_col)
        if show_asm and 'ASM' in dsm_data.columns:
            display_cols.append('ASM')
        if show_rs and 'RS' in dsm_data.columns:
            display_cols.append('RS')
        if show_po and 'ПО' in dsm_data.columns:
            display_cols.append('ПО')
        
        # Колонки, которые показываются всегда
        always_show = [
            'План проекта, шт.', 
            'Факт проекта, шт.', 
            'План/Факт проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            'План/Факт на дату,%',
            '△План/Факт на дату, шт',
            '△План/Факт на дату, %',
            'Фокус',
            'Прогноз на месяц, шт.',
            'Длительность',
            'Дней до конца проекта',
            'Утилизация тайминга, %',
            'Ср. план на день для 100% плана'
        ]
        
        for col in always_show:
            if col in dsm_data.columns:
                display_cols.append(col)
        
        existing_display = [c for c in display_cols if c in dsm_data.columns]
        
        st.dataframe(dsm_data[existing_display], use_container_width=True, hide_index=True)
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dsm_data.to_excel(writer, sheet_name='DSM', index=False)
        
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

    def create_dynamics_tab(self, data, visits_df, calc_params):
        """
        Создает вкладку Динамика с фактом визитов по дням
        """
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        if visits_df is None or visits_df.empty:
            st.warning("⚠️ Нет данных визитов для динамики")
            return
        
        st.subheader("📈 Динамика факта визитов")
        
        # Переименовываем колонки
        rename_cols = {'ЗОД': 'DSM', 'АСС': 'ASM', 'ЭМ': 'RS'}
        data = data.rename(columns=rename_cols)
        
        if 'RS' not in visits_df.columns and 'ЭМ' in visits_df.columns:
            visits_df = visits_df.rename(columns={'ЭМ': 'RS'})
        
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        visits_region_col = region_col
        if visits_region_col not in visits_df.columns and 'Регион short' in visits_df.columns:
            visits_region_col = 'Регион short'
        
        # Базовые агрегации
        all_dsm = sorted(data['DSM'].dropna().unique()) if 'DSM' in data.columns else []
        all_asm = sorted(data['ASM'].dropna().unique()) if 'ASM' in data.columns else []
        all_clients = sorted(data['Клиент'].dropna().unique()) if 'Клиент' in data.columns else []
        
        region_map = {}
        regions_display = []
        if region_col in data.columns:
            unique_codes = data[region_col].dropna().unique()
            for code in unique_codes:
                long_name = self._get_long_region(code)
                region_map[long_name] = code
                regions_display.append(long_name)
            regions_display.sort()
        
        # Форма с фильтрами
        with st.form("dynamics_filters_form"):
            st.markdown("### 🔍 Фильтры")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**DSM**")
                dsm_mode = st.radio("Режим", ["Включить", "Исключить"], key="dynamics_dsm_mode", horizontal=True, index=0)
                dsm_selected = st.multiselect("Выбрать DSM", all_dsm, key="dynamics_dsm_values", default=[])
            
            with col2:
                st.markdown("**ASM**")
                asm_mode = st.radio("Режим", ["Включить", "Исключить"], key="dynamics_asm_mode", horizontal=True, index=0)
                asm_selected = st.multiselect("Выбрать ASM", all_asm, key="dynamics_asm_values", default=[])
            
            with col3:
                st.markdown("**Регион**")
                region_mode = st.radio("Режим", ["Включить", "Исключить"], key="dynamics_region_mode", horizontal=True, index=0)
                region_selected_display = st.multiselect("Выбрать регион", regions_display, key="dynamics_region_values", default=[])
                region_selected = [region_map.get(name, name) for name in region_selected_display]
            
            with col4:
                st.markdown("**Клиент**")
                client_mode = st.radio("Режим", ["Включить", "Исключить"], key="dynamics_client_mode", horizontal=True, index=0)
                client_selected = st.multiselect("Выбрать клиента", all_clients, key="dynamics_client_values", default=[])
            
            st.markdown("---")
            st.markdown("### 📊 Детализация")
            
            col_detail1, col_detail2, col_detail3, col_detail4, col_detail5, col_detail6 = st.columns(6)
            
            with col_detail1:
                show_project = st.checkbox("Код проекта", key="dynamics_show_project")
            with col_detail2:
                show_wave = st.checkbox("Волна", key="dynamics_show_wave")
            with col_detail3:
                show_region_detail = st.checkbox("Регионы", key="dynamics_show_region")
            with col_detail4:
                show_dsm = st.checkbox("DSM", key="dynamics_show_dsm")
            with col_detail5:
                show_asm = st.checkbox("ASM", key="dynamics_show_asm")
            with col_detail6:
                show_rs = st.checkbox("RS", key="dynamics_show_rs")
            
            apply_filters = st.form_submit_button("✅ Применить фильтры", type="primary", use_container_width=True)
        
        # Применяем фильтры
        filtered_data = data.copy()
        
        if dsm_selected:
            if dsm_mode == 'Включить':
                filtered_data = filtered_data[filtered_data['DSM'].isin(dsm_selected)]
            else:
                filtered_data = filtered_data[~filtered_data['DSM'].isin(dsm_selected)]
        
        if asm_selected:
            if asm_mode == 'Включить':
                filtered_data = filtered_data[filtered_data['ASM'].isin(asm_selected)]
            else:
                filtered_data = filtered_data[~filtered_data['ASM'].isin(asm_selected)]
        
        if region_selected:
            if region_mode == 'Включить':
                filtered_data = filtered_data[filtered_data[region_col].isin(region_selected)]
            else:
                filtered_data = filtered_data[~filtered_data[region_col].isin(region_selected)]
        
        if client_selected:
            if client_mode == 'Включить':
                filtered_data = filtered_data[filtered_data['Клиент'].isin(client_selected)]
            else:
                filtered_data = filtered_data[~filtered_data['Клиент'].isin(client_selected)]
        
        # Формируем группы для детализации
        group_cols = ['Клиент']
        
        if show_project and 'Проект' in filtered_data.columns:
            group_cols.append('Проект')
        if show_wave and 'Волна' in filtered_data.columns:
            group_cols.append('Волна')
        if show_region_detail and region_col in filtered_data.columns:
            group_cols.append(region_col)
        if show_dsm and 'DSM' in filtered_data.columns:
            group_cols.append('DSM')
        if show_asm and 'ASM' in filtered_data.columns:
            group_cols.append('ASM')
        if show_rs and 'RS' in filtered_data.columns:
            group_cols.append('RS')
        
        # Фильтруем визиты (оптимизированно)
        filter_keys = ['Клиент', 'Проект', 'Волна', region_col, 'DSM', 'ASM', 'RS']
        
        # 🔧 ПЕРЕИМЕНОВЫВАЕМ колонку 'Имя клиента' в 'Клиент' для единообразия
        visits_df_copy = visits_df.copy()
        if 'Имя клиента' in visits_df_copy.columns and 'Клиент' not in visits_df_copy.columns:
            visits_df_copy = visits_df_copy.rename(columns={'Имя клиента': 'Клиент'})
        
        existing_filter_keys = [k for k in filter_keys if k in filtered_data.columns and k in visits_df_copy.columns]
        
        if not existing_filter_keys:
            st.warning("⚠️ Нет общих полей для фильтрации визитов")
            return
        
        visits_df_copy['_filter_key'] = visits_df_copy[existing_filter_keys].astype(str).agg('|'.join, axis=1)
        
        valid_keys = set()
        for _, row in filtered_data[existing_filter_keys].drop_duplicates().iterrows():
            key = '|'.join(str(row[k]) for k in existing_filter_keys)
            valid_keys.add(key)
        
        visits_filtered = visits_df_copy[visits_df_copy['_filter_key'].isin(valid_keys)].copy()
        visits_filtered = visits_filtered.drop('_filter_key', axis=1)
        
        if visits_filtered.empty:
            st.warning("⚠️ Нет визитов, соответствующих выбранным фильтрам")
            return
        
        # Расчет динамики
        dynamics_df = visit_calculator.calculate_dynamics_fact(visits_filtered, calc_params, group_cols)
        
        if dynamics_df.empty:
            st.warning("⚠️ Нет данных для отображения динамики")
            return
        
        # Сводная таблица
        pivot_df = dynamics_df.pivot_table(
            index=group_cols,
            columns='Дата',
            values='Факт',
            fill_value=0,
            aggfunc='sum'
        )
        
        pivot_df.columns = [col.strftime('%d.%b') if hasattr(col, 'strftime') else str(col) for col in pivot_df.columns]
        
        try:
            pivot_df = pivot_df.reindex(sorted(pivot_df.columns, key=lambda x: x.split('.')[0].zfill(2) if '.' in x else x), axis=1)
        except:
            pass
        
        result_df = pivot_df.reset_index()
        
        if show_region_detail and region_col in result_df.columns:
            result_df[region_col] = result_df[region_col].apply(self._get_long_region)
        
        # Отображение
        total_fact = dynamics_df['Факт'].sum()
        st.metric("📊 Всего визитов за период", f"{total_fact:,.0f}")
        st.markdown("---")
        
        st.dataframe(result_df, use_container_width=True, hide_index=True)
        
        # Скачивание
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name='Динамика_факта', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"динамика_факта_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )


# Глобальный экземпляр
dataviz = DataVisualizer()
