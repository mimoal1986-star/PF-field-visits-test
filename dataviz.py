# utils/dataviz.py
# draft 3.1 
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
            'ЭМ': 'RS'
        }
        data = data.rename(columns=rename_cols)
        
        # 🔍 ФИЛЬТРЫ (каскадные)
        with st.expander("🔍 Фильтры", expanded=True):
            # Определяем, какую колонку региона использовать
            region_col = 'Регион'
            if 'Регион short' in data.columns and 'Регион' not in data.columns:
                region_col = 'Регион short'
            
            # 👇 ИЗМЕНЕНО: для фильтра регионов используем длинные названия
            # Получаем уникальные значения для фильтров
            all_dsm = data['DSM'].dropna().unique() if 'DSM' in data.columns else []
            all_asm = data['ASM'].dropna().unique() if 'ASM' in data.columns else []
            
            # Создаем список длинных названий регионов для фильтра
            if region_col in data.columns:
                # Получаем уникальные коды регионов
                unique_codes = data[region_col].dropna().unique()
                # Создаем словарь для маппинга длинных названий обратно в коды
                self.region_display_map = {}
                all_regions_display = []
                for code in unique_codes:
                    long_name = self._get_long_region(code)
                    self.region_display_map[long_name] = code
                    all_regions_display.append(long_name)
                # Сортируем по алфавиту
                all_regions_display.sort()
            else:
                all_regions_display = []
                self.region_display_map = {}
            
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
                # 👇 ИЗМЕНЕНО: показываем длинные названия в фильтре
                selected_region_display = st.multiselect('Регион', all_regions_display, key='filter_region')
                # Преобразуем обратно в коды для фильтрации
                selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
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
        
        # 👇 ИЗМЕНЕНО: применяем длинные названия к колонке региона в таблице
        if region_col in df_display.columns:
            df_display[region_col] = df_display[region_col].apply(self._get_long_region)
        
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
        
        # 🔍 ФИЛЬТРЫ (каскадные)
        with st.expander("🔍 Фильтры", expanded=True):
            # Получаем уникальные значения для фильтров
            all_dsm = data['DSM'].dropna().unique() if 'DSM' in data.columns else []
            all_asm = data['ASM'].dropna().unique() if 'ASM' in data.columns else []
            
            # 👇 ИЗМЕНЕНО: для фильтра регионов используем длинные названия
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
            
            all_clients = data['Клиент'].dropna().unique() if 'Клиент' in data.columns else []
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                selected_dsm = st.multiselect('DSM', all_dsm, key='region_filter_dsm')
            with col2:
                asm_options = all_asm
                if selected_dsm and 'DSM' in data.columns and 'ASM' in data.columns:
                    asm_options = data[data['DSM'].isin(selected_dsm)]['ASM'].dropna().unique()
                selected_asm = st.multiselect('ASM', asm_options, key='region_filter_asm')
            with col3:
                client_options = all_clients
                filtered_for_client = data.copy()
                if selected_dsm and 'DSM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['DSM'].isin(selected_dsm)]
                if selected_asm and 'ASM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['ASM'].isin(selected_asm)]
                if 'Клиент' in filtered_for_client.columns:
                    client_options = filtered_for_client['Клиент'].dropna().unique()
                selected_client = st.multiselect('Клиент', client_options, key='region_filter_client')
            with col4:
                # 👇 ИЗМЕНЕНО: показываем длинные названия в фильтре регионов
                selected_region_display = st.multiselect('Регион', all_regions_display, key='region_filter_region')
                selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
        
        # Применяем фильтры к данным
        filtered_data = data.copy()
        
        if selected_dsm and 'DSM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['DSM'].isin(selected_dsm)]
        if selected_asm and 'ASM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if selected_client and 'Клиент' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        if selected_region and region_col in filtered_data.columns:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        
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
        group_cols = [region_col]  # Только регион
        
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
        if len(group_cols) > 1:  # больше чем просто Регион
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
            
            # Пересчитываем метрики
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
        display_columns = [region_col]  # Только регион
        
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
        
        # Добавляем метрики
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
        
        # Добавляем дополнительные колонки из региональной сводки
        extra_columns = ['Кол-во клиентов', 'Кол-во проектов', 'Кол-во сотрудников', 'ПО']
        for col in extra_columns:
            if col in region_data.columns:
                metric_columns.append(col)
        
        display_columns.extend([col for col in metric_columns if col in region_data.columns])
        
        # Только существующие колонки
        existing_cols = [col for col in display_columns if col in region_data.columns]
        df_display = region_data[existing_cols].copy()
        
        # 👇 ИЗМЕНЕНО: применяем длинные названия к колонке региона
        if region_col in df_display.columns:
            df_display[region_col] = df_display[region_col].apply(self._get_long_region)
        
        # Форматирование
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        
        if 'План/Факт проекта,%' in df_display.columns:
            df_display['План/Факт проекта,%'] = df_display['План/Факт проекта,%'].map(lambda x: f"{x:.1f}%")
        
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        # Таблица
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )
        
        # Кнопка скачивания
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
        
        # Определяем колонку региона для полных названий
        region_col = 'Регион'
        if 'Регион short' in data.columns and 'Регион' not in data.columns:
            region_col = 'Регион short'
        
        # 🔍 ФИЛЬТРЫ (каскадные)
        with st.expander("🔍 Фильтры", expanded=True):
            # Получаем уникальные значения для фильтров
            all_asm = data['ASM'].dropna().unique() if 'ASM' in data.columns else []
            all_clients = data['Клиент'].dropna().unique() if 'Клиент' in data.columns else []
            all_projects = data['Проект'].dropna().unique() if 'Проект' in data.columns else []
            all_waves = data['Волна'].dropna().unique() if 'Волна' in data.columns else []
            
            # 👇 ИЗМЕНЕНО: для фильтра регионов используем длинные названия
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
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                selected_asm = st.multiselect('ASM', all_asm, key='dsm_filter_asm')
            with col2:
                client_options = all_clients
                filtered_for_client = data.copy()
                if selected_asm and 'ASM' in filtered_for_client.columns:
                    filtered_for_client = filtered_for_client[filtered_for_client['ASM'].isin(selected_asm)]
                if 'Клиент' in filtered_for_client.columns:
                    client_options = filtered_for_client['Клиент'].dropna().unique()
                selected_client = st.multiselect('Клиент', client_options, key='dsm_filter_client')
            with col3:
                project_options = all_projects
                filtered_for_project = data.copy()
                if selected_asm and 'ASM' in filtered_for_project.columns:
                    filtered_for_project = filtered_for_project[filtered_for_project['ASM'].isin(selected_asm)]
                if selected_client and 'Клиент' in filtered_for_project.columns:
                    filtered_for_project = filtered_for_project[filtered_for_project['Клиент'].isin(selected_client)]
                if 'Проект' in filtered_for_project.columns:
                    project_options = filtered_for_project['Проект'].dropna().unique()
                selected_project = st.multiselect('Код проекта', project_options, key='dsm_filter_project')
            with col4:
                wave_options = all_waves
                filtered_for_wave = data.copy()
                if selected_asm and 'ASM' in filtered_for_wave.columns:
                    filtered_for_wave = filtered_for_wave[filtered_for_wave['ASM'].isin(selected_asm)]
                if selected_client and 'Клиент' in filtered_for_wave.columns:
                    filtered_for_wave = filtered_for_wave[filtered_for_wave['Клиент'].isin(selected_client)]
                if selected_project and 'Проект' in filtered_for_wave.columns:
                    filtered_for_wave = filtered_for_wave[filtered_for_wave['Проект'].isin(selected_project)]
                if 'Волна' in filtered_for_wave.columns:
                    wave_options = filtered_for_wave['Волна'].dropna().unique()
                selected_wave = st.multiselect('Волна', wave_options, key='dsm_filter_wave')
            with col5:
                # 👇 ИЗМЕНЕНО: показываем длинные названия в фильтре регионов
                selected_region_display = st.multiselect('Регион', all_regions_display, key='dsm_filter_region')
                selected_region = [self.region_display_map.get(name, name) for name in selected_region_display]
        
        # Применяем фильтры к данным
        filtered_data = data.copy()
        
        if selected_asm and 'ASM' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['ASM'].isin(selected_asm)]
        if selected_client and 'Клиент' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['Клиент'].isin(selected_client)]
        if selected_project and 'Проект' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['Проект'].isin(selected_project)]
        if selected_wave and 'Волна' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['Волна'].isin(selected_wave)]
        if selected_region and region_col in filtered_data.columns:
            filtered_data = filtered_data[filtered_data[region_col].isin(selected_region)]
        
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
        
        # Формируем groupby в зависимости от чек-боксов
        group_cols = ['DSM']  # Только DSM
        
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
        
        # Агрегируем данные с учетом развертки
        if len(group_cols) > 1:  # больше чем просто DSM
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
            
            # Пересчитываем метрики
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
        
        # Добавляем метрики
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
        
        # Добавляем дополнительные колонки из сводки по DSM
        extra_columns = ['Кол-во проектов', 'Кол-во сотрудников', 'ПО']
        for col in extra_columns:
            if col in dsm_data.columns:
                metric_columns.append(col)
        
        display_columns.extend([col for col in metric_columns if col in dsm_data.columns])
        
        # Только существующие колонки
        existing_cols = [col for col in display_columns if col in dsm_data.columns]
        df_display = dsm_data[existing_cols].copy()
        
        # 👇 ИЗМЕНЕНО: применяем длинные названия к колонке региона если она есть
        if show_region and region_col in df_display.columns:
            df_display[region_col] = df_display[region_col].apply(self._get_long_region)
        
        # Форматирование
        if 'План/Факт на дату,%' in df_display.columns:
            df_display['План/Факт на дату,%'] = df_display['План/Факт на дату,%'].map(lambda x: f"{x:.1f}%")
        
        if 'План/Факт проекта,%' in df_display.columns:
            df_display['План/Факт проекта,%'] = df_display['План/Факт проекта,%'].map(lambda x: f"{x:.1f}%")
        
        if '△План/Факт на дату, %' in df_display.columns:
            df_display['△План/Факт на дату, %'] = df_display['△План/Факт на дату, %'].map(lambda x: f"{x:+.1f}%")
        
        # Таблица
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )
        
        # Кнопка скачивания
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

# Глобальный экземпляр
dataviz = DataVisualizer()
