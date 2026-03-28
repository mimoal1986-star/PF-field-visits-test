# app.py
# draft 4.1 - simplified
import streamlit as st
import pandas as pd
import sys
import os
import traceback
from datetime import date, datetime, timedelta
from io import BytesIO
from github_settings import get_settings_manager
from data_cleaner import is_field_project_by_code

# data_cleaner.py
try:
    from utils.data_cleaner import data_cleaner
except ImportError:
    from data_cleaner import DataCleaner
    data_cleaner = DataCleaner()
    
# visit_calculator.py
from visit_calculator import VisitCalculator
visit_calculator = VisitCalculator()

# dataviz.py
try:
    from utils.dataviz import dataviz
except ImportError:
    from dataviz import DataVisualizer
    dataviz = DataVisualizer()

# Настройка путей
current_dir = os.path.dirname(os.path.abspath(__file__))
utils_path = os.path.join(current_dir, 'utils')
if utils_path not in sys.path:
    sys.path.append(utils_path)

# Настройка страницы
st.set_page_config(
    page_title="ИУ Аудиты - Аналитика",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Инициализация session_state
DEFAULT_STATE = {
    'uploaded_files': {},
    'cleaned_data': {},
    'processing_complete': False,
    'last_error': None,
    'visit_report': {},
    'plan_calc_params': None
}

for key, default_value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_value

# Вспомогательные функции
def validate_file_upload(file_obj, file_name):
    """Проверка и загрузка файла"""
    if file_obj is None:
        return None
    
    try:
        df = pd.read_excel(file_obj, dtype=str)
        if df.empty:
            return None
        return df
    except Exception as e:
        return None

def display_file_preview(df, title):
    """Отображение предпросмотра файла"""
    if df is not None and not df.empty:
        with st.expander(f"👀 {title}"):
            st.dataframe(df.head(10), use_container_width=True)

def process_all_data(settings_manager=None):
    """Полная обработка данных и расчет план/факт - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ"""
    try:
        # Проверяем наличие основных файлов
        required_files = ['портал', 'сервизория']
        missing_files = [f for f in required_files if f not in st.session_state.uploaded_files]
        
        if missing_files:
            return False
        
        # ШАГ 1: ОЧИСТКА ИСТОЧНИКОВ (БЕЗ ОПРЕДЕЛЕНИЯ ПОЛЕВОЙ)
        
        # Получаем данные
        portal_raw = st.session_state.uploaded_files['портал']
        google_raw = st.session_state.uploaded_files['сервизория']
        
        # Очистка портала
        portal_cleaned = data_cleaner.clean_array(portal_raw)
        if portal_cleaned is None:
            portal_cleaned = portal_raw
        st.session_state.cleaned_data['портал'] = portal_cleaned
        
        # Очистка проектов
        google_cleaned = data_cleaner.clean_google(google_raw)
        if google_cleaned is None:
            google_cleaned = google_raw
        st.session_state.cleaned_data['сервизория'] = google_cleaned
        
        # Обогащение массива кодами проектов
        enriched_result = data_cleaner.enrich_array_with_project_codes(
            st.session_state.cleaned_data['портал'],
            st.session_state.cleaned_data['сервизория']
        )
        
        if enriched_result:
            enriched_array, discrepancy_df, stats = enriched_result
            st.session_state.cleaned_data['портал'] = enriched_array
        
        # Добавление признака полевой проект в Google
        google_with_field = data_cleaner.update_field_projects_flag(st.session_state.cleaned_data['сервизория'])
        st.session_state.cleaned_data['сервизория'] = google_with_field
        
        # Добавляем поле в портал
        array_with_field = data_cleaner.add_field_flag_to_array(st.session_state.cleaned_data['портал'])
        array_with_portal = data_cleaner.add_portal_to_array(array_with_field, google_with_field)
        
        # Удаляем проекты CXWAY из портала
        array_with_portal = data_cleaner.remove_cxway_from_portal(array_with_portal, google_with_field)
        

        # ШАГ 2: СБОР ВСЕХ ИСТОЧНИКОВ В ОДИН ДАТАСЕТ
        
        all_sources = []
        
        # 1. Портал (уже очищен)
        if array_with_portal is not None and not array_with_portal.empty:
            all_sources.append(array_with_portal)
        
        # 2. CXWAY
        cxway_raw = st.session_state.uploaded_files.get('cxway')
        if cxway_raw is not None:
            cxway_processed = data_cleaner.clean_cxway(cxway_raw, None, google_with_field)
            if cxway_processed is not None and not cxway_processed.empty:
                all_sources.append(cxway_processed)
                st.session_state.cleaned_data['cxway_processed'] = cxway_processed
        
        # 3. Easymerch
        easymerch_raw = st.session_state.uploaded_files.get('easymerch')
        if easymerch_raw is not None:
            easymerch_processed = data_cleaner.clean_easymerch(easymerch_raw, google_with_field)
            if easymerch_processed is not None and not easymerch_processed.empty:
                all_sources.append(easymerch_processed)
                st.session_state.cleaned_data['easymerch_processed'] = easymerch_processed
        
        # 4. Optima
        optima_raw = st.session_state.uploaded_files.get('optima')
        if optima_raw is not None:
            try:
                optima_processed = data_cleaner.clean_optima(optima_raw, google_with_field)
                if optima_processed is not None and not optima_processed.empty:
                    all_sources.append(optima_processed)
                    st.session_state.cleaned_data['optima_processed'] = optima_processed
            except Exception as e:
                st.warning(f"⚠️ Ошибка при обработке Optima: {e}")
        
        # 5. ProData (Мониторинги)
        prodata_raw = st.session_state.uploaded_files.get('prodata')
        if prodata_raw is not None:
            try:
                prodata_processed = data_cleaner.clean_prodata(prodata_raw, google_with_field)
                if prodata_processed is not None and not prodata_processed.empty:
                    all_sources.append(prodata_processed)
                    st.session_state.cleaned_data['prodata_processed'] = prodata_processed
            except Exception as e:
                st.warning(f"⚠️ Ошибка при обработке ПроДата: {e}")
        

        # ШАГ 3: ОБЪЕДИНЕНИЕ ВСЕХ ИСТОЧНИКОВ
        
        if not all_sources:
            st.error("❌ Нет данных для обработки")
            return False
        
        all_projects_raw = pd.concat(all_sources, ignore_index=True)
        

        # ШАГ 4: ЕДИНОЕ ОПРЕДЕЛЕНИЕ ПОЛЕВОГО ПРОЕКТА
        
        if 'Код анкеты' in all_projects_raw.columns:
            all_projects_raw['Полевой'] = all_projects_raw['Код анкеты'].apply(is_field_project_by_code)
        else:
            all_projects_raw['Полевой'] = 0
        

        # ШАГ 5: ПРИМЕНЕНИЕ НАСТРОЕК (ИСКЛЮЧЕНИЯ И ДОБАВЛЕНИЯ)
        
        if settings_manager is None:
            settings_manager = get_settings_manager()
        
        excluded_df = settings_manager.get_excluded_projects()
        included_df = settings_manager.get_included_projects()
        
        # 5.1 Применяем исключенные проекты
        if not excluded_df.empty:
            for _, row in excluded_df.iterrows():
                mask = (
                    (all_projects_raw['Имя клиента'].astype(str).str.strip() == str(row['Название проекта']).strip()) &
                    (all_projects_raw['Название проекта'].astype(str).str.strip() == str(row['Волна']).strip()) &
                    (all_projects_raw['Код анкеты'].astype(str).str.strip() == str(row['Код проекта']).strip())
                )
                if mask.any():
                    all_projects_raw.loc[mask, 'Полевой'] = 0
        
        # 5.2 Применяем добавленные проекты
        if not included_df.empty:
            for _, row in included_df.iterrows():
                mask = (
                    (all_projects_raw['Имя клиента'].astype(str).str.strip() == str(row['Название проекта']).strip()) &
                    (all_projects_raw['Название проекта'].astype(str).str.strip() == str(row['Волна']).strip()) &
                    (all_projects_raw['Код анкеты'].astype(str).str.strip() == str(row['Код проекта']).strip())
                )
                if mask.any():
                    all_projects_raw.loc[mask, 'Полевой'] = 1
        
        # ШАГ 6: РАЗДЕЛЕНИЕ НА ПОЛЕВЫЕ И НЕПОЛЕВЫЕ
        
        # Создаем полевые проекты
        field_projects = all_projects_raw[all_projects_raw['Полевой'] == 1].copy()
        non_field_projects = all_projects_raw[all_projects_raw['Полевой'] == 0].copy()
        
        st.session_state.cleaned_data['полевые_проекты'] = field_projects
        st.session_state.cleaned_data['неполевые_проекты'] = non_field_projects
        
        # ШАГ 7: ДОБАВЛЕНИЕ ЗОД
        
        if not field_projects.empty:
            field_with_zod = data_cleaner.add_zod_from_hierarchy(field_projects)
            st.session_state.cleaned_data['полевые_проекты'] = field_with_zod
        
        # ШАГ 8: СОЗДАНИЕ ИЕРАРХИИ И РАСЧЕТ ПЛАН/ФАКТ
        
        # Создание иерархии
        base_data = visit_calculator.extract_hierarchical_data(
            st.session_state.cleaned_data['полевые_проекты'],
            st.session_state.cleaned_data['сервизория']
        )
        
        st.session_state.visit_report['base_data'] = base_data
        st.session_state.visit_report['timestamp'] = datetime.now().isoformat()
        
        # Расчет план/факт
        if st.session_state.plan_calc_params and not base_data.empty:
            params = st.session_state.plan_calc_params
            source_df = st.session_state.cleaned_data['полевые_проекты']
            
            plan_result = visit_calculator.calculate_hierarchical_plan_on_date(
                base_data, source_df, params, st.session_state.cleaned_data['сервизория']
            )
            
            if plan_result is not None and not plan_result.empty:
                fact_result = visit_calculator.calculate_hierarchical_fact_on_date(
                    plan_result, source_df, params
                )
                
                final_result = visit_calculator._calculate_metrics(
                    fact_result, params, plan_result
                )
                
                st.session_state.visit_report['calculated_data'] = final_result
        
        st.session_state.processing_complete = True
        return True
        
    except Exception as e:
        st.session_state.last_error = {
            'step': 'Общая обработка',
            'error': str(e),
            'traceback': traceback.format_exc()
        }
        return False

# ==============================================
# САЙДБАР
# ==============================================
with st.sidebar:
    st.header("📊 Навигация")
    st.markdown("---")
    
    if st.button("🗑️ Сбросить все данные", type="secondary", use_container_width=True):
        for key in list(DEFAULT_STATE.keys()):
            st.session_state[key] = DEFAULT_STATE[key]
        st.rerun()
     
    st.markdown("---")
    st.subheader("📅 Параметры расчета")
    
    st.write("**Период расчета:**")
    today = date.today()
    first_day = date(today.year, today.month, 1)
    yesterday = today - timedelta(days=1)
    
    if yesterday < first_day:
        yesterday = first_day
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Дата начала", value=first_day)
    with col2:
        end_date = st.date_input("Дата окончания", value=yesterday)
    
    if start_date.month != end_date.month:
        end_date = start_date.replace(day=28)

    st.markdown("---")
    st.subheader("📊 Коэффициенты этапов")
    
    stage_weights = []
    default_weights = [0.8, 1.2, 1.0, 0.9]
    
    for i in range(1, 5):
        weight = st.slider(
            f"Вес этапа {i}",
            min_value=0.0,
            max_value=2.0,
            value=default_weights[i-1],
            step=0.1,
            key=f"stage_slider_{i}"
        )
        stage_weights.append(weight)
    
    total_weight = sum(stage_weights)
    if total_weight > 0:
        coefficients = [w/total_weight for w in stage_weights]
    else:
        coefficients = [0.25, 0.25, 0.25, 0.25]
    
    st.session_state.plan_calc_params = {
        'start_date': start_date,
        'end_date': end_date,
        'coefficients': coefficients
    }

# ==============================================
# ОСНОВНОЙ ИНТЕРФЕЙС
# ==============================================
tab1, tab2, tab3 = st.tabs(["📤 Загрузка данных", "📈 Отчеты", "⚙️ Настройки проектов"])

with tab1:
    st.title("📤 Загрузка исходных данных")
    st.markdown("Загрузите необходимые Excel файлы")
    
    # Загрузка файлов
    col1, col2,col3,col4,col5,col6  = st.columns(6)
    
    with col1:
        st.subheader("1. 📋 Портал (Массив.xlsx)")
        portal_file = st.file_uploader(
            "Загрузите файл Массив.xlsx",
            type=['xlsx', 'xls'],
            key="portal"
        )
        if portal_file:
            portal_df = validate_file_upload(portal_file, "Массив.xlsx")
            if portal_df is not None:
                st.session_state.uploaded_files['портал'] = portal_df
                st.success("✅ Портал загружен")
                display_file_preview(portal_df, "Просмотр данных портала")

                # Кнопка скачивания исходного массива
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    portal_df.to_excel(writer, sheet_name='Исходный_массив', index=False)
                
                st.download_button(
                    label="📥 Скачать исходный массив",
                    data=output.getvalue(),
                    file_name=f"исходный_массив_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                    use_container_width=True,
                    key="download_original_portal"
                )
    
    with col2:
        st.subheader("2. 📅 Проекты Сервизория")
        projects_file = st.file_uploader(
            "Загрузите файл Гугл таблица.xlsx",
            type=['xlsx', 'xls'],
            key="projects"
        )
        if projects_file:
            projects_df = validate_file_upload(projects_file, "Гугл таблица.xlsx")
            if projects_df is not None:
                st.session_state.uploaded_files['сервизория'] = projects_df
                st.success("✅ Проекты загружены")
                display_file_preview(projects_df, "Просмотр проектов")
    
    with col3:
        st.subheader("3. 📡 CXWAY (дополнительно)")
        cxway_file = st.file_uploader(
            "Загрузите файл CXWAY.xlsx",
            type=['xlsx', 'xls'],
            key="cxway"
        )
        if cxway_file:
            cxway_df = validate_file_upload(cxway_file, "CXWAY.xlsx")
            if cxway_df is not None:
                st.session_state.uploaded_files['cxway'] = cxway_df
                st.success("✅ CXWAY загружен")
                display_file_preview(cxway_df, "Просмотр данных CXWAY")
            
    with col4:
        st.subheader("4. 📱 Easymerch (дополнительно)")
        easymerch_file = st.file_uploader(
            "Загрузите файл Easymerch.xlsx",
            type=['xlsx', 'xls'],
            key="easymerch"
        )
        if easymerch_file:
            easymerch_df = validate_file_upload(easymerch_file, "Easymerch.xlsx")
            if easymerch_df is not None:
                st.session_state.uploaded_files['easymerch'] = easymerch_df
                st.success("✅ Easymerch загружен")
                display_file_preview(easymerch_df, "Просмотр данных Easymerch")

    with col5:
        st.subheader("5. 📱 Optima (дополнительно)")
        optima_file = st.file_uploader(
            "Загрузите файл Optima.xlsx",
            type=['xlsx', 'xls'],
            key="optima"
        )
        if optima_file:
            optima_df = validate_file_upload(optima_file, "Optima.xlsx")
            if optima_df is not None:
                st.session_state.uploaded_files['optima'] = optima_df
                st.success("✅ Optima загружен")
                display_file_preview(optima_df, "Просмотр данных Optima")

    with col6:
        st.subheader("6. 📱 ПроДата (дополнительно)")
        prodata_file = st.file_uploader(
            "Загрузите файл ПроДата.xlsx",
            type=['xlsx', 'xls'],
            key="prodata"
        )
        if prodata_file:
            prodata_df = validate_file_upload(prodata_file, "ПроДата.xlsx")
            if prodata_df is not None:
                st.session_state.uploaded_files['prodata'] = prodata_df
                st.success("✅ ПроДата загружен")
                display_file_preview(prodata_df, "Просмотр данных ПроДата")
        
    
    st.markdown("---")
    
    # Кнопка расчета
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        required_files = ['портал', 'сервизория']
        all_loaded = all(f in st.session_state.uploaded_files for f in required_files)
        
        if st.button(
            "🚀 РАССЧИТАТЬ ПЛАН/ФАКТ",
            type="primary",
            use_container_width=True,
            disabled=not all_loaded
        ):
            with st.spinner("Обработка данных и расчет план/факт..."):
                # Загружаем менеджер настроек
                if 'settings_manager' in st.session_state:
                    settings_manager = st.session_state.settings_manager
                else:
                    settings_manager = get_settings_manager()
                    st.session_state.settings_manager = settings_manager
                # Передаем в функцию
                success = process_all_data(settings_manager)

                if success:
                    st.success("✅ Расчет завершен! Перейдите на вкладку 'Отчеты'")
                else:
                    st.error("❌ Ошибка при расчете")
        
        if not all_loaded:
            st.info("📌 Загрузите оба основных файла для расчета")
            

with tab2:
    st.title("📈 Отчеты по полевым визитам")
    
    if ('visit_report' not in st.session_state or 
        st.session_state.visit_report.get('calculated_data') is None):
        st.info("Сначала выполните расчет на странице 'Загрузка данных'")
    else:
        tab_projects, tab_regions, tab_dsm = st.tabs(["📊 ПФ проекты", "🗺️ Регионы", "👥 DSM"])
        
        with tab_projects:
            data = st.session_state.visit_report['calculated_data']
            dataviz.create_planfact_tab(data, None)
            
            prodata_df = st.session_state.cleaned_data.get('prodata_processed')
            if prodata_df is not None and not prodata_df.empty:
                dataviz.create_prodata_table(prodata_df)
        
        with tab_regions:
            data = st.session_state.visit_report['calculated_data']
            dataviz.create_region_tab(data, None)
        
        with tab_dsm:
            data = st.session_state.visit_report['calculated_data']
            dataviz.create_dsm_tab(data, None)

# ============================================
# ВЫГРУЗКА ПОЛЕВЫХ ПРОЕКТОВ
# ============================================
if st.session_state.cleaned_data.get('полевые_проекты') is not None:
    st.markdown("---")
    st.subheader("📥 Выгрузка данных")
    
    field_projects_df = st.session_state.cleaned_data['полевые_проекты']
    
    
    # Исключаем ПроДата из выгрузки
    if 'Источник' in field_projects_df.columns:
        field_projects_df = field_projects_df[field_projects_df['Источник'] != 'Мониторинги']
    
    if not field_projects_df.empty:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            field_projects_df.to_excel(writer, sheet_name='Полевые_проекты', index=False)
        
        st.download_button(
            label="📥 Скачать все полевые проекты",
            data=output.getvalue(),
            file_name=f"полевые_проекты_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.info("Нет данных для выгрузки")

# ============================================
# ВКЛАДКА 3: НАСТРОЙКИ ПРОЕКТОВ
# ============================================
with tab3:
    st.title("⚙️ Пользовательские настройки проектов")
    st.markdown("Управляйте списком проектов для расчета план/факта")
    

    # Инициализируем менеджер настроек
    if 'settings_manager' not in st.session_state:
        st.session_state.settings_manager = get_settings_manager()
    
    manager = st.session_state.settings_manager
    
    if not manager.available:
        st.error("❌ Менеджер настроек недоступен. Проверьте секреты GitHub.")
        st.stop()
    
    # Загружаем текущие настройки
    excluded_df = manager.get_excluded_projects()
    included_df = manager.get_included_projects()
    
    # === СОСТОЯНИЯ ДЛЯ ИНТЕРФЕЙСА ===
    if 'selected_to_exclude' not in st.session_state:
        st.session_state.selected_to_exclude = []
    if 'selected_to_include' not in st.session_state:
        st.session_state.selected_to_include = []
    if 'show_history' not in st.session_state:
        st.session_state.show_history = False
    
    # === ОСНОВНОЙ ИНТЕРФЕЙС ===
    
    # Информация о текущих настройках
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📋 Исключенных проектов", len(excluded_df))
    with col2:
        st.metric("➕ Добавленных проектов", len(included_df))
    with col3:
        st.metric("🔄 Всего изменений", len(manager.get_history(days=9999)))
    
    st.markdown("---")
    
    # ДВЕ КОЛОНКИ ДЛЯ ТАБЛИЦ
    col_left, col_right = st.columns(2)
    
    # === ЛЕВАЯ КОЛОНКА: Проекты в расчете (можно исключить) ===
    with col_left:
        st.subheader("📊 Проекты В РАСЧЕТЕ")
        st.caption("Отметьте проекты, которые нужно ИСКЛЮЧИТЬ из расчета")
        
        # Получаем проекты из расчета (из session_state после process_all_data)
        if 'cleaned_data' in st.session_state and 'полевые_проекты' in st.session_state.cleaned_data:
            field_df = st.session_state.cleaned_data['полевые_проекты'].copy()
            field_df = field_df[field_df['Полевой'] == 1]
            
            # 🔥 ПРИМЕНЯЕМ НАСТРОЙКИ 🔥
            if field_df is not None and not field_df.empty:
                
                # 1. Убираем исключенные проекты
                if not excluded_df.empty:
                    for _, row in excluded_df.iterrows():
                        mask = (
                            (field_df['Имя клиента'] == row['Название проекта']) &
                            (field_df['Название проекта'] == row['Волна']) &
                            (field_df['Код анкеты'] == row['Код проекта'])
                        )
                        field_df = field_df[~mask]
                
                # 2. Добавляем проекты из included_df
                if not included_df.empty:
                    for _, row in included_df.iterrows():
                        new_row = {
                            'Имя клиента': row['Название проекта'],
                            'Название проекта': row['Волна'],
                            'Код анкеты': row['Код проекта'],
                            'ПО': row['ПО']
                        }
                        field_df = pd.concat([field_df, pd.DataFrame([new_row])], ignore_index=True)
                
                # Формируем DataFrame для отображения
                projects_in_calc = field_df[['Имя клиента', 'Название проекта', 'Код анкеты', 'ПО']].copy()
                projects_in_calc = projects_in_calc.rename(columns={
                    'Имя клиента': 'Название проекта',
                    'Название проекта': 'Волна',
                    'Код анкеты': 'Код проекта'
                })
                
                # Удаляем дубликаты
                projects_in_calc = projects_in_calc.drop_duplicates(keep='first').reset_index(drop=True)
                
                st.dataframe(projects_in_calc, use_container_width=True)
                
            
                
                # Мультиселект по Название проекта
                selected_clients = st.multiselect(
                    "Выберите проекты для исключения:",
                    options=projects_in_calc['Название проекта'].unique(),
                    key='multiselect_exclude'
                )
                
                if selected_clients and st.button("🗑️ Убрать выбранные из расчета", type="secondary", use_container_width=True):
                    # Берем строки с выбранными клиентами
                    selected_df = projects_in_calc[projects_in_calc['Название проекта'].isin(selected_clients)].copy()
                    
                    # Передаем как есть
                    success, msg = manager.add_to_excluded(selected_df)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                    
            else:
                st.info("⏳ Сначала выполните расчет на вкладке 'Загрузка данных'")
        else:
            st.info("⏳ Сначала выполните расчет на вкладке 'Загрузка данных'")
    
    # === ПРАВАЯ КОЛОНКА: Проекты НЕ в расчете  ===
    with col_right:
        st.subheader("📊 Проекты НЕ В РАСЧЕТЕ")
        st.caption("Отметьте проекты, которые нужно ДОБАВИТЬ в расчет")
        
        # Здесь будут неполевые проекты
        if 'cleaned_data' in st.session_state and 'неполевые_проекты' in st.session_state.cleaned_data:
            non_field_df = st.session_state.cleaned_data['неполевые_проекты'].copy()
            non_field_df = non_field_df[non_field_df['Полевой'] == 0]
            
            if non_field_df is not None and not non_field_df.empty:
                # Формируем DataFrame для отображения
                projects_not_in_calc = non_field_df[['Имя клиента', 'Название проекта', 'Код анкеты', 'ПО']].copy()
                projects_not_in_calc = projects_not_in_calc.rename(columns={
                    'Имя клиента': 'Название проекта',
                    'Название проекта': 'Волна',
                    'Код анкеты': 'Код проекта'
                })
                # Оставляем только нужные колонки
                projects_not_in_calc = projects_not_in_calc[['Название проекта', 'Волна', 'Код проекта', 'ПО']]
                
                # Удаляем дубликаты
                projects_not_in_calc = projects_not_in_calc.drop_duplicates(
                    subset=['Название проекта', 'Волна', 'Код проекта'], 
                    keep='first'
                ).reset_index(drop=True)
                
                # Исключаем уже добавленные проекты (из included_df)
                if not included_df.empty:
                    for _, row in included_df.iterrows():
                        mask = (
                            (projects_not_in_calc['Название проекта'] == row['Название проекта']) &
                            (projects_not_in_calc['Волна'] == row['Волна']) &
                            (projects_not_in_calc['Код проекта'] == row['Код проекта'])
                        )
                        projects_not_in_calc = projects_not_in_calc[~mask]
                
                if not projects_not_in_calc.empty:
                    st.dataframe(projects_not_in_calc, use_container_width=True)
                    
                    # Мультиселект для выбора проектов
                    project_options = projects_not_in_calc.apply(
                        lambda row: f"{row['Название проекта']} | {row['Волна']} | {row['Код проекта']}", 
                        axis=1
                    ).tolist()
                    
                    selected_projects = st.multiselect(
                        "Выберите проекты для добавления:",
                        options=project_options,
                        key='multiselect_include'
                    )
                    
                    if selected_projects and st.button("➕ Добавить выбранные в расчет", type="primary", use_container_width=True):
                        selected_rows = []
                        for s in selected_projects:
                            parts = s.split(' | ')
                            if len(parts) >= 3:
                                # Находим оригинальную строку в projects_not_in_calc
                                mask = (
                                    (projects_not_in_calc['Название проекта'] == parts[0]) &
                                    (projects_not_in_calc['Волна'] == parts[1]) &
                                    (projects_not_in_calc['Код проекта'] == parts[2])
                                )
                                if mask.any():
                                    original_row = projects_not_in_calc[mask].iloc[0]
                                    row_data = {
                                        'Название проекта': parts[0],
                                        'Волна': parts[1],
                                        'Код проекта': parts[2],
                                        'ПО': original_row['ПО'],
                                        'ФИО ОМ': ''
                                    }
                                    selected_rows.append(row_data)
                        
                        if selected_rows:
                            selected_df = pd.DataFrame(selected_rows)
                            success, msg = manager.add_to_included(selected_df)
                            if success:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                else:
                    st.info("✅ Нет неполевых проектов для добавления")
            else:
                st.info("⏳ Неполевые проекты появятся после расчета")
        else:
            st.info("⏳ Сначала выполните расчет на вкладке 'Загрузка данных'")
        
        st.markdown("---") 
    
    # === ТЕКУЩИЕ НАСТРОЙКИ ===
    with st.expander("📋 Текущие списки исключенных/добавленных проектов"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🗑️ Исключенные проекты")
            if not excluded_df.empty:
                st.dataframe(excluded_df, use_container_width=True)
                
                if st.button("Очистить список исключенных", key="clear_excluded"):
                    success, msg = manager.remove_from_excluded(excluded_df)
                    if success:
                        st.success("Список исключенных очищен")
                        st.rerun()
            else:
                st.info("Список пуст")
        
        with col2:
            st.subheader("➕ Добавленные проекты")
            if not included_df.empty:
                st.dataframe(included_df, use_container_width=True)
                
                if st.button("Очистить список добавленных", key="clear_included"):
                    success, msg = manager.remove_from_included(included_df)
                    if success:
                        st.success("Список добавленных очищен")
                        st.rerun()
            else:
                st.info("Список пуст")
    
    # === ИСТОРИЯ ИЗМЕНЕНИЙ ===
    with st.expander("📜 История изменений"):
        history_df = manager.get_history(days=30)
        if not history_df.empty:
            # Переименовываем колонки для единообразия
            history_display = history_df.rename(columns={
                'project_name': 'Название проекта',
                'wave_name': 'Волна', 
                'project_code': 'Код проекта'
            })
            
            # Выводим все нужные колонки
            st.dataframe(history_display, use_container_width=True)
        else:
            st.info("История изменений пуста")
        
    # ============================================
    # ПРОБЛЕМНЫЕ ПРОЕКТЫ
    # ============================================
    with st.expander("🔴 Проблемные проекты", expanded=False):
        if 'cleaned_data' in st.session_state and 'сервизория' in st.session_state.cleaned_data:
            google_df = st.session_state.cleaned_data['сервизория']
            field_df = st.session_state.cleaned_data.get('полевые_проекты')
            
            if field_df is not None:
                problematic_projects = data_cleaner.check_problematic_projects(
                    google_df, field_df
                )
                
                if not problematic_projects.empty:
                    st.dataframe(problematic_projects, use_container_width=True)
                    
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        problematic_projects.to_excel(writer, sheet_name='Проблемные_проекты', index=False)
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="⬇️ Скачать проблемные проекты",
                        data=excel_buffer,
                        file_name="проблемные_проекты.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary",
                        use_container_width=True
                    )
                else:
                    st.info("✅ Проблемных проектов не найдено")
            else:
                st.info("Нет данных для проверки")
            
    
    # === КНОПКИ СОХРАНЕНИЯ И ПЕРЕСЧЕТА ===
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        if st.button("💾 Сохранить настройки", type="primary", use_container_width=True):
            st.success("Настройки сохранены в GitHub!")
    
    with col2:
        if st.button("🔄 Пересчитать", type="secondary", use_container_width=True):
            with st.spinner("🔄 Пересчет план/факта с учетом настроек..."):
                try:
                    # Просто вызываем ту же функцию, что и при первом расчете
                    success = process_all_data(manager)
                    
                    if success:
                        st.success("✅ Пересчет завершен! Перейдите на вкладку 'Отчеты'")
                        st.rerun()
                    else:
                        st.error("❌ Ошибка при пересчете")
                except Exception as e:
                    st.error(f"❌ Ошибка: {str(e)}")
                
    
    with col3:
        if st.button("🗑️ Сбросить все настройки", use_container_width=True):
            success, msg = manager.clear_all_settings()
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)








