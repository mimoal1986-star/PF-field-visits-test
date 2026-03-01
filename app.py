# app.py
# draft 3.1 - cleaned (without diagnostics)
import streamlit as st
import pandas as pd
import sys
import os
import traceback
from datetime import date, datetime, timedelta
from io import BytesIO

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
    'excel_files': {},
    'processing_complete': False,
    'processing_stats': {},
    'last_error': None,
    'visit_report': {} 
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
        if 'автокодификация' in file_name.lower() or 'автокодификац' in file_name.lower():
            excel_reader = pd.ExcelFile(file_obj)
            sheet_names = excel_reader.sheet_names
            
            target_sheet = None
            for sheet in sheet_names:
                sheet_lower = sheet.lower()
                if any(name in sheet_lower for name in ['код', 'codes', 'data', 'основная', 'main']):
                    target_sheet = sheet
                    break
            
            if target_sheet is None:
                return None
            else:
                pass
            
            df = pd.read_excel(file_obj, sheet_name=target_sheet, dtype=str)
        else:
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

def process_single_step(step_func, step_name, *args):
    """Обработка одного этапа с обработкой ошибок"""
    try:
        result = step_func(*args)
        return result, None
    except Exception as e:
        error_msg = f"Ошибка на этапе '{step_name}': {str(e)[:200]}"
        st.session_state['last_error'] = {
            'step': step_name,
            'error': str(e),
            'timestamp': datetime.now().isoformat()
        }
        return None, error_msg

def process_field_projects_with_stats():
    """Основная функция обработки полевых проектов"""
    try:
        required_keys = ['сервизория', 'портал', 'автокодификация']
        missing_keys = [k for k in required_keys if k not in st.session_state.cleaned_data and 
                       k not in st.session_state.uploaded_files]
        
        if missing_keys:
            return False
        
        google_df = st.session_state.cleaned_data.get('сервизория')
        if google_df is None:
            return False 
            
        array_df = st.session_state.cleaned_data.get('портал')
        if array_df is None:
            return False 
            
        autocoding_df = st.session_state.uploaded_files.get('автокодификация')
        hierarchy_df = st.session_state.uploaded_files.get('иерархия')
        
        if google_df is None or array_df is None or autocoding_df is None:
            return False
        
        google_updated = data_cleaner.update_field_projects_flag(google_df)
        if google_updated is None:
            return False
        st.session_state.cleaned_data['сервизория_с_полем'] = google_updated
        st.session_state.cleaned_data['сервизория'] = google_updated 
        
        array_updated = data_cleaner.add_field_flag_to_array(array_df)
        array_with_portal = data_cleaner.add_portal_to_array(array_updated, google_updated)
        array_updated = array_with_portal
        if array_updated is None:
            return False
        st.session_state.cleaned_data['портал_с_полем'] = array_updated
        
        field_df, non_field_df = data_cleaner.split_array_by_field_flag(array_updated)
        if field_df is None and non_field_df is None:
            return False
        
        st.session_state.cleaned_data['полевые_проекты'] = field_df
        st.session_state.cleaned_data['неполевые_проекты'] = non_field_df
        
        cxway_df = st.session_state.uploaded_files.get('cxway')
        cxway_processed = None
        
        if cxway_df is not None:
            cxway_processed = data_cleaner.clean_cxway(
                cxway_df, 
                hierarchy_df, 
                google_updated
            )
            
            if cxway_processed is not None and not cxway_processed.empty:
                if field_df is not None and not field_df.empty:
                    combined_field_projects = data_cleaner.merge_field_projects(
                        field_df, 
                        cxway_processed
                    )
                    field_df = combined_field_projects
                    st.session_state.cleaned_data['полевые_проекты'] = field_df
                else:
                    field_df = cxway_processed
                    st.session_state.cleaned_data['полевые_проекты'] = field_df
        
        excel_output = data_cleaner.export_split_array_to_excel(field_df, non_field_df)
        if excel_output:
            st.session_state.excel_files['разделенный_массив'] = excel_output

        if field_df is not None and not field_df.empty:
            field_projects_df = st.session_state.cleaned_data.get('полевые_проекты')
            
            if field_projects_df is not None and not field_projects_df.empty:
                base_data = visit_calculator.extract_hierarchical_data(field_projects_df, google_df)
            else:
                base_data = pd.DataFrame()
            
            if 'visit_report' not in st.session_state:
                st.session_state.visit_report = {}
            
            st.session_state.visit_report['base_data'] = base_data
            st.session_state.visit_report['timestamp'] = datetime.now().isoformat()
        
        return True
        
    except Exception as e:
        return False

# ==============================================
# САЙДБАР
# ==============================================
with st.sidebar:
    st.subheader("Тестовая среда")
    st.header("📊 Навигация")
    st.markdown("---")
    
    if st.button("🗑️ Сбросить все данные", type="secondary", use_container_width=True):
        for key in list(DEFAULT_STATE.keys()):
            st.session_state[key] = DEFAULT_STATE[key]
        st.rerun()
     
    st.markdown("---")
    st.subheader("📅 Параметры расчета план/факта")
    
    st.write("**Период расчета:**")
    today = date.today()
    first_day = date(today.year, today.month, 1)
    yesterday = today - timedelta(days=1)
    
    if yesterday < first_day:
        yesterday = first_day
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Дата начала",
            value=first_day,
        )
    with col2:
        end_date = st.date_input(
            "Дата окончания",
            value=yesterday,
        )
    
    if start_date.month != end_date.month:
        end_date = start_date.replace(day=28)

    st.markdown("---")
    st.subheader("📊 Коэффициенты этапов")
    
    st.write("**Распределение весов по этапам (0-2):**")
    
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
    
    st.session_state['plan_calc_params'] = {
        'start_date': start_date,
        'end_date': end_date,
        'coefficients': coefficients
    }

# Основной интерфейс
tab1, tab2 = st.tabs(["📤 Загрузка данных", "📈 Отчеты"])

with tab1:
    st.title("📤 Загрузка исходных данных")
    st.markdown("Загрузите 4 Excel файла для формирования отчетов")
    
    # ==============================================
    # СЕКЦИЯ 1: ЗАГРУЗКА ФАЙЛОВ
    # ==============================================
    
    upload_cols = st.columns(2)
    
    with upload_cols[0]:
        st.subheader("1. 📋 Портал (Массив.xlsx)")
        portal_file = st.file_uploader(
            "Загрузите файл Массив.xlsx",
            type=['xlsx', 'xls'],
            key="portal"
        )
        portal_df = validate_file_upload(portal_file, "Массив.xlsx")
        if portal_df is not None:
            st.session_state.uploaded_files['портал'] = portal_df
            display_file_preview(portal_df, "Просмотр данных портала")
    
        st.subheader("2. 🏷️ Автокодификация")
        autocoding_file = st.file_uploader(
            "Загрузите файл Автокодификация.xlsx",
            type=['xlsx', 'xls'],
            key="autocoding"
        )
        autocoding_df = validate_file_upload(autocoding_file, "Автокодификация.xlsx")
        if autocoding_df is not None:
            st.session_state.uploaded_files['автокодификация'] = autocoding_df
            display_file_preview(autocoding_df, "Просмотр автокодификации")
    
    with upload_cols[1]:
        st.subheader("3. 📅 Проекты Сервизория")
        projects_file = st.file_uploader(
            "Загрузите файл Гугл таблица.xlsx",
            type=['xlsx', 'xls'],
            key="projects"
        )
        projects_df = validate_file_upload(projects_file, "Гугл таблица.xlsx")
        if projects_df is not None:
            st.session_state.uploaded_files['сервизория'] = projects_df
            display_file_preview(projects_df, "Просмотр проектов")
    
        st.subheader("4. 👥 Иерархия ЗОД-АСС")
        hierarchy_file = st.file_uploader(
            "Загрузите файл ЗОД+АСС.xlsx",
            type=['xlsx', 'xls'],
            key="hierarchy"
        )
        hierarchy_df = validate_file_upload(hierarchy_file, "ЗОД+АСС.xlsx")
        if hierarchy_df is not None:
            st.session_state.uploaded_files['иерархия'] = hierarchy_df
            display_file_preview(hierarchy_df, "Просмотр иерархии")

        st.subheader("5. 📡 CXWAY (дополнительно)")
        cxway_file = st.file_uploader(
            "Загрузите файл CXWAY.xlsx",
            type=['xlsx', 'xls'],
            key="cxway"
        )
        cxway_df = validate_file_upload(cxway_file, "CXWAY.xlsx")
        if cxway_df is not None:
            st.session_state.uploaded_files['cxway'] = cxway_df
            display_file_preview(cxway_df, "Просмотр данных CXWAY")
    
    # ==============================================
    # СЕКЦИЯ 2: СТАТУС И ОБРАБОТКА
    # ==============================================
    st.markdown("---")
    
    if st.session_state.uploaded_files:
        required_files = ['портал', 'автокодификация', 'сервизория', 'иерархия']
        loaded_count = sum(1 for f in required_files if f in st.session_state.uploaded_files)
        
        if loaded_count == 4:
            st.markdown("---")
            st.subheader("🚀 Полная обработка данных")
            
            col1, col2 = st.columns([3, 1])
            with col2:
                process_disabled = st.session_state.processing_complete
                if st.button("🚀 ЗАПУСТИТЬ ОБРАБОТКУ", 
                            type="primary",
                            disabled=process_disabled,
                            use_container_width=True):
                    
                    st.session_state.processing_complete = False
                    st.session_state.excel_files.clear()
                    st.session_state.processing_stats.clear()
    
                    try:
                        from data_cleaner import data_cleaner
                        
                        with st.expander("📊 **Ход обработки**", expanded=False):
                            # ЭТАП 1: Проверка
                            missing_files = [f for f in required_files if f not in st.session_state.uploaded_files]
                            if missing_files:
                                raise ValueError(f"Отсутствуют файлы: {', '.join(missing_files)}")
                            
                            # ЭТАП 2: Очистка портала
                            portal_raw = st.session_state.uploaded_files['портал']
                            portal_cleaned, portal_error = process_single_step(
                                data_cleaner.clean_array, "Очистка портала", portal_raw
                            )
                            
                            if portal_error:
                                portal_cleaned = portal_raw
                            
                            st.session_state.cleaned_data['портал'] = portal_cleaned
                            
                            # ЭТАП 3: Очистка проектов
                            projects_raw = st.session_state.uploaded_files['сервизория']
                            projects_cleaned, projects_error = process_single_step(
                                data_cleaner.clean_google, "Очистка проектов", projects_raw
                            )
                            
                            if projects_error:
                                projects_cleaned = projects_raw
                            
                            st.session_state.cleaned_data['сервизория'] = projects_cleaned
                            
                            # ЭТАП 4: Обогащение массива
                            if 'портал' in st.session_state.cleaned_data and 'сервизория' in st.session_state.cleaned_data:
                                enriched_result, enrich_error = process_single_step(
                                    data_cleaner.enrich_array_with_project_codes,
                                    "Обогащение массива",
                                    st.session_state.cleaned_data['портал'],
                                    st.session_state.cleaned_data['сервизория']
                                )
                                
                                if enrich_error:
                                    enriched_array = st.session_state.cleaned_data['портал']
                                    discrepancy_df = pd.DataFrame()
                                    stats = {'filled': 0, 'total': 0}
                                else:
                                    enriched_array, discrepancy_df, stats = enriched_result
                                    st.session_state.cleaned_data['портал'] = enriched_array
                                    
                                    if not discrepancy_df.empty:
                                        st.session_state['array_discrepancies'] = discrepancy_df
                                        st.session_state['discrepancy_stats'] = stats

                            # ЭТАП 4.5: Добавление ЗОД в массив
                            if 'портал' in st.session_state.cleaned_data and 'иерархия' in st.session_state.uploaded_files:
                                hierarchy_df = st.session_state.uploaded_files['иерархия']
                                array_with_zod, zod_error = process_single_step(
                                    data_cleaner.add_zod_from_hierarchy,
                                    "Добавление ЗОД",
                                    st.session_state.cleaned_data['портал'],
                                    hierarchy_df
                                )
                                
                                if not zod_error and array_with_zod is not None:
                                    st.session_state.cleaned_data['портал'] = array_with_zod
                                    
                            # ЭТАП 5: Разделение на полевые/неполевые проекты
                            try:
                                field_success = process_field_projects_with_stats()
                            except Exception as e:
                                field_success = False
                            
                            # ЭТАП 6: Выгрузка в Excel
                            
                            # Массив
                            if 'портал' in st.session_state.cleaned_data:
                                array_with_field = st.session_state.cleaned_data.get('портал_с_полем', st.session_state.cleaned_data['портал'])
                                array_excel, array_export_error = process_single_step(
                                    data_cleaner.export_array_to_excel,
                                    "Выгрузка массива",
                                    array_with_field
                                )
                                
                                if array_excel:
                                    st.session_state.excel_files['массив'] = array_excel
                            
                            # Проекты
                            if 'сервизория' in st.session_state.cleaned_data:
                                projects_excel, projects_export_error = process_single_step(
                                    data_cleaner.export_to_excel,
                                    "Выгрузка проектов",
                                    st.session_state.uploaded_files['сервизория'],
                                    st.session_state.cleaned_data['сервизория'],
                                    "очищенные_проекты"
                                )
                                
                                if projects_excel:
                                    st.session_state.excel_files['проекты'] = projects_excel
                            
                            # Сохранение статистики
                            st.session_state.processing_stats = {
                                'timestamp': datetime.now().isoformat(),
                                'portal_rows': len(portal_cleaned),
                                'projects_rows': len(projects_cleaned),
                                'excel_files': len(st.session_state.excel_files),
                                'enriched_codes': stats.get('filled', 0) if 'stats' in locals() else 0
                            }
                            
                            st.session_state.processing_complete = True
                            
                    except ImportError as e:
                        st.session_state['last_error'] = {
                            'step': 'Общая обработка',
                            'error': str(e),
                            'traceback': traceback.format_exc()
                        }
                    except Exception as e:
                        st.session_state['last_error'] = {
                            'step': 'Общая обработка',
                            'error': str(e),
                            'traceback': traceback.format_exc()
                        }
    
    # ==============================================
    # СЕКЦИЯ 3: РЕЗУЛЬТАТЫ
    # ==============================================
    if st.session_state.processing_complete:
        st.markdown("---")
        st.subheader("✅ Результаты обработки")
        
        # Загрузка файлов
        st.markdown("### 📥 Загрузка результатов")
        
        download_cols = st.columns(2)
        
        with download_cols[0]:
            if 'массив' in st.session_state.excel_files:
                st.download_button(
                    label="⬇️ Скачать очищенный_массив.xlsx",
                    data=st.session_state.excel_files['массив'],
                    file_name="очищенный_массив.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        
        with download_cols[1]:
            if 'проекты' in st.session_state.excel_files:
                st.download_button(
                    label="⬇️ Скачать очищенные_проекты.xlsx",
                    data=st.session_state.excel_files['проекты'],
                    file_name="очищенные_проекты.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
        # НОВАЯ КНОПКА - Разделенный массив
        st.markdown("---")
        st.subheader("🎯 Разделенный массив")
        
        if 'разделенный_массив' in st.session_state.excel_files:
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="⬇️ Скачать разделенный_массив.xlsx",
                    data=st.session_state.excel_files['разделенный_массив'],
                    file_name="разделенный_массив.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            st.markdown("---")
            st.subheader("📥 Отдельные файлы")
            
            cols = st.columns(2)
            with cols[0]:
                if 'полевые_проекты' in st.session_state.cleaned_data:
                    field_excel = data_cleaner.export_field_projects_only(
                        st.session_state.cleaned_data['полевые_проекты']
                    )
                    if field_excel:
                        st.download_button(
                            label="⬇️ ТОЛЬКО Полевые проекты",
                            data=field_excel,
                            file_name="полевые_проекты.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
            
            with cols[1]:
                if 'неполевые_проекты' in st.session_state.cleaned_data:
                    non_field_excel = data_cleaner.export_non_field_projects_only(
                        st.session_state.cleaned_data['неполевые_проекты']
                    )
                    if non_field_excel:
                        st.download_button(
                            label="⬇️ ТОЛЬКО Неполевые проекты (без дубликатов)",
                            data=non_field_excel,
                            file_name="неполевые_проекты.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                        
        # ============================================
        # 🆕 ДАННЫЕ ДЛЯ РАСЧЕТА ПЛАН/ФАКТА
        # ============================================

        # Кнопка расчета план/факт
        if st.button("📊 Рассчитать план/факт", type="primary", use_container_width=True):
            if 'plan_calc_params' in st.session_state and 'visit_report' in st.session_state:
                required_keys = ['сервизория', 'портал']
                missing_keys = [k for k in required_keys if k not in st.session_state.cleaned_data]
                
                if not missing_keys:
                    if 'base_data' not in st.session_state.visit_report:
                        st.stop() 
    
                    base_data = st.session_state.visit_report['base_data']
                    source_df = st.session_state.cleaned_data['полевые_проекты']
                    params = st.session_state['plan_calc_params']
                    
                    plan_result = visit_calculator.calculate_hierarchical_plan_on_date(
                        base_data,
                        source_df,
                        params
                    )
                    
                    if plan_result is None or plan_result.empty:
                        st.stop()
                    
                    fact_result = visit_calculator.calculate_hierarchical_fact_on_date(
                        plan_result, 
                        source_df, 
                        params
                    )
                    
                    final_result = visit_calculator._calculate_metrics(
                        fact_result,
                        params,
                        plan_result
                    )

                    st.session_state['visit_report'] = {
                        'calculated_data': final_result,
                        'hierarchy': base_data,
                        'timestamp': datetime.now().isoformat()
                    }

                    st.markdown("---")
                    st.subheader("📊 Результаты расчета план/факт")
                    
                    if not final_result.empty:
                        with st.expander("👀 Просмотреть таблицу ПланФакт", expanded=True):
                            st.dataframe(final_result, use_container_width=True, height=300)
                        
                        excel_buffer = BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            final_result.to_excel(writer, sheet_name='Данные_план_факт', index=False)
                        excel_buffer.seek(0)
                        
                        st.download_button(
                            label="⬇️ Скачать Excel",
                            data=excel_buffer,
                            file_name="данные_план_факт.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
        
        # ============================================
        # 🆕 ПРОВЕРКА ПРОБЛЕМНЫХ ПРОЕКТОВ
        # ============================================
        if 'cleaned_data' in st.session_state and 'сервизория' in st.session_state.cleaned_data:
            st.markdown("---")
            st.subheader("🔴 Проблемные проекты")
            
            google_df = st.session_state.cleaned_data['сервизория']
            autocoding_df = st.session_state.uploaded_files.get('автокодификация')
            array_df = st.session_state.cleaned_data.get('портал_с_полем', 
                   st.session_state.cleaned_data.get('портал'))
    
            problematic_projects = data_cleaner.check_problematic_projects(
                google_df, autocoding_df, array_df
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
        
        # Просмотр данных
        st.markdown("---")
        st.subheader("🔍 Просмотр данных")

# ==============================================
# ВТОРАЯ ВКЛАДКА: ОТЧЕТЫ
# ==============================================
with tab2:
    st.title("📈 Отчеты по полевым визитам")
    
    if ('visit_report' not in st.session_state or 
        st.session_state.visit_report.get('calculated_data') is None):
        st.info("Сначала запустите расчет на странице 'Загрузка данных'")
    else:
        tab_projects, tab_regions, tab_dsm = st.tabs(["📊 ПФ проекты", "🗺️ Регионы", "👥 DSM"])
        
        with tab_projects:
            data = st.session_state.visit_report['calculated_data']
            hierarchy_df = st.session_state.uploaded_files.get('иерархия')
            dataviz.create_planfact_tab(data, hierarchy_df)
        
        with tab_regions:
            data = st.session_state.visit_report['calculated_data']
            hierarchy_df = st.session_state.uploaded_files.get('иерархия')
            dataviz.create_region_tab(data, hierarchy_df)
        
        with tab_dsm:
            data = st.session_state.visit_report['calculated_data']
            hierarchy_df = st.session_state.uploaded_files.get('иерархия')
            dataviz.create_dsm_tab(data, hierarchy_df)
