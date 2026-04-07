# app.py
# draft 4.1 - simplified
import streamlit as st
import pandas as pd
import sys
import os
import traceback
from datetime import date, datetime, timedelta
from io import BytesIO
from github_settings import get_settings_manager, get_plan_adjustment_manager
# Инициализация временных корректировок
if 'temp_adjustments' not in st.session_state:
    st.session_state.temp_adjustments = []

# ФУНКЦИЯ КЭШИРОВАНИЯ ЗАГРУЗКИ EXCEL
@st.cache_data
def load_excel(file_obj, file_key):
    """Загружает Excel с кэшированием. file_key - уникальный ключ для кэша"""
    if file_obj is None:
        return None
    try:
        return pd.read_excel(file_obj, dtype=str)
    except Exception as e:
        st.error(f"Ошибка загрузки файла {file_key}: {e}")
        return None

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
    'plan_calc_params': None,
    'data_calculated': False,  # ← ДОБАВИТЬ
    'last_calculation_hash': None  # ← ДОБАВИТЬ
}

for key, default_value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_value

# Вспомогательные функции

def process_all_data(settings_manager=None, force_recalc=False):
    """Полная обработка данных и расчет план/факт"""
    
    # Если данные уже посчитаны - сразу выходим
    if not force_recalc and st.session_state.get('data_calculated', False):
        return True
    
    # Загрузка, очистка    
    try:
        import time
        start_total = time.time()
        start = start_total

        if 'debug_times' not in st.session_state:
            st.session_state.debug_times = []
        st.session_state.debug_times = []
        
        # Проверяем наличие основных файлов
        required_files = ['портал', 'сервизория']
        missing_files = [f for f in required_files if f not in st.session_state.uploaded_files]
        
        if missing_files:
            return False
        
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
        
        # Добавление признака полевой проект
        google_with_field = data_cleaner.update_field_projects_flag(st.session_state.cleaned_data['сервизория'])
        st.session_state.cleaned_data['сервизория'] = google_with_field
        st.session_state.debug_times.append(f"[DEBUG] Очистка: {time.time() - start:.2f} сек")
        start = time.time()

        
        # ============================================
        # ОБОГАЩЕНИЕ ДАННЫХ (ОПТИМИЗИРОВАННО)
        # ============================================
        
        # Добавляем поле 'Полевой' и 'ПО'
        array_with_field = data_cleaner.add_field_flag_to_array(st.session_state.cleaned_data['портал'])
        array_with_portal = data_cleaner.add_portal_to_array(array_with_field, google_with_field)
        array_with_portal = data_cleaner.remove_cxway_from_portal(array_with_portal, google_with_field)
        st.session_state.cleaned_data['портал_с_полем'] = array_with_portal
        
        # Разделение на полевые/неполевые
        field_df, non_field_df = data_cleaner.split_array_by_field_flag(array_with_portal)
        
        # Загружаем настройки
        if settings_manager is None:
            settings_manager = get_settings_manager()
        
        excluded_df = settings_manager.get_excluded_projects()
        included_df = settings_manager.get_included_projects()
        
        # ============================================
        # ВЕКТОРИЗОВАННОЕ ПРИМЕНЕНИЕ НАСТРОЕК
        # ============================================
        
        # Применяем исключенные проекты (делаем их неполевыми)
        if not excluded_df.empty and field_df is not None and not field_df.empty:
            # Создаем временные ключи для быстрого поиска
            field_df['_temp_key'] = (
                field_df['Имя клиента'].astype(str) + '|' + 
                field_df['Название проекта'].astype(str) + '|' + 
                field_df['Код анкеты'].astype(str)
            )
            excluded_df['_temp_key'] = (
                excluded_df['Название проекта'].astype(str) + '|' + 
                excluded_df['Волна'].astype(str) + '|' + 
                excluded_df['Код проекта'].astype(str)
            )
            
            # Одна операция вместо цикла
            mask = field_df['_temp_key'].isin(excluded_df['_temp_key'])
            field_df.loc[mask, 'Полевой'] = 0
            
            # Удаляем временные колонки
            field_df = field_df.drop('_temp_key', axis=1)
            # excluded_df не сохраняем, не нужно удалять
        
        # Применяем добавленные проекты (делаем их полевыми)
        if not included_df.empty and non_field_df is not None and not non_field_df.empty:
            # Создаем временные ключи
            non_field_df['_temp_key'] = (
                non_field_df['Имя клиента'].astype(str) + '|' + 
                non_field_df['Название проекта'].astype(str) + '|' + 
                non_field_df['Код анкеты'].astype(str)
            )
            included_df['_temp_key'] = (
                included_df['Название проекта'].astype(str) + '|' + 
                included_df['Волна'].astype(str) + '|' + 
                included_df['Код проекта'].astype(str)
            )
            
            # Одна операция вместо цикла
            mask = non_field_df['_temp_key'].isin(included_df['_temp_key'])
            non_field_df.loc[mask, 'Полевой'] = 1
            
            # Удаляем временные колонки
            non_field_df = non_field_df.drop('_temp_key', axis=1)
        
        # Объединяем все проекты в один датасет
        all_projects = pd.concat([field_df, non_field_df], ignore_index=True)
        st.session_state.cleaned_data['all_projects'] = all_projects

        # 2. Создаем обновленные датасеты
        field_df = all_projects[all_projects['Полевой'] == 1].copy()
        non_field_df = all_projects[all_projects['Полевой'] == 0].copy()
        
        # Создаем датасеты на основе актуального значения Полевой
        st.session_state.cleaned_data['полевые_проекты'] = all_projects[all_projects['Полевой'] == 1].copy()
        st.session_state.cleaned_data['неполевые_проекты'] = all_projects[all_projects['Полевой'] == 0].copy()
        
        # ============================================
        # ВЕКТОРИЗОВАННОЕ ДОБАВЛЕНИЕ ЗОД
        # ============================================
        
        # Добавление ЗОД из встроенного справочника (векторизовано)
        if field_df is not None and not field_df.empty:
            # Получаем ЗОД через словарь (быстрее чем apply)
            field_df_with_zod = data_cleaner.add_zod_from_hierarchy(field_df)
            
            # Создаем словарь для быстрого обновления
            if not field_df_with_zod.empty:
                # Создаем временный ключ для поиска
                all_projects['_temp_key'] = (
                    all_projects['Имя клиента'].astype(str) + '|' + 
                    all_projects['Название проекта'].astype(str) + '|' + 
                    all_projects['Код анкеты'].astype(str)
                )
                
                # Создаем маппинг ЗОД по ключу
                zod_mapping = {}
                for _, row in field_df_with_zod.iterrows():
                    key = (
                        str(row['Имя клиента']) + '|' + 
                        str(row['Название проекта']) + '|' + 
                        str(row['Код анкеты'])
                    )
                    zod_mapping[key] = row['ЗОД']
                
                # Обновляем ЗОД одним проходом
                mask = all_projects['_temp_key'].isin(zod_mapping.keys())
                all_projects.loc[mask, 'ЗОД'] = all_projects.loc[mask, '_temp_key'].map(zod_mapping)
                
                # Удаляем временную колонку
                all_projects = all_projects.drop('_temp_key', axis=1)
        
        # ============================================
        # ОБРАБОТКА ДОПОЛНИТЕЛЬНЫХ ИСТОЧНИКОВ
        # ============================================
        
        # Обработка Easymerch (если есть)
        easymerch_processed = None
        easymerch_raw = st.session_state.uploaded_files.get('easymerch')
        if easymerch_raw is not None:
            easymerch_processed = data_cleaner.clean_easymerch(easymerch_raw, google_with_field)
            if easymerch_processed is not None and not easymerch_processed.empty:
                st.session_state.cleaned_data['easymerch_processed'] = easymerch_processed
        
        # Обработка Optima (если есть)
        optima_processed = None
        optima_raw = st.session_state.uploaded_files.get('optima')
        if optima_raw is not None:
            try:
                optima_processed = data_cleaner.clean_optima(optima_raw, google_with_field)
                if optima_processed is not None and not optima_processed.empty:
                    st.session_state.cleaned_data['optima_processed'] = optima_processed
            except Exception as e:
                st.warning(f"⚠️ Ошибка при обработке Optima: {e}")

        # Обработка ПроДата (Мониторинги)
        prodata_processed = None
        prodata_raw = st.session_state.uploaded_files.get('prodata')
        if prodata_raw is not None:
            try:
                prodata_processed = data_cleaner.clean_prodata(prodata_raw, google_with_field)
                if prodata_processed is not None and not prodata_processed.empty:
                    st.session_state.cleaned_data['prodata_processed'] = prodata_processed
            except Exception as e:
                st.warning(f"⚠️ Ошибка при обработке ПроДата: {e}")
        
        # Обработка CXWAY (если есть)
        cxway_processed = None
        cxway_raw = st.session_state.uploaded_files.get('cxway')
        if cxway_raw is not None:
            cxway_processed = data_cleaner.clean_cxway(cxway_raw, None, google_with_field)
            
            # Разделяем CXWAY на полевые и неполевые
            if cxway_processed is not None and not cxway_processed.empty:
                cxway_field = cxway_processed[cxway_processed['Полевой'] == 1]
                cxway_non_field = cxway_processed[cxway_processed['Полевой'] == 0]
                
                # Неполевые добавляем в неполевые проекты сразу
                if not cxway_non_field.empty:
                    st.session_state.cleaned_data['неполевые_проекты'] = pd.concat([
                        st.session_state.cleaned_data['неполевые_проекты'],
                        cxway_non_field
                    ], ignore_index=True)
        
        # ============================================
        # ФИНАЛЬНОЕ ОБЪЕДИНЕНИЕ ВСЕХ ИСТОЧНИКОВ
        # ============================================
        
        sources_for_merge = []
        
        if field_df is not None and not field_df.empty:
            sources_for_merge.append(field_df_with_zod)
            
        if cxway_processed is not None and not cxway_processed.empty:
            cxway_field_only = cxway_processed[cxway_processed['Полевой'] == 1].copy()
            if not cxway_field_only.empty:
                sources_for_merge.append(cxway_field_only)
        
        if easymerch_processed is not None and not easymerch_processed.empty:
            sources_for_merge.append(easymerch_processed)
            
        if optima_processed is not None and not optima_processed.empty:
            sources_for_merge.append(optima_processed)
        
        if prodata_processed is not None and not prodata_processed.empty:
            st.session_state.cleaned_data['prodata_processed'] = prodata_processed
        
        if sources_for_merge:
            all_field_projects = pd.concat(sources_for_merge, ignore_index=True)
            st.session_state.cleaned_data['полевые_проекты'] = all_field_projects
        else:
            st.session_state.cleaned_data['полевые_проекты'] = pd.DataFrame()

        # ============================================
        # ДОБАВЛЯЕМ ЗОД ДЛЯ ВСЕХ ПОЛЕВЫХ ПРОЕКТОВ
        # ============================================
        if not st.session_state.cleaned_data['полевые_проекты'].empty:
            all_field_projects_with_zod = data_cleaner.add_zod_from_hierarchy(
                st.session_state.cleaned_data['полевые_проекты']
            )
            st.session_state.cleaned_data['полевые_проекты'] = all_field_projects_with_zod
        
        # Добавляем ЗОД для неполевых проектов
        if not st.session_state.cleaned_data['неполевые_проекты'].empty:
            non_field_with_zod = data_cleaner.add_zod_from_hierarchy(
                st.session_state.cleaned_data['неполевые_проекты']
            )
            st.session_state.cleaned_data['неполевые_проекты'] = non_field_with_zod

        # ============================================
        # ПРИМЕНЕНИЕ НАСТРОЕК ДЛЯ CXWAY/EASYMERCH/OPTIMA
        # ============================================
        
        # 1. Обрабатываем excluded_df (исключаем проекты из расчета)
        if not excluded_df.empty and not all_field_projects.empty:
            excluded_df['_key'] = (
                excluded_df['Название проекта'].astype(str).str.strip() + '|' +
                excluded_df['Волна'].astype(str).str.strip() + '|' +
                excluded_df['Код проекта'].astype(str).str.strip()
            )
            
            all_field_projects['_key'] = (
                all_field_projects['Имя клиента'].astype(str).str.strip() + '|' +
                all_field_projects['Название проекта'].astype(str).str.strip() + '|' +
                all_field_projects['Код анкеты'].astype(str).str.strip()
            )
            
            all_field_projects = all_field_projects[~all_field_projects['_key'].isin(excluded_df['_key'])]
            all_field_projects = all_field_projects.drop('_key', axis=1)
            excluded_df = excluded_df.drop('_key', axis=1)
            
            st.session_state.cleaned_data['полевые_проекты'] = all_field_projects
        
        # 2. Обрабатываем included_df (добавляем проекты в расчет)
        if not included_df.empty:
            included_df['_key'] = (
                included_df['Название проекта'].astype(str).str.strip() + '|' +
                included_df['Волна'].astype(str).str.strip() + '|' +
                included_df['Код проекта'].astype(str).str.strip()
            )
            
            # Создаем ключи в all_field_projects
            if not all_field_projects.empty:
                all_field_projects['_key'] = (
                    all_field_projects['Имя клиента'].astype(str).str.strip() + '|' +
                    all_field_projects['Название проекта'].astype(str).str.strip() + '|' +
                    all_field_projects['Код анкеты'].astype(str).str.strip()
                )
                existing_keys = set(all_field_projects['_key'])
            else:
                existing_keys = set()
            
            # Находим проекты, которых еще нет
            new_projects = included_df[~included_df['_key'].isin(existing_keys)].copy()
            
            if not new_projects.empty:
                all_source_rows = []
                not_found_projects = []
                new_keys = set(new_projects['_key'])
                
                # Создаем ключи в источниках
                if 'cxway_processed' in locals() and cxway_processed is not None and not cxway_processed.empty:
                    if '_key' not in cxway_processed.columns:
                        cxway_processed['_key'] = (
                            cxway_processed['Имя клиента'].astype(str).str.strip() + '|' +
                            cxway_processed['Название проекта'].astype(str).str.strip() + '|' +
                            cxway_processed['Код анкеты'].astype(str).str.strip()
                        )
                
                if 'optima_processed' in locals() and optima_processed is not None and not optima_processed.empty:
                    if '_key' not in optima_processed.columns:
                        optima_processed['_key'] = (
                            optima_processed['Имя клиента'].astype(str).str.strip() + '|' +
                            optima_processed['Название проекта'].astype(str).str.strip() + '|' +
                            optima_processed['Код анкеты'].astype(str).str.strip()
                        )
                
                if 'easymerch_processed' in locals() and easymerch_processed is not None and not easymerch_processed.empty:
                    if '_key' not in easymerch_processed.columns:
                        easymerch_processed['_key'] = (
                            easymerch_processed['Имя клиента'].astype(str).str.strip() + '|' +
                            easymerch_processed['Название проекта'].astype(str).str.strip() + '|' +
                            easymerch_processed['Код анкеты'].astype(str).str.strip()
                        )
                
                # Портал
                portal_df = None
                if 'портал_с_полем' in st.session_state.cleaned_data:
                    portal_df = st.session_state.cleaned_data['портал_с_полем'].copy()
                    if portal_df is not None and not portal_df.empty and '_key' not in portal_df.columns:
                        portal_df['_key'] = (
                            portal_df['Имя клиента'].astype(str).str.strip() + '|' +
                            portal_df['Название проекта'].astype(str).str.strip() + '|' +
                            portal_df['Код анкеты'].astype(str).str.strip()
                        )
                
                # Поиск в CXWAY
                if 'cxway_processed' in locals() and cxway_processed is not None and not cxway_processed.empty:
                    matches = cxway_processed[cxway_processed['_key'].isin(new_keys)].copy()
                    if not matches.empty:
                        matches['Полевой'] = 1
                        matches['Источник'] = 'CXWAY (добавлен вручную)'
                        all_source_rows.append(matches)
                        new_keys -= set(matches['_key'])
                
                # Поиск в портале
                if new_keys and portal_df is not None and not portal_df.empty:
                    matches = portal_df[portal_df['_key'].isin(new_keys)].copy()
                    if not matches.empty:
                        matches['Полевой'] = 1
                        matches['Источник'] = 'Портал (добавлен вручную)'
                        all_source_rows.append(matches)
                        new_keys -= set(matches['_key'])
                
                # Поиск в Optima
                if new_keys and 'optima_processed' in locals() and optima_processed is not None and not optima_processed.empty:
                    matches = optima_processed[optima_processed['_key'].isin(new_keys)].copy()
                    if not matches.empty:
                        matches['Полевой'] = 1
                        matches['Источник'] = 'Optima (добавлен вручную)'
                        all_source_rows.append(matches)
                        new_keys -= set(matches['_key'])
                
                # Поиск в Easymerch
                if new_keys and 'easymerch_processed' in locals() and easymerch_processed is not None and not easymerch_processed.empty:
                    matches = easymerch_processed[easymerch_processed['_key'].isin(new_keys)].copy()
                    if not matches.empty:
                        matches['Полевой'] = 1
                        matches['Источник'] = 'Easymerch (добавлен вручную)'
                        all_source_rows.append(matches)
                        new_keys -= set(matches['_key'])
                
                # Проекты не найдены
                if new_keys:
                    not_found_df = new_projects[new_projects['_key'].isin(new_keys)]
                    for _, row in not_found_df.iterrows():
                        not_found_projects.append({
                            'Клиент': row['Название проекта'],
                            'Волна': row['Волна'],
                            'Код проекта': row['Код проекта'],
                            'ПО': row.get('ПО', ''),
                            'Проверенные источники': 'CXWAY, Портал, Optima, Easymerch'
                        })
                
                # Добавляем найденные строки
                if all_source_rows:
                    new_df = pd.concat(all_source_rows, ignore_index=True)
                    if all_field_projects.empty:
                        all_field_projects = new_df
                    else:
                        all_field_projects = pd.concat([all_field_projects, new_df], ignore_index=True)
                
                # Сохраняем список ненайденных проектов
                if not_found_projects:
                    st.session_state.not_found_projects = pd.DataFrame(not_found_projects)
                else:
                    st.session_state.not_found_projects = pd.DataFrame()
                
                # Удаляем временные колонки
                if '_key' in all_field_projects.columns:
                    all_field_projects = all_field_projects.drop('_key', axis=1)
            
            # Удаляем временные колонки
            included_df = included_df.drop('_key', axis=1)
            
            # Сохраняем обновленные полевые проекты
            st.session_state.cleaned_data['полевые_проекты'] = all_field_projects
            
    

        all_projects_export = pd.concat([
            st.session_state.cleaned_data['полевые_проекты'],
            st.session_state.cleaned_data['неполевые_проекты']
        ], ignore_index=True)
        st.session_state.cleaned_data['all_projects'] = all_projects_export
    

        
        # Создание иерархии
        import time as tm
        start_hier = tm.time()
        st.write(f"🔍 НАЧАЛО ИЕРАРХИИ: {tm.time() - start_total:.2f} сек от старта")
        
        base_data = visit_calculator.extract_hierarchical_data(
            st.session_state.cleaned_data['полевые_проекты'],
            st.session_state.cleaned_data['сервизория']
        )

        st.write(f"🔍 КОНЕЦ ИЕРАРХИИ: {tm.time() - start_hier:.2f} сек (время выполнения)")
        st.write(f"🔍 ВСЕГО СТРОК В ИЕРАРХИИ: {len(base_data)}")
        
        st.session_state.visit_report['base_data'] = base_data
        st.session_state.visit_report['timestamp'] = datetime.now().isoformat()

        st.session_state.debug_times.append(f"[DEBUG] Иерархия: {time.time() - start:.2f} сек")
        start = time.time()
        
        # Расчет план/факт
        if st.session_state.plan_calc_params and not base_data.empty:
            params = st.session_state.plan_calc_params
            source_df = st.session_state.cleaned_data['полевые_проекты']
            
            plan_result = visit_calculator.calculate_hierarchical_plan_on_date(
                base_data, source_df, params, st.session_state.cleaned_data['сервизория']
            )
            
            st.session_state.debug_times.append(f"[DEBUG] План: {time.time() - start:.2f} сек")
            start = time.time()
            
            if plan_result is not None and not plan_result.empty:
                fact_result = visit_calculator.calculate_hierarchical_fact_on_date(
                    plan_result, source_df, params
                )
                
                st.session_state.debug_times.append(f"[DEBUG] Факт: {time.time() - start:.2f} сек")
                start = time.time()
                
                final_result = visit_calculator._calculate_metrics(
                    fact_result, params, plan_result
                )
                
                st.session_state.visit_report['calculated_data'] = final_result
            
                st.session_state.debug_times.append(f"[DEBUG] Метрики: {time.time() - start:.2f} сек")
                
        st.session_state.debug_times.append(f"[DEBUG] ВСЕГО: {time.time() - start_total:.2f} сек")
                
        # Выводим предупреждение о ненайденных проектах
        if 'not_found_projects' in st.session_state and not st.session_state.not_found_projects.empty:
            st.warning("⚠️ Следующие проекты не найдены в загруженных данных:")
            st.dataframe(st.session_state.not_found_projects, width='stretch')
            st.info("💡 Проверьте: возможно, визиты по этим проектам не были загружены, или указан неверный портал.")
            
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
    
    if st.button("🗑️ Сбросить все данные", type="secondary", width='stretch'):
        # Очищаем кэш загрузки Excel
        load_excel.clear()
        # Сбрасываем все переменные session_state
        for key in list(DEFAULT_STATE.keys()):
            st.session_state[key] = DEFAULT_STATE[key]
        st.success("✅ Все данные и кэш очищены")
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
    # Показываем сообщения о расчете после перезагрузки
    if 'show_messages' in st.session_state and st.session_state.show_messages:
        st.write("### ⏱️ Время выполнения:")
        for msg in st.session_state.calculation_messages:
            st.write(msg)
        st.success("✅ Расчет завершен!")
        st.session_state.show_messages = False
        
    st.title("📤 Загрузка исходных данных")
    st.markdown("Загрузите необходимые Excel файлы")
    
    # Размещаем загрузчики БЕЗ формы (чтобы они обновляли session_state сразу)
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.subheader("1. 📋 Портал (Массив.xlsx)")
        portal_file = st.file_uploader(
            "Загрузите файл Массив.xlsx",
            type=['xlsx', 'xls'],
            key="portal_uploader",
            label_visibility="collapsed"
        )
    
    with col2:
        st.subheader("2. 📅 Проекты Сервизория")
        projects_file = st.file_uploader(
            "Загрузите файл Гугл таблица.xlsx",
            type=['xlsx', 'xls'],
            key="projects_uploader",
            label_visibility="collapsed"
        )
    
    with col3:
        st.subheader("3. 📡 CXWAY (дополнительно)")
        cxway_file = st.file_uploader(
            "Загрузите файл CXWAY.xlsx",
            type=['xlsx', 'xls'],
            key="cxway_uploader",
            label_visibility="collapsed"
        )
    
    with col4:
        st.subheader("4. 📱 Easymerch (дополнительно)")
        easymerch_file = st.file_uploader(
            "Загрузите файл Easymerch.xlsx",
            type=['xlsx', 'xls'],
            key="easymerch_uploader",
            label_visibility="collapsed"
        )
    
    with col5:
        st.subheader("5. 📱 Optima (дополнительно)")
        optima_file = st.file_uploader(
            "Загрузите файл Optima.xlsx",
            type=['xlsx', 'xls'],
            key="optima_uploader",
            label_visibility="collapsed"
        )
    
    with col6:
        st.subheader("6. 📱 ПроДата (дополнительно)")
        prodata_file = st.file_uploader(
            "Загрузите файл ПроДата.xlsx",
            type=['xlsx', 'xls'],
            key="prodata_uploader",
            label_visibility="collapsed"
        )
    
    st.markdown("---")
    
    # КНОПКА РАССЧИТАТЬ (вне формы)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Проверяем наличие файлов (они уже в session_state)
        portal_exists = st.session_state.get('portal_uploader') is not None
        projects_exists = st.session_state.get('projects_uploader') is not None
        
        if portal_exists and projects_exists:
            if st.button("🚀 РАССЧИТАТЬ ПЛАН/ФАКТ", type="primary", width='stretch'):
                with st.spinner("📥 Загрузка файлов и обработка данных..."):
                    
                    # 1. ЗАГРУЖАЕМ ФАЙЛЫ из session_state
                    portal_file_obj = st.session_state.get('portal_uploader')
                    projects_file_obj = st.session_state.get('projects_uploader')
                    cxway_file_obj = st.session_state.get('cxway_uploader')
                    easymerch_file_obj = st.session_state.get('easymerch_uploader')
                    optima_file_obj = st.session_state.get('optima_uploader')
                    prodata_file_obj = st.session_state.get('prodata_uploader')
                    
                    portal_df = load_excel(portal_file_obj, "портал")
                    projects_df = load_excel(projects_file_obj, "проекты")
                    cxway_df = load_excel(cxway_file_obj, "cxway")
                    easymerch_df = load_excel(easymerch_file_obj, "easymerch")
                    optima_df = load_excel(optima_file_obj, "optima")
                    prodata_df = load_excel(prodata_file_obj, "prodata")
                    
                    # 2. СОХРАНЯЕМ В SESSION_STATE.uploaded_files
                    if portal_df is not None:
                        st.session_state.uploaded_files['портал'] = portal_df
                    if projects_df is not None:
                        st.session_state.uploaded_files['сервизория'] = projects_df
                    if cxway_df is not None:
                        st.session_state.uploaded_files['cxway'] = cxway_df
                    if easymerch_df is not None:
                        st.session_state.uploaded_files['easymerch'] = easymerch_df
                    if optima_df is not None:
                        st.session_state.uploaded_files['optima'] = optima_df
                    if prodata_df is not None:
                        st.session_state.uploaded_files['prodata'] = prodata_df
                    
                    # Сбрасываем флаг перед расчетом
                    st.session_state.data_calculated = False
                    
                    # 3. ЗАПУСКАЕМ РАСЧЕТ
                    if 'settings_manager' in st.session_state:
                        settings_manager = st.session_state.settings_manager
                    else:
                        settings_manager = get_settings_manager()
                        st.session_state.settings_manager = settings_manager
                    
                    success = process_all_data(settings_manager)
                    
                    if success:
                        # Сохраняем сообщения в память
                        st.session_state.calculation_messages = st.session_state.debug_times.copy()
                        st.session_state.show_messages = True 
                        st.session_state.data_calculated = True
                        st.rerun()
                    else:
                        st.error("❌ Ошибка при расчете")
        else:
            st.info("📌 Загрузите оба основных файла для расчета")
            st.button("🚀 РАССЧИТАТЬ ПЛАН/ФАКТ", type="primary", width='stretch', disabled=True)
            

with tab2:
    st.title("📈 Отчеты по полевым визитам")

    if not st.session_state.get('data_calculated', False):
        st.info("📌 Сначала выполните расчет на вкладке 'Загрузка данных'")
    
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
            width='stretch'
        )
    else:
        st.info("Нет данных для выгрузки")

# # ============================================
# # ВЫГРУЗКА НЕПОЛЕВЫХ ПРОЕКТОВ
# # ============================================
# if st.session_state.cleaned_data.get('неполевые_проекты') is not None:
#     st.markdown("---")
    
#     non_field_projects_df = st.session_state.cleaned_data['неполевые_проекты']
    
#     if not non_field_projects_df.empty:
#         output = BytesIO()
#         with pd.ExcelWriter(output, engine='openpyxl') as writer:
#             non_field_projects_df.to_excel(writer, sheet_name='Неполевые_проекты', index=False)
        
#         st.download_button(
#             label="📥 Скачать все неполевые проекты",
#             data=output.getvalue(),
#             file_name=f"неполевые_проекты_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             type="secondary",
#             width='stretch'
#         )
#     else:
#         st.info("Нет данных для выгрузки")

# # ============================================
# # ВЫГРУЗКА ALL_PROJECTS (ОБЪЕДИНЕННЫЙ ДАТАСЕТ)
# # ============================================
# if 'all_projects' in st.session_state.cleaned_data:
#     st.markdown("---")
    
#     all_projects_df = st.session_state.cleaned_data['all_projects']
    
#     if not all_projects_df.empty:
#         output = BytesIO()
#         with pd.ExcelWriter(output, engine='openpyxl') as writer:
#             all_projects_df.to_excel(writer, sheet_name='Все_проекты', index=False)
        
#         st.download_button(
#             label="📥 Скачать all_projects (все проекты)",
#             data=output.getvalue(),
#             file_name=f"all_projects_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             type="secondary",
#             width='stretch'
#         )
#     else:
#         st.info("Нет данных для выгрузки")

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
                
                st.dataframe(projects_in_calc, width='stretch')
                
            
                
                # Мультиселект по Название проекта
                selected_clients = st.multiselect(
                    "Выберите проекты для исключения:",
                    options=projects_in_calc['Название проекта'].unique(),
                    key='multiselect_exclude'
                )
                
                if selected_clients and st.button("🗑️ Убрать выбранные из расчета", type="secondary", width='stretch'):
                    # Берем строки с выбранными клиентами
                    selected_df = projects_in_calc[projects_in_calc['Название проекта'].isin(selected_clients)].copy()
                    
                    # Передаем как есть
                    success, msg = manager.add_to_excluded(selected_df)
                    if success:
                        st.success(msg)
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
                    st.dataframe(projects_not_in_calc, width='stretch')
                    
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
                    
                    if selected_projects and st.button("➕ Добавить выбранные в расчет", type="primary", width='stretch'):
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
                st.dataframe(excluded_df, width='stretch')
                
                if st.button("Очистить список исключенных", key="clear_excluded"):
                    success, msg = manager.remove_from_excluded(excluded_df)
                    if success:
                        st.success("Список исключенных очищен")
            else:
                st.info("Список пуст")
        
        with col2:
            st.subheader("➕ Добавленные проекты")
            if not included_df.empty:
                st.dataframe(included_df, width='stretch')
                
                if st.button("Очистить список добавленных", key="clear_included"):
                    success, msg = manager.remove_from_included(included_df)
                    if success:
                        st.success("Список добавленных очищен")
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
            st.dataframe(history_display, width='stretch')
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
                    st.dataframe(problematic_projects, width='stretch')
                    
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
                        width='stretch'
                    )
                else:
                    st.info("✅ Проблемных проектов не найдено")
            else:
                st.info("Нет данных для проверки")
            
    
    # === КНОПКИ СОХРАНЕНИЯ И ПЕРЕСЧЕТА ===
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        if st.button("💾 Сохранить настройки", type="primary", width='stretch'):
            st.success("Настройки сохранены в GitHub!")
    
    with col2:
        if st.button("🔄 Пересчитать", type="secondary", width='stretch'):
            st.warning("⚠️ Чтобы применить настройки, нажмите 'РАССЧИТАТЬ' на вкладке 'Загрузка данных'")
                    
        
        with col3:
            if st.button("🗑️ Сбросить все настройки", width='stretch'):
                success, msg = manager.clear_all_settings()
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    # === КОРРЕКТИРОВКА ПЛАНА ===
    
    st.markdown("---")
    st.subheader("✂️ Корректировка плана")
    st.caption("✂️ Срез: введите отрицательное число (например, -1000). ➕ Добавление: введите положительное число (например, 1000)")

    # Получаем менеджер корректировок
    plan_adj_manager = get_plan_adjustment_manager()

    # Кэшируем корректировки из GitHub (один раз)
    if 'cached_saved_adjustments' not in st.session_state:
        st.session_state.cached_saved_adjustments = plan_adj_manager.get_all_adjustments()
    
    saved_adjustments = st.session_state.cached_saved_adjustments
    
    # Выбираем проект для корректировки
    if 'cleaned_data' in st.session_state and 'полевые_проекты' in st.session_state.cleaned_data:
        field_df = st.session_state.cleaned_data['полевые_проекты'].copy()
        field_df = field_df[field_df['Полевой'] == 1]
        
        if not field_df.empty:
            # Кэшируем уникальные проекты для корректировки
            field_df_hash = hash(field_df.values.tobytes()) if not field_df.empty else 0
            
            if st.session_state.get('cached_projects_hash') != field_df_hash:
                st.session_state.cached_unique_projects = field_df[['Имя клиента', 'Название проекта', 'Код анкеты']].drop_duplicates()
                st.session_state.cached_project_options = st.session_state.cached_unique_projects.apply(
                    lambda row: f"{row['Имя клиента']} | {row['Название проекта']} | {row['Код анкеты']}",
                    axis=1
                ).tolist()
                st.session_state.cached_projects_hash = field_df_hash
            
            unique_projects = st.session_state.get('cached_unique_projects', pd.DataFrame())
            project_options = st.session_state.get('cached_project_options', [])

            
            # ========== ПОСТОЯННАЯ ТАБЛИЦА КОРРЕКТИРОВОК (только проекты с изменениями) ==========
            with st.expander("📋 Текущие корректировки проектов", expanded=True):
           
                # Добавляем временные корректировки
                all_adjustments = saved_adjustments.copy()
                for temp in st.session_state.temp_adjustments:
                    key = (temp['project_name'], temp['wave_name'], temp['project_code'])
                    current = all_adjustments.get(key, 0)
                    all_adjustments[key] = current + temp['adjustment_value']
                
                adjustments_list = []
                for (p_name, w_name, p_code), adj_value in all_adjustments.items():
                    if adj_value < 0:
                        status = f"⬇️ срез {abs(adj_value)}"
                    elif adj_value > 0:
                        status = f"⬆️ добавление {adj_value}"
                    else:
                        status = "✅ без изменений"
                    
                    adjustments_list.append({
                        'Название проекта': p_name,
                        'Волна': w_name,
                        'Код проекта': p_code,
                        'Корректировка': adj_value,
                        'Статус': status
                    })
                
                if adjustments_list:
                    df_adj_table = pd.DataFrame(adjustments_list)
                    st.dataframe(df_adj_table, width='stretch', hide_index=True)
                else:
                    st.info("✅ Нет проектов с корректировками")
            # =============================================================
            
            
            selected_projects = st.multiselect(
                "Выберите проект для корректировки",
                options=project_options,
                key="plan_adjustment_select",
                max_selections=1
            )
            
            if selected_projects:
                parts = selected_projects[0].split(' | ')
                if len(parts) >= 3:
                    project_name = parts[0]
                    wave_name = parts[1]
                    project_code = parts[2]
                    
                    # Текущая суммарная корректировка (сохраненная + временная)
                    saved_adj = plan_adj_manager.get_total_adjustment(project_name, wave_name, project_code)
                    temp_adj = sum(t['adjustment_value'] for t in st.session_state.temp_adjustments 
                                  if t['project_name'] == project_name 
                                  and t['wave_name'] == wave_name 
                                  and t['project_code'] == project_code)
                    current_adjustment = saved_adj + temp_adj
                    
                    # Отображаем текущую корректировку
                    if current_adjustment < 0:
                        st.warning(f"📊 Текущая корректировка: **{current_adjustment}** (план уменьшен)")
                    elif current_adjustment > 0:
                        st.info(f"📊 Текущая корректировка: **+{current_adjustment}** (план увеличен)")
                    else:
                        st.success("📊 Текущая корректировка: **0** (план не изменен)")
                    
                    # Ввод новой корректировки
                    col1, col2 = st.columns(2)
                    with col1:
                        adjustment_value = st.number_input(
                            "Введите значение (отрицательное = срез, положительное = добавление)",
                            value=0,
                            step=1,
                            key="plan_adjustment_value"
                        )
                    
                    with col2:
                        if st.button("➕ Добавить корректировку", key="add_temp_adjustment_btn"):
                            if adjustment_value != 0:
                                # Добавляем во временный список
                                st.session_state.temp_adjustments.append({
                                    'project_name': project_name,
                                    'wave_name': wave_name,
                                    'project_code': project_code,
                                    'adjustment_value': adjustment_value
                                })
                                st.success(f"✅ Корректировка {adjustment_value} добавлена во временный список")
                            else:
                                st.warning("⚠️ Введите ненулевое значение")
                    
                    # Кнопка очистки временных корректировок для проекта
                    if temp_adj != 0:
                        if st.button("🗑️ Очистить временные корректировки для проекта", key="clear_temp_adjustments_btn"):
                            st.session_state.temp_adjustments = [t for t in st.session_state.temp_adjustments 
                                if not (t['project_name'] == project_name 
                                       and t['wave_name'] == wave_name 
                                       and t['project_code'] == project_code)]
                            st.success("✅ Временные корректировки очищены")
                    
                    # ========== ИСТОРИЯ КОРРЕКТИРОВОК (как в "История изменений") ==========
                    with st.expander("📜 История корректировок", expanded=False):
                        # Получаем всю историю
                        all_history = plan_adj_manager.get_adjustments()
                        
                        if all_history:
                            history_df = pd.DataFrame(all_history)
                            # Фильтруем: оставляем только записи с ненулевой корректировкой
                            history_df = history_df[history_df['adjustment_value'] != 0]
                            
                            if not history_df.empty:
                                # Форматируем дату
                                history_df['created_at'] = pd.to_datetime(history_df['created_at']).dt.strftime('%Y-%m-%d %H:%M')
                                
                                # Переводим действие на русский
                                history_df['Действие'] = history_df['adjustment_value'].apply(
                                    lambda x: '✂️ Срез' if x < 0 else '➕ Добавление'
                                )
                                
                                # Переименовываем колонки как в "История изменений"
                                history_df = history_df.rename(columns={
                                    'created_at': 'Дата',
                                    'created_by': 'Пользователь',
                                    'project_name': 'Проект',
                                    'wave_name': 'Волна',
                                    'project_code': 'Код',
                                    'adjustment_value': 'Значение'
                                })
                                
                                # Выводим нужные колонки
                                st.dataframe(
                                    history_df[['Дата', 'Пользователь', 'Действие', 'Проект', 'Волна', 'Код', 'Значение']],
                                    width='stretch',
                                    hide_index=True
                                )
                            else:
                                st.info("История корректировок пуста")
                        else:
                            st.info("История корректировок пуста")
                    
                    # ========== ВРЕМЕННЫЕ КОРРЕКТИРОВКИ (не сохраненные) ==========
                    if st.session_state.temp_adjustments:
                        st.markdown("---")
                        st.caption("⏳ Несохраненные корректировки:")
                        temp_df = pd.DataFrame(st.session_state.temp_adjustments)
                        temp_df = temp_df.rename(columns={
                            'project_name': 'Проект',
                            'wave_name': 'Волна',
                            'project_code': 'Код',
                            'adjustment_value': 'Значение'
                        })
                        st.dataframe(temp_df[['Проект', 'Волна', 'Код', 'Значение']], width='stretch', hide_index=True)
            
            # ========== КНОПКИ ДЛЯ КОРРЕКТИРОВОК ==========
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 1, 1])
            with col1:
                if st.button("💾 Сохранить корректировки", key="save_adjustments_btn"):
                    if st.session_state.temp_adjustments:
                        for adj in st.session_state.temp_adjustments:
                            plan_adj_manager.add_adjustment(
                                adj['project_name'], 
                                adj['wave_name'], 
                                adj['project_code'], 
                                adj['adjustment_value']
                            )
                        st.session_state.temp_adjustments = []
                        st.session_state.pop('cached_saved_adjustments', None)
                        st.success("✅ Корректировки сохранены в GitHub!")
                    else:
                        st.info("Нет временных корректировок для сохранения")
            with col2:
                if st.button("🔄 Пересчитать с корректировками", key="recalc_with_adjustments_btn"):
                    with st.spinner("🔄 Пересчет план/факта с учетом корректировок..."):
                        try:
                            st.session_state.data_calculated = False
                            success = process_all_data(manager, force_recalc=True)
                            if success:
                                st.session_state.data_calculated = True
                                st.success("✅ Пересчет завершен!")
                                st.rerun()
                            else:
                                st.error("❌ Ошибка при пересчете")
                        except Exception as e:
                            st.error(f"❌ Ошибка: {str(e)}")
            with col3:
                if st.button("🗑️ Сбросить все временные корректировки", key="reset_temp_adjustments_btn"):
                    if st.session_state.temp_adjustments:
                        st.session_state.temp_adjustments = []
                        st.success("✅ Все временные корректировки очищены")
                    else:
                        st.info("Нет временных корректировок")
            # =============================================
        else:
            st.info("⏳ Нет полевых проектов для корректировки")
    else:
        st.info("⏳ Сначала выполните расчет")





