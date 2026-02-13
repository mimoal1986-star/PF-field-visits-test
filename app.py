# app.py
# draft 3.0 
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
    page_title="ИУ Аудиты - Аналитика Тест",
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

# # ПРОВЕРКА ВРЕМЕННАЯ
# def check_required_columns():
#     """Простая проверка всех необходимых колонок"""
#     errors = []
    
#     # 1. Проверяем основные данные
#     required_data = ['портал', 'сервизория', 'полевые_проекты']
#     for key in required_data:
#         if key not in st.session_state.cleaned_data:
#             errors.append(f"❌ Нет данных: '{key}'")
    
#     # 2. Проверяем колонки в массиве (портал)
#     if 'портал' in st.session_state.cleaned_data:
#         portal_df = st.session_state.cleaned_data['портал']
#         portal_required = ['Код анкеты', 'Дата визита']
#         for col in portal_required:
#             if col not in portal_df.columns:
#                 errors.append(f"❌ В массиве нет колонки: '{col}'")
    
#     # 3. Проверяем visit_report
#     if 'visit_report' not in st.session_state:
#         errors.append("❌ Нет visit_report (не запущен расчёт иерархии)")
#     elif 'base_data' not in st.session_state.visit_report:
#         errors.append("❌ В visit_report нет base_data")
    
#     # 4. Проверяем параметры расчёта
#     if 'plan_calc_params' not in st.session_state:
#         errors.append("❌ Не настроены параметры расчёта в сайдбаре")
    
#     # Возвращаем ошибки (пустой список = всё ок)
#     return errors
# # ПРОВЕРКА ВРЕМЕННАЯ

# Вспомогательные функции
def validate_file_upload(file_obj, file_name):
    """Проверка и загрузка файла"""
    if file_obj is None:
        return None
    
    try:
        # Для автокодификации читаем вкладку "Коды"
        if 'автокодификация' in file_name.lower() or 'автокодификац' in file_name.lower():
            # Создаем Excel reader для проверки вкладок
            excel_reader = pd.ExcelFile(file_obj)
            sheet_names = excel_reader.sheet_names
            
            # Ищем вкладку с кодами
            target_sheet = None
            for sheet in sheet_names:
                sheet_lower = sheet.lower()
                if any(name in sheet_lower for name in ['код', 'codes', 'data', 'основная', 'main']):
                    target_sheet = sheet
                    break
            
            # Если не нашли - берем первую вкладку
            if target_sheet is None:
                st.error(f"❌ В файле автокодификации не найдена вкладка с кодами. Доступные вкладки: {', '.join(sheet_names)}")
                st.error("❌ Нужна вкладка с названием содержащим 'код' или 'codes'")
                return None  
            else:
                st.info(f"✅ Используем вкладку АК: '{target_sheet}'")
            
            df = pd.read_excel(file_obj, sheet_name=target_sheet, dtype=str)
        else:
            # Для остальных файлов читаем как обычно
            df = pd.read_excel(file_obj, dtype=str)
            
        if df.empty:
            st.warning(f"Файл {file_name} пуст")
            return None
        return df
    except Exception as e:
        st.error(f"Ошибка чтения {file_name}: {str(e)[:200]}")
        return None

def display_file_preview(df, title):
    """Отображение предпросмотра файла"""
    if df is not None and not df.empty:
        with st.expander(f"👀 {title}"):
            st.dataframe(df.head(10), use_container_width=True)
            st.caption(f"Всего строк: {len(df):,}, колонок: {len(df.columns)}")

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
        # Проверяем наличие данных
        required_keys = ['сервизория', 'портал', 'автокодификация']
        missing_keys = [k for k in required_keys if k not in st.session_state.cleaned_data and 
                       k not in st.session_state.uploaded_files]
        
        if missing_keys:
            st.error(f"❌ Отсутствуют данные: {', '.join(missing_keys)}")
            return False
        
        # Получаем данные
        google_df = st.session_state.cleaned_data.get('сервизория')
        if google_df is None:
            st.error("❌ Не удалось получить данные сервизория")
            return False 
            
        array_df = st.session_state.cleaned_data.get('портал')
        if array_df is None:
            st.error("❌ Не удалось получить данные портал")
            return False 
            
        autocoding_df = st.session_state.uploaded_files.get('автокодификация')
        
        hierarchy_df = st.session_state.uploaded_files.get('иерархия')
        if hierarchy_df is None:
            st.warning("Файл иерархии не загружен, CXWAY не будет обработан")
            cxway_processed = None
        
        if google_df is None or array_df is None or autocoding_df is None:
            st.error("❌ Не удалось получить все необходимые данные")
            return False
        
        st.write("### 🎯 Шаг 1: Определение полевых проектов")
        with st.spinner("Анализирую автокодификацию..."):
            google_updated = data_cleaner.update_field_projects_flag(google_df)
            if google_updated is None:
                return False
            st.session_state.cleaned_data['сервизория_с_полем'] = google_updated
            st.session_state.cleaned_data['сервизория'] = google_updated 
        
        st.write("### 🎯 Шаг 2: Добавление признака в массив")
        with st.spinner("Сопоставляю коды проектов..."):
            array_updated = data_cleaner.add_field_flag_to_array(array_df)
            array_with_portal = data_cleaner.add_portal_to_array(array_updated, google_updated)
            array_updated = array_with_portal
            if array_updated is None:
                return False
            st.session_state.cleaned_data['портал_с_полем'] = array_updated

        
        st.write("### 🎯 Шаг 3: Разделение на полевые/неполевые")
        with st.spinner("Фильтрую данные..."):
            field_df, non_field_df = data_cleaner.split_array_by_field_flag(array_updated)
            if field_df is None and non_field_df is None:
                return False
            
            st.session_state.cleaned_data['полевые_проекты'] = field_df
            st.session_state.cleaned_data['неполевые_проекты'] = non_field_df
        
        # ПРОВЕРКА ВРЕМЕННАЯ
        st.write("### 🔍 Проверка field_df")
        st.write(f"Колонки в field_df: {list(field_df.columns)}")
        
        # Выгрузка в Excel
        excel_buffer = BytesIO()
        field_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        
        st.download_button(
            label="⬇️ Скачать field_df",
            data=excel_buffer,
            file_name="field_df_проверка.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # ПРОВЕРКА ВРЕМЕННАЯ

        
        st.write("### 🎯 Шаг 4: Обработка CXWAY (если есть)")
        
        # Проверяем есть ли файл CXWAY
        cxway_df = st.session_state.uploaded_files.get('cxway')
        cxway_processed = None
        
        if cxway_df is not None:
            with st.spinner("Обрабатываю CXWAY..."):
                # Очищаем и преобразуем CXWAY
                cxway_processed = data_cleaner.clean_cxway(
                    cxway_df, 
                    hierarchy_df, 
                    google_updated
                )

                # 🔴 ВЫГРУЗКА ПРЯМО ЗДЕСЬ
                if cxway_processed is not None and not cxway_processed.empty:
                    excel_buffer = BytesIO()
                    cxway_processed.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)
                    st.download_button(
                        label="📥 СКАЧАТЬ CXWAY ДО ОБЪЕДИНЕНИЯ",
                        data=excel_buffer,
                        file_name="cxway_processed.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="cxway_download"
                    )
                
                if cxway_processed is not None and not cxway_processed.empty:
                    st.success(f"✅ CXWAY обработан: {len(cxway_processed)} полевых проектов")
                    
                    # Объединяем с полевыми проектами из массива
                    if field_df is not None and not field_df.empty:
                        combined_field_projects = data_cleaner.merge_field_projects(
                            field_df, 
                            cxway_processed
                        )
                        
                        # Обновляем поле field_df объединенными данными
                        field_df = combined_field_projects
                        st.session_state.cleaned_data['полевые_проекты'] = field_df
                        st.success(f"✅ Объединено с CXWAY: {len(field_df)} полевых проектов")
              
                        # 🔴 Выгрузка visits_df
                        visits_df = st.session_state.cleaned_data.get('полевые_проекты')
                        if visits_df is not None and not visits_df.empty:
                            excel_buffer = BytesIO()
                            visits_df.to_excel(excel_buffer, index=False)
                            excel_buffer.seek(0)
                            st.download_button(
                                label="📥 СКАЧАТЬ visits_df (полевые_проекты)",
                                data=excel_buffer,
                                file_name=f"visits_df_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                    else:
                        # Если в массиве не было полевых, используем только CXWAY
                        field_df = cxway_processed
                        st.session_state.cleaned_data['полевые_проекты'] = field_df
                        st.success(f"✅ Используем только CXWAY: {len(field_df)} проектов")
                else:
                    st.info("ℹ️ CXWAY не содержит полевых проектов или пуст")
        else:
            st.info("ℹ️ Файл CXWAY не загружен, пропускаем")
        
        st.write("### 🎯 Шаг 5: Создание отчета")
        with st.spinner("Формирую Excel файл..."):
            excel_output = data_cleaner.export_split_array_to_excel(field_df, non_field_df)
            if excel_output:
                st.session_state.excel_files['разделенный_массив'] = excel_output
                st.success("✅ Отчет создан успешно!")
            else:
                st.warning("⚠️ Не удалось создать Excel файл")

        # ============================================
        # 🆕 ИЗВЛЕЧЕНИЕ БАЗОВЫХ ДАННЫХ ДЛЯ ПЛАН/ФАКТА
        # ============================================
        if field_df is not None and not field_df.empty:
            try:
                # Берем полевые_проекты (уже с CXWAY!)
                field_projects_df = st.session_state.cleaned_data.get('полевые_проекты')
                
                st.write("### 🔍 ДИАГНОСТИКА ПОЛЕЙ В field_projects_df")
                st.write("Колонки:", list(field_projects_df.columns))
                st.write("Тип колонки 'Имя клиента':", field_projects_df['Имя клиента'].dtype)
                st.write("Первые 5 значений:", field_projects_df['Имя клиента'].head(5).tolist())

                if field_projects_df is not None and not field_projects_df.empty:
                    base_data = visit_calculator.extract_hierarchical_data(field_projects_df, google_df)
                    st.write(f"🔥 base_data создан, строк: {len(base_data)}")
                    
                    # ========== ВЫГРУЗКА HIERARCHY_DF ==========
                    if not base_data.empty:
                        excel_buffer = BytesIO()
                        base_data.to_excel(excel_buffer, index=False)
                        excel_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 СКАЧАТЬ ИЕРАРХИЮ (hierarchy_df)",
                            data=excel_buffer,
                            file_name=f"hierarchy_df_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_hierarchy"
                        )
                        st.success(f"✅ Иерархия выгружена: {len(base_data)} уникальных цепочек")
                    else:
                        st.warning("⚠️ base_data пустой!")
                    # ============================================
                
                else:
                    base_data = pd.DataFrame()
                    st.warning("⚠️ Нет полевых проектов для построения иерархии")
                
                # 🔴 СОХРАНЯЕМ base_data В visit_report (ЭТО ВНЕ ELSE!)
                if 'visit_report' not in st.session_state:
                    st.session_state.visit_report = {}
                
                st.session_state.visit_report['base_data'] = base_data
                st.session_state.visit_report['timestamp'] = datetime.now().isoformat()
                
            except Exception as e:
                st.warning(f"⚠️ Ошибка извлечения базовых данных: {str(e)[:100]}")
                base_data = pd.DataFrame()


                
                # ========== ПРОВЕРКИ ==========
                # 1. Проверка поля ПО в гугл таблице
                portal_col = 'Портал на котором идет проект (для работы полевой команды)'
                if portal_col in google_df.columns:
                    checker_count = (google_df[portal_col] == 'Чеккер').sum()
                    st.success(f"✅ 1. Поле '{portal_col}' найдено. Проектов на Чеккере: {checker_count}")
                else:
                    st.warning(f"⚠️ 1. Поле '{portal_col}' НЕ найдено в гугл таблице")
                
                # 2. Проверка загрузки в таблицу ПланФакт
                if not base_data.empty:
                    checker_projects_count = len(base_data)
                    st.success(f"✅ 2. Загружено {checker_projects_count} проектов в таблицу ПланФакт")
                    
                    # Проверяем колонку ПО
                    if 'ПО' in base_data.columns:
                        po_values = base_data['ПО'].unique()
                        st.write(f"   Колонка 'ПО' содержит: {list(po_values)}")
                    else:
                        st.warning("   ⚠️ Колонка 'ПО' отсутствует")
                        
                    # Сохраняем данные
                    st.session_state['visit_report'] = {
                        'base_data': base_data,
                        'timestamp': datetime.now().isoformat()
                    }
                else:
                    st.warning("⚠️ 2. Не удалось загрузить проекты в таблицу ПланФакт")
        
        # Показываем статистику
        st.write("### 📊 Статистика после обработки")
        
        cols = st.columns(4)
        with cols[0]:
            total_field = len(field_df) if field_df is not None else 0
            st.metric("Полевые проекты", total_field)
            
        with cols[1]:
            total_non_field = len(non_field_df) if non_field_df is not None else 0
            st.metric("Неполевые проекты", total_non_field)
            
        with cols[2]:
            total_all = total_field + total_non_field
            st.metric("Всего проектов", total_all)
            
        with cols[3]:
            if cxway_processed is not None and not cxway_processed.empty:
                cxway_count = len(cxway_processed)
                source_text = f"Из CXWAY: {cxway_count}"
            else:
                source_text = "CXWAY: нет"
            st.metric("Источник", source_text)
        
        # Дополнительная информация
        if cxway_processed is not None and not cxway_processed.empty:
            field_from_array = total_field - cxway_count
            st.info(f"📊 Детали: {cxway_count} проектов из CXWAY + {field_from_array} из Массива")
        
        return True
        
    except Exception as e:
        st.error(f"❌ Ошибка в process_field_projects_with_stats: {str(e)[:200]}")
        import traceback
        st.error(f"Детали: {traceback.format_exc()[:500]}")
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
    
    # Календарь периода
    st.write("**Период расчета:**")
    today = date.today()
    first_day = date(today.year, today.month, 1)
    yesterday = today - timedelta(days=1)
    
    # Если yesterday раньше first_day (первый день месяца)
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
    
    # Проверка: даты в одном месяце
    if start_date.month != end_date.month:
        st.warning("⚠️ Даты должны быть в одном месяце")
        end_date = start_date.replace(day=28)  # Корректируем
    
    st.info(f"Период: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")

    # Этапы
    
    
    st.markdown("---")
    st.subheader("📊 Коэффициенты этапов")
    
    # Слайдеры для весов этапов с ограничением 0-2
    st.write("**Распределение весов по этапам (0-2):**")
    
    stage_weights = []
    default_weights = [0.8, 1.2, 1.0, 0.9]  # новые дефолтные значения
    
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
    
    # Расчет коэффициентов
    total_weight = sum(stage_weights)
    if total_weight > 0:
        coefficients = [w/total_weight for w in stage_weights]
    else:
        coefficients = [0.25, 0.25, 0.25, 0.25]  # равные если все нули
        st.warning("⚠️ Сумма весов = 0, используем равные коэффициенты")
    
    # Визуализация распределения
    st.write("**Распределение коэффициентов:**")
    for i, coeff in enumerate(coefficients, 1):
        st.progress(coeff, text=f"Этап {i}: {coeff:.1%}")
    
    # Сохраняем в session_state
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
            st.success(f"✅ Портал загружен: {len(portal_df):,} строк")
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
            st.success(f"✅ Автокодификация загружена: {len(autocoding_df):,} строк")
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
            st.success(f"✅ Проекты загружены: {len(projects_df):,} строк")
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
            st.success(f"✅ Иерархия загружена: {len(hierarchy_df):,} строк")
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
            st.success(f"✅ CXWAY загружен: {len(cxway_df):,} строк")
            display_file_preview(cxway_df, "Просмотр данных CXWAY")
    
    # ==============================================
    # СЕКЦИЯ 2: СТАТУС И ОБРАБОТКА
    # ==============================================
    st.markdown("---")
    
    if st.session_state.uploaded_files:
        st.subheader("📊 Статус загрузки")
        
        required_files = ['портал', 'автокодификация', 'сервизория', 'иерархия']
        loaded_count = sum(1 for f in required_files if f in st.session_state.uploaded_files)
        
        if loaded_count == 4:
            st.success(f"🎉 Все 4 файла загружены!")
            
            summary_data = []
            for name in required_files:
                df = st.session_state.uploaded_files[name]
                summary_data.append({
                    'Файл': name,
                    'Строк': f"{len(df):,}",
                    'Колонок': len(df.columns),
                    'Пример': ', '.join(list(df.columns)[:2])
                })
            
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True)
            
            st.markdown("---")
            st.subheader("🚀 Полная обработка данных")
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.info("""
                **Процесс обработки включает:**
                1. Очистку портала (массива)
                2. Очистку проектов (гугл таблицы)
                3. Обогащение массива кодами проектов
                4. Выгрузку результатов в Excel
                """)
            
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
                            st.write("🔍 **1. Проверка файлов...**")
                            missing_files = [f for f in required_files if f not in st.session_state.uploaded_files]
                            if missing_files:
                                raise ValueError(f"Отсутствуют файлы: {', '.join(missing_files)}")
                            st.write("✅ Все файлы проверены")
                            
                            # ЭТАП 2: Очистка портала
                            st.write("🧹 **2. Очистка портала...**")
                            portal_raw = st.session_state.uploaded_files['портал']
                            portal_cleaned, portal_error = process_single_step(
                                data_cleaner.clean_array, "Очистка портала", portal_raw
                            )
                            
                            if portal_error:
                                st.warning(f"⚠️ {portal_error}")
                                portal_cleaned = portal_raw
                            
                            st.session_state.cleaned_data['портал'] = portal_cleaned
                            st.write(f"✅ Очищено: {len(portal_cleaned):,} строк")
                            
                            # ЭТАП 3: Очистка проектов
                            st.write("🧹 **3. Очистка проектов...**")
                            projects_raw = st.session_state.uploaded_files['сервизория']
                            projects_cleaned, projects_error = process_single_step(
                                data_cleaner.clean_google, "Очистка проектов", projects_raw
                            )
                            
                            if projects_error:
                                st.warning(f"⚠️ {projects_error}")
                                projects_cleaned = projects_raw
                            
                            st.session_state.cleaned_data['сервизория'] = projects_cleaned
                            st.write(f"✅ Очищено: {len(projects_cleaned):,} строк")
                            
                            # ЭТАП 4: Обогащение массива
                            st.write("🔗 **4. Обогащение массива...**")
                            if 'портал' in st.session_state.cleaned_data and 'сервизория' in st.session_state.cleaned_data:
                                enriched_result, enrich_error = process_single_step(
                                    data_cleaner.enrich_array_with_project_codes,
                                    "Обогащение массива",
                                    st.session_state.cleaned_data['портал'],
                                    st.session_state.cleaned_data['сервизория']
                                )
                                
                                if enrich_error:
                                    st.warning(f"⚠️ {enrich_error}")
                                    enriched_array = st.session_state.cleaned_data['портал']
                                    discrepancy_df = pd.DataFrame()
                                    stats = {'filled': 0, 'total': 0}
                                else:
                                    enriched_array, discrepancy_df, stats = enriched_result
                                    st.session_state.cleaned_data['портал'] = enriched_array
                                    
                                    if not discrepancy_df.empty:
                                        st.session_state['array_discrepancies'] = discrepancy_df
                                        st.session_state['discrepancy_stats'] = stats
                                
                                st.write(f"✅ Обогащено кодов: {stats.get('filled', 0):,}")
    
                            # ЭТАП 4.5: Добавление ЗОД в массив
                            st.write("👥 **4.5. Добавление ЗОД из справочника...**")
                            if 'портал' in st.session_state.cleaned_data and 'иерархия' in st.session_state.uploaded_files:
                                hierarchy_df = st.session_state.uploaded_files['иерархия']
                                array_with_zod, zod_error = process_single_step(
                                    data_cleaner.add_zod_from_hierarchy,
                                    "Добавление ЗОД",
                                    st.session_state.cleaned_data['портал'],
                                    hierarchy_df
                                )
                                
                                if zod_error:
                                    st.warning(f"⚠️ {zod_error}")
                                elif array_with_zod is not None:
                                    st.session_state.cleaned_data['портал'] = array_with_zod
                                    st.write(f"✅ ЗОД добавлен в массив")
                                    
                            # ЭТАП 5: Разделение на полевые/неполевые проекты
                            st.write("🎯 **5. Разделение на полевые/неполевые проекты...**")
                            
                            field_success = False
                            try:
                                field_success = process_field_projects_with_stats()
                            except Exception as e:
                                st.write(f"⚠️ Ошибка: {str(e)[:100]}")
                            
                            if field_success:
                                st.write("✅ Проекты разделены")
                                if 'разделенный_массив' in st.session_state.excel_files:
                                    st.write("📁 Файл 'разделенный_массив.xlsx' создан")
                            else:
                                st.write("⚠️ Разделение не удалось")
                                
                            # ЭТАП 6: Выгрузка в Excel
                            st.write("📊 **6. Выгрузка в Excel...**")
                            
                            # Массив
                            if 'портал' in st.session_state.cleaned_data:
                                array_with_field = st.session_state.cleaned_data.get('портал_с_полем', st.session_state.cleaned_data['портал'])
                                array_excel, array_export_error = process_single_step(
                                    data_cleaner.export_array_to_excel,
                                    "Выгрузка массива",
                                    array_with_field  # ← передаем массив с колонкой 'Полевой'
                                )
                                
                                if array_excel:
                                    st.session_state.excel_files['массив'] = array_excel
                                    st.write("   ✅ Файл 'очищенный_массив.xlsx' создан")
                                elif array_export_error:
                                    st.write(f"   ⚠️ {array_export_error}")
                            
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
                                    st.write("   ✅ Файл 'очищенные_проекты.xlsx' создан")
                                elif projects_export_error:
                                    st.write(f"   ⚠️ {projects_export_error}")
                            
                            # Сохранение статистики
                            st.session_state.processing_stats = {
                                'timestamp': datetime.now().isoformat(),
                                'portal_rows': len(portal_cleaned),
                                'projects_rows': len(projects_cleaned),
                                'excel_files': len(st.session_state.excel_files),
                                'enriched_codes': stats.get('filled', 0) if 'stats' in locals() else 0
                            }
                            
                            st.success("✅ **Обработка завершена!**")
                            st.session_state.processing_complete = True
                            
                    except ImportError as e:
                        st.error(f"❌ Ошибка импорта модулей: {e}")
                        st.code(traceback.format_exc())
                    except Exception as e:
                        st.error(f"❌ Критическая ошибка обработки: {e}")
                        st.session_state['last_error'] = {
                            'step': 'Общая обработка',
                            'error': str(e),
                            'traceback': traceback.format_exc()
                        }
        else:
            st.warning(f"⚠️ Загружено {loaded_count} из 4 файлов")
            missing = [f for f in required_files if f not in st.session_state.uploaded_files]
            st.write(f"Ожидаются: {', '.join(missing)}")
    
    
    # ==============================================
    # СЕКЦИЯ 3: РЕЗУЛЬТАТЫ
    # ==============================================
    if st.session_state.processing_complete:
        st.markdown("---")
        st.subheader("✅ Результаты обработки")
        
        stats_cols = st.columns(4)
        with stats_cols[0]:
            st.metric("Файлов обработано", len(st.session_state.cleaned_data))
        with stats_cols[1]:
            if 'портал' in st.session_state.cleaned_data:
                st.metric("Строк в массиве", f"{len(st.session_state.cleaned_data['портал']):,}")
        with stats_cols[2]:
            if 'сервизория' in st.session_state.cleaned_data:
                st.metric("Строк в проектах", f"{len(st.session_state.cleaned_data['сервизория']):,}")
        with stats_cols[3]:
            if 'enriched_codes' in st.session_state.processing_stats:
                st.metric("Заполнено кодов", f"{st.session_state.processing_stats['enriched_codes']:,}")
    
        # === ПРОВЕРКА ===
        st.write("**🔍 ПРОВЕРКА:**")
        
        # 1. АК - полевые проекты
        autocoding_df = st.session_state.uploaded_files.get('автокодификация')
        if autocoding_df is not None:
            if 'Направление' in autocoding_df.columns:
                ak_01 = (autocoding_df['Направление'].astype(str).str.strip() == '.01').sum()
                ak_02 = (autocoding_df['Направление'].astype(str).str.strip() == '.02').sum()
                st.write(f"1️⃣ АК: {ak_01 + ak_02} полевых (.01={ak_01}, .02={ak_02})")
                
                # Дополнительная проверка: все направления
                all_directions = autocoding_df['Направление'].astype(str).str.strip().unique()
                st.write(f"   Все направления в АК: {list(all_directions)[:10]}")
            else:
                st.write("1️⃣ АК: нет колонки 'Направление'")
                # Покажем какие колонки есть
                st.write(f"   Колонки в АК: {list(autocoding_df.columns)[:10]}")
        
        # 2. Проверка совпадения кодов
        st.write("**2. Проверка совпадения кодов:**")
        
        # Гугл таблица
        google_df = st.session_state.cleaned_data.get('сервизория')
        if google_df is not None:
            # Ищем колонку с кодом
            google_code_cols = [col for col in google_df.columns if 'код' in str(col).lower()]
            if google_code_cols:
                google_code_col = google_code_cols[0]
                google_codes = google_df[google_code_col].astype(str).str.strip()
                google_codes_valid = google_codes[~google_codes.isin(['', 'nan', 'None'])]
                st.write(f"   Уникальных кодов в гугл: {len(google_codes_valid.unique())}")
            else:
                st.write("   ❌ В гугл нет колонки с 'код'")
        
        # Массив
        array_df = st.session_state.cleaned_data.get('портал_с_полем')
        if array_df is not None:
            # Ищем колонку с кодом
            array_code_cols = [col for col in array_df.columns if 'код' in str(col).lower() and 'анкет' in str(col).lower()]
            if array_code_cols:
                array_code_col = array_code_cols[0]
                array_codes = array_df[array_code_col].astype(str).str.strip()
                array_codes_valid = array_codes[~array_codes.isin(['', 'nan', 'None'])]
                st.write(f"   Уникальных кодов в массиве: {len(array_codes_valid.unique())}")
            else:
                st.write("   ❌ В массиве нет колонки 'Код анкеты'")
                
        # 3. Проверка почему 0 полевых
        st.write("**3. Анализ (почему 0 полевых):**")
        
        if 'автокодификация' in st.session_state.uploaded_files:
            ak_df = st.session_state.uploaded_files['автокодификация']
            
            # Вариант A: В АК нет направлений .01/.02
            if 'Направление' in ak_df.columns:
                dir_01_count = (ak_df['Направление'].astype(str).str.strip() == '.01').sum()
                dir_02_count = (ak_df['Направление'].astype(str).str.strip() == '.02').sum()
                
                if dir_01_count == 0 and dir_02_count == 0:
                    st.write("   ❌ **ВАРИАНТ А:** В АК нет проектов с направлениями .01/.02")
                    # Покажем какие направления есть
                    other_dirs = ak_df['Направление'].astype(str).str.strip().unique()
                    st.write(f"   Есть направления: {list(other_dirs)[:10]}")
                else:
                    st.write(f"   ✅ В АК есть .01/.02: .01={dir_01_count}, .02={dir_02_count}")
            
            # Вариант B: Коды не совпадают
            st.write("   **ВАРИАНТ Б:** Коды проектов не совпадают между источниками")
            st.write("   Проверьте:")
            st.write("   - Коды в АК совпадают с кодами в гугл таблице")
            st.write("   - Коды в массиве совпадают с кодами в гугл таблице")
    
        # === ПРОВЕРКА ===
        
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
                    use_container_width=True,
                    help="Очищенный массив"
                )
        
        with download_cols[1]:
            if 'проекты' in st.session_state.excel_files:
                st.download_button(
                    label="⬇️ Скачать очищенные_проекты.xlsx",
                    data=st.session_state.excel_files['проекты'],
                    file_name="очищенные_проекты.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    help="3 вкладки: Оригинал, Очищенный, Сравнение"
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
                    use_container_width=True,
                    help="3 вкладки: Полевые проекты, Неполевые проекты, Статистика"
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
                            use_container_width=True,
                            help="Только полевые проекты (все записи)"
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
                            use_container_width=True,
                            help="Только неполевые проекты (удалены дубликаты по Код/Клиент/Проект)"
                        )
                        
        # ============================================
        # 🆕 ДАННЫЕ ДЛЯ РАСЧЕТА ПЛАН/ФАКТА
        # ============================================

        # Кнопка расчета план/факт
        if st.button("📊 Рассчитать план/факт", type="primary", use_container_width=True):

                
            # 🔵 ПРОВЕРКА ПЕРЕД РАСЧЕТОМ
            # check_errors = check_required_columns()
            
            # if check_errors:
            #     # Показываем ошибки
            #     st.error("🚫 **Не могу рассчитать:**")
            #     for error in check_errors:
            #         st.write(f"- {error}")
                
            
            # 🔵 ЕСЛИ ОШИБОК НЕТ - ПРОВЕРЯЕМ ДАННЫЕ
            if 'plan_calc_params' in st.session_state and 'visit_report' in st.session_state:
                # 1. ПРОВЕРКА - все ли есть для расчета
                required_keys = ['сервизория', 'портал']
                missing_keys = [k for k in required_keys if k not in st.session_state.cleaned_data]
                
                if missing_keys:
                    st.error(f"❌ Отсутствуют данные: {', '.join(missing_keys)}. Сначала запустите обработку.")
                else:
                    # 2. Получаем все данные
                    if 'base_data' not in st.session_state.visit_report:
                        st.error("❌ Нет базовых данных для расчёта. Сначала запустите обработку.")
                        st.stop() 
    
                    base_data = st.session_state.visit_report['base_data']
                    
                    source_df = st.session_state.cleaned_data['полевые_проекты']
                    params = st.session_state['plan_calc_params']
                    
                    # 3. Считаем план
                    plan_result = visit_calculator.calculate_hierarchical_plan_on_date(
                        base_data,
                        source_df,
                        params
                    )
                    # 🔴 ЗАЩИТА: если план пустой - не идем дальше
                    if plan_result is None or plan_result.empty:
                        st.error("❌ План не посчитался. Дальнейшие расчеты невозможны.")
                        st.stop()  # <- остановка выполнения
                    
                    # ========== ПРОСТАЯ ДИАГНОСТИКА ==========
                    st.write("### 🔍 ДИАГНОСТИКА ПЛАНА")
                    
                    # 1. Проверяем base_data
                    st.write(f"1️⃣ base_data: {len(base_data)} строк в иерархии")
                    
                    # 2. Берем первый проект из иерархии для примера
                    if not base_data.empty:
                        example = base_data.iloc[0]
                        project = example['Проект']
                        wave = example['Волна']
                        region = example['Регион']
                        
                        st.write(f"2️⃣ Пример проекта из иерархии: `{project} | {wave} | {region}`")
                        
                        # 3. Проверяем есть ли такой ключ в source_df
                        match = source_df[
                            (source_df['Код анкеты'] == project) &
                            (source_df['Название проекта'] == wave) &
                            (source_df['Регион short'] == region)
                        ]
                        
                        st.write(f"3️⃣ Совпадение в source_df: **{len(match)} строк**")
                        
                        if len(match) == 0:
                            st.error("❌ КЛЮЧ НЕ НАЙДЕН! Причина:")
                            
                            # Проверяем по отдельности
                            by_project = source_df[source_df['Код анкеты'] == project].shape[0]
                            by_wave = source_df[source_df['Название проекта'] == wave].shape[0]
                            by_region = source_df[source_df['Регион short'] == region].shape[0]
                            
                            st.write(f"   - Проект '{project}' встречается: {by_project} раз")
                            st.write(f"   - Волна '{wave}' встречается: {by_wave} раз")
                            st.write(f"   - Регион '{region}' встречается: {by_region} раз")
                    # ==========================================
                    
                    # ПРОВЕРКА
                    if base_data is not None and not base_data.empty:
                        # Проверяем проекты с пустыми датами
                        projects_no_dates = base_data[
                            base_data['Дата старта'].isna() | 
                            base_data['Дата финиша'].isna() | 
                            (base_data['Длительность'] <= 0)
                        ]
                        
                        if not projects_no_dates.empty:
                            excel_buffer = BytesIO()
                            projects_no_dates.to_excel(excel_buffer, index=False)
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label=f"📥 СКАЧАТЬ ПРОЕКТЫ БЕЗ ДАТ ({len(projects_no_dates)})",
                                data=excel_buffer,
                                file_name="проекты_без_дат.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    # 4. Считаем факт
                    fact_result = visit_calculator.calculate_hierarchical_fact_on_date(
                        plan_result, 
                        source_df, 
                        params
                    )
                    # 5. Рассчитываем метрики
                    final_result = visit_calculator._calculate_metrics(
                        fact_result,  # fact_df (обязательный)
                        params,       # calc_params (опциональный)
                        plan_result   # plan_df (опциональный)
                    )

                    # 6. Сохраняем результат
                    st.session_state['visit_report'] = {
                        'calculated_data': final_result,
                        'hierarchy': base_data,
                        'timestamp': datetime.now().isoformat()
                    }

                    # 7. СРАЗУ ПОКАЗЫВАЕМ РЕЗУЛЬТАТ
                    st.markdown("---")
                    st.subheader("📊 Результаты расчета план/факт")

                    # ========== КНОПКА PLAN_DF ==========   ← ВОТ ЗДЕСЬ!
                    if not final_result.empty:
                        excel_buffer = BytesIO()
                        final_result.to_excel(excel_buffer, index=False)
                        excel_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 СКАЧАТЬ PLAN_DF",
                            data=excel_buffer,
                            file_name=f"plan_df.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    # ====================================
                    
                    if not final_result.empty:
                        # Показать таблицу
                        with st.expander("👀 Просмотреть таблицу ПланФакт", expanded=True):
                            st.dataframe(final_result, use_container_width=True, height=300)
                        
                        # Кнопка выгрузки
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
                            use_container_width=True,
                            help="Таблица План/Факт в формате Excel"
                        )
                    else:
                        st.warning("Нет данных для отображения")
                    
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
                
                # Экспорт в Excel
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
        
        # Просмотр данных
        st.markdown("---")
        st.subheader("🔍 Просмотр данных")
        
        # # ПРОВЕРКА ВРЕМЕННАЯ
        # # Быстрая проверка готовности
        # st.markdown("### 🔍 Проверка готовности к расчёту")
        
        # check_ready = check_required_columns()
        
        # if check_ready:
        #     st.error("❌ **Есть проблемы:**")
        #     for problem in check_ready:
        #         st.write(f"- {problem}")
        # else:
        #     st.success("✅ **Всё готово к расчёту!**")
        #     st.write("Настройте параметры в сайдбаре и нажмите кнопку ниже")
        # # ПРОВЕРКА ВРЕМЕННАЯ

    
# ==============================================
# ВТОРАЯ ВКЛАДКА: ОТЧЕТЫ
# ==============================================
with tab2:
    st.title("📈 Отчеты по полевым визитам")
    
    if ('visit_report' not in st.session_state or 
        st.session_state.visit_report.get('calculated_data') is None):
            
        st.warning("⚠️ Нет рассчитанных данных")
        st.info("Сначала запустите расчет на странице 'Загрузка данных'")
    else:
        tab1, tab2 = st.tabs(["📊 ПланФакт на дату", "📈 Другие"])
        
        with tab1:
            data = st.session_state.visit_report['calculated_data']
            hierarchy_df = st.session_state.uploaded_files.get('иерархия')
            dataviz.create_planfact_tab(data, hierarchy_df)
        
        with tab2:
            st.info("Другие отчеты в разработке")





























