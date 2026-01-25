# app.py
import streamlit as st
import pandas as pd
import sys
import os
import traceback
from datetime import datetime
from io import BytesIO

# Настройка путей
current_dir = os.path.dirname(os.path.abspath(__file__))
utils_path = os.path.join(current_dir, 'utils')
if utils_path not in sys.path:
    sys.path.append(utils_path)

# Настройка страницы
st.set_page_config(
    page_title="ИУ Аудиты - Загрузка данных",
    page_icon="📤",
    layout="wide"
)

# Инициализация session_state
DEFAULT_STATE = {
    'uploaded_files': {},
    'cleaned_data': {},
    'excel_files': {},
    'processing_complete': False,
    'processing_stats': {},
    'last_error': None
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

def create_status_container():
    """Создание контейнера для отображения статуса"""
    return st.status("📊 **Подготовка к обработке...**", expanded=True)

# Основной интерфейс
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
                    
                    with create_status_container() as status:
                        # ЭТАП 1: Проверка
                        status.write("🔍 **1. Проверка файлов...**")
                        missing_files = [f for f in required_files if f not in st.session_state.uploaded_files]
                        if missing_files:
                            raise ValueError(f"Отсутствуют файлы: {', '.join(missing_files)}")
                        status.write("✅ Все файлы проверены")
                        
                        # ЭТАП 2: Очистка портала
                        status.write("🧹 **2. Очистка портала...**")
                        portal_raw = st.session_state.uploaded_files['портал']
                        portal_cleaned, portal_error = process_single_step(
                            data_cleaner.clean_array, "Очистка портала", portal_raw
                        )
                        
                        if portal_error:
                            st.warning(f"⚠️ {portal_error}")
                            portal_cleaned = portal_raw
                        
                        st.session_state.cleaned_data['портал'] = portal_cleaned
                        status.write(f"✅ Очищено: {len(portal_cleaned):,} строк")
                        
                        # ЭТАП 3: Очистка проектов
                        status.write("🧹 **3. Очистка проектов...**")
                        projects_raw = st.session_state.uploaded_files['сервизория']
                        projects_cleaned, projects_error = process_single_step(
                            data_cleaner.clean_google, "Очистка проектов", projects_raw
                        )
                        
                        if projects_error:
                            st.warning(f"⚠️ {projects_error}")
                            projects_cleaned = projects_raw
                        
                        st.session_state.cleaned_data['сервизория'] = projects_cleaned
                        status.write(f"✅ Очищено: {len(projects_cleaned):,} строк")
                        
                        # ЭТАП 4: Обогащение массива
                        status.write("🔗 **4. Обогащение массива...**")
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
                            
                            status.write(f"✅ Обогащено кодов: {stats.get('filled', 0):,}")
                        
                        # ЭТАП 5: Выгрузка в Excel
                        status.write("📊 **5. Выгрузка в Excel...**")
                        
                        # Массив
                        if 'портал' in st.session_state.cleaned_data:
                            array_excel, array_export_error = process_single_step(
                                data_cleaner.export_array_to_excel,
                                "Выгрузка массива",
                                st.session_state.cleaned_data['портал']
                            )
                            
                            if array_excel:
                                st.session_state.excel_files['массив'] = array_excel
                                status.write("   ✅ Файл 'очищенный_массив.xlsx' создан")
                            elif array_export_error:
                                status.write(f"   ⚠️ {array_export_error}")
                        
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
                                status.write("   ✅ Файл 'очищенные_проекты.xlsx' создан")
                            elif projects_export_error:
                                status.write(f"   ⚠️ {projects_export_error}")
                        
                        # Сохранение статистики
                        st.session_state.processing_stats = {
                            'timestamp': datetime.now().isoformat(),
                            'portal_rows': len(portal_cleaned),
                            'projects_rows': len(projects_cleaned),
                            'excel_files': len(st.session_state.excel_files),
                            'enriched_codes': stats.get('filled', 0) if 'stats' in locals() else 0
                        }
                        
                        status.update(label="✅ **Обработка завершена!**", state="complete")
                        st.session_state.processing_complete = True
                        st.rerun()
                        
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
                help="3 вкладки: Очищенный массив, Строки с Н Д, Нули в датах"
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
    
    # Просмотр данных
    st.markdown("---")
    st.subheader("🔍 Просмотр данных")
    
    if st.session_state.cleaned_data:
        selected_key = st.selectbox(
            "Выберите набор данных для просмотра",
            options=list(st.session_state.cleaned_data.keys()),
            format_func=lambda x: {
                'портал': '📊 Очищенный массив',
                'сервизория': '📅 Очищенные проекты',
                'автокодификация': '🏷️ Автокодификация',
                'иерархия': '👥 Иерархия'
            }.get(x, x.capitalize())
        )
        
        if selected_key in st.session_state.cleaned_data:
            df = st.session_state.cleaned_data[selected_key]
            st.dataframe(df, use_container_width=True, height=400)
            st.caption(f"Всего: {len(df):,} строк × {len(df.columns)} колонок")
    
    # Действия
    st.markdown("---")
    action_cols = st.columns(3)
    with action_cols[0]:
        if st.button("🔄 Обработать заново", use_container_width=True):
            st.session_state.processing_complete = False
            st.session_state.excel_files.clear()
            st.rerun()
    
    with action_cols[1]:
        if st.button("📋 Экспорт сводки", use_container_width=True):
            summary_df = pd.DataFrame([{
                'Этап': 'Обработка данных',
                'Статус': 'Завершено',
                'Время': st.session_state.processing_stats.get('timestamp', 'N/A'),
                'Файлов': len(st.session_state.cleaned_data),
                'Excel файлов': len(st.session_state.excel_files)
            }])
            st.download_button(
                label="⬇️ Скачать сводку",
                data=summary_df.to_csv(index=False).encode('utf-8'),
                file_name="сводка_обработки.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with action_cols[2]:
        if st.session_state.get('last_error'):
            if st.button("🐛 Детали ошибки", use_container_width=True):
                st.error(f"Последняя ошибка: {st.session_state['last_error'].get('step')}")
                st.code(st.session_state['last_error'].get('error', 'Нет информации'))

# ==============================================
# САЙДБАР
# ==============================================
with st.sidebar:
    st.header("ℹ️ Информация")
    
    st.metric("Загружено файлов", len(st.session_state.uploaded_files))
    
    if st.session_state.uploaded_files:
        with st.expander("📁 Детали файлов"):
            for name, df in st.session_state.uploaded_files.items():
                st.write(f"**{name}**: {len(df):,} строк")
    
    st.markdown("---")
    
    if st.session_state.processing_complete:
        st.success("✅ Обработка завершена")
        st.metric("Создано Excel", len(st.session_state.excel_files))
    else:
        st.info("⏳ Ожидание обработки")
    
    st.markdown("---")
    
    if st.button("🗑️ Сбросить все данные", type="secondary", use_container_width=True):
        for key in list(DEFAULT_STATE.keys()):
            st.session_state[key] = DEFAULT_STATE[key]
        st.rerun()
    
    st.markdown("---")
    
    if st.session_state.get('processing_stats'):
        with st.expander("📈 Статистика обработки"):
            stats = st.session_state.processing_stats
            for key, value in stats.items():
                if key != 'timestamp':
                    st.write(f"**{key.replace('_', ' ').title()}**: {value}")



