# app.py
import streamlit as st
import pandas as pd
import sys
import os

# Настройка путей для импорта
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

st.title("📤 Загрузка исходных данных")
st.markdown("Загрузите 4 Excel файла для формирования отчетов")

# Инициализация session_state
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = {}
if 'cleaned_data' not in st.session_state:
    st.session_state.cleaned_data = {}

# ==============================================
# СЕКЦИЯ 1: ЗАГРУЗКА ФАЙЛОВ
# ==============================================

# 1. ЗАГРУЗКА ПОРТАЛА
st.subheader("1. 📋 Портал (Массив.xlsx)")
portal_file = st.file_uploader(
    "Загрузите файл Массив.xlsx с данными портала",
    type=['xlsx', 'xls'],
    key="portal"
)

if portal_file is not None:
    try:
        portal_df = pd.read_excel(portal_file, dtype=str)
        st.session_state.uploaded_files['портал'] = portal_df
        st.success(f"✅ Портал загружен: {len(portal_df)} строк, {len(portal_df.columns)} колонок")
        
        with st.expander("👀 Просмотр данных портала"):
            st.dataframe(portal_df.head(10))
    except Exception as e:
        st.error(f"❌ Ошибка загрузки: {e}")

# 2. ЗАГРРУЗКА АВТОКОДИФИКАЦИИ
st.subheader("2. 🏷️ Автокодификация")
autocoding_file = st.file_uploader(
    "Загрузите файл Автокодификация.xlsx",
    type=['xlsx', 'xls'],
    key="autocoding"
)

if autocoding_file is not None:
    try:
        autocoding_df = pd.read_excel(autocoding_file, dtype=str)
        st.session_state.uploaded_files['автокодификация'] = autocoding_df
        st.success(f"✅ Автокодификация загружена: {len(autocoding_df)} строк")
        
        with st.expander("👀 Просмотр автокодификации"):
            st.dataframe(autocoding_df.head(10))
    except Exception as e:
        st.error(f"❌ Ошибка загрузки: {e}")

# 3. ЗАГРУЗКА ПРОЕКТОВ СЕРВИЗОРИЯ
st.subheader("3. 📅 Проекты Сервизория")
projects_file = st.file_uploader(
    "Загрузите файл Гугл таблица.xlsx",
    type=['xlsx', 'xls'],
    key="projects"
)

if projects_file is not None:
    try:
        projects_df = pd.read_excel(projects_file, dtype=str)
        st.session_state.uploaded_files['сервизория'] = projects_df
        st.success(f"✅ Проекты загружены: {len(projects_df)} строк")
        
        with st.expander("👀 Просмотр проектов"):
            st.dataframe(projects_df.head(10))
    except Exception as e:
        st.error(f"❌ Ошибка загрузки: {e}")

# 4. ЗАГРУЗКА ИЕРАРХИИ
st.subheader("4. 👥 Иерархия ЗОД-АСС")
hierarchy_file = st.file_uploader(
    "Загрузите файл ЗОД+АСС.xlsx",
    type=['xlsx', 'xls'],
    key="hierarchy"
)

if hierarchy_file is not None:
    try:
        hierarchy_df = pd.read_excel(hierarchy_file, dtype=str)
        st.session_state.uploaded_files['иерархия'] = hierarchy_df
        st.success(f"✅ Иерархия загружена: {len(hierarchy_df)} строк")
        
        with st.expander("👀 Просмотр иерархии"):
            st.dataframe(hierarchy_df.head(10))
    except Exception as e:
        st.error(f"❌ Ошибка загрузки: {e}")

# ==============================================
# СЕКЦИЯ 2: СТАТУС ЗАГРУЗКИ
# ==============================================
st.markdown("---")
st.subheader("📊 Статус загрузки")

if len(st.session_state.uploaded_files) == 4:
    st.success("🎉 Все 4 файла успешно загружены!")
    
    # Сводная таблица
    st.subheader("Сводка по загруженным данным")
    
    summary_data = []
    for name, df in st.session_state.uploaded_files.items():
        summary_data.append({
            'Файл': name,
            'Строк': len(df),
            'Колонок': len(df.columns),
            'Пример колонок': ', '.join(list(df.columns)[:3])
        })
    
    st.dataframe(pd.DataFrame(summary_data))
    
    # Временная кнопка без перехода
    if st.button("🚀 Начать обработку данных", type="primary"):
        st.info("Обработка данных будет реализована в следующем шаге")
        
else:
    st.warning(f"⚠️ Загружено {len(st.session_state.uploaded_files)} из 4 файлов")
    missing = [f for f in ['портал', 'автокодификация', 'сервизория', 'иерархия'] 
               if f not in st.session_state.uploaded_files]
    st.write(f"Ожидаются: {', '.join(missing)}")

# ==============================================
# СЕКЦИЯ 3: ТЕСТИРОВАНИЕ ОЧИСТКИ ДАННЫХ
# ==============================================
if len(st.session_state.uploaded_files) > 0:
    st.markdown("---")
    st.subheader("🧹 Тестирование очистки данных")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🧪 Протестировать очистку Гугл таблицы", type="secondary"):
            if 'сервизория' in st.session_state.uploaded_files:
                try:
                    from data_cleaner import data_cleaner
                    
                    google_df = st.session_state.uploaded_files['сервизория']
                    
                    with st.spinner("Очищаю Гугл таблицу..."):
                        cleaned_google = data_cleaner.clean_google(google_df)
                        
                        if cleaned_google is not None:
                            st.session_state.cleaned_data['сервизория'] = cleaned_google
                            st.success("✅ Очистка завершена! Данные сохранены в session_state")
                            
                            # НОВОЕ: Кнопка для выгрузки в Excel
                            st.markdown("---")
                            st.subheader("📥 Выгрузка для сверки")
                            
                            # Создаем Excel файл для сравнения
                            excel_file = data_cleaner.export_to_excel(
                                google_df, 
                                cleaned_google,
                                filename="очищенная_гугл_таблица"
                            )
                            
                            if excel_file:
                                st.download_button(
                                    label="⬇️ Скачать Excel с сравнением",
                                    data=excel_file,
                                    file_name="очищенная_гугл_таблица.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    help="Файл содержит 3 вкладки: ОРИГИНАЛ, ОЧИЩЕННЫЙ, СРАВНЕНИЕ"
                                )
                                
                                st.info("""
                                **Файл содержит 3 вкладки:**
                                1. 📋 **ОРИГИНАЛ** - исходные данные
                                2. ✅ **ОЧИЩЕННЫЙ** - после всех преобразований
                                3. 🔍 **СРАВНЕНИЕ** - статистика изменений
                                """)

# ==============================================
# СЕКЦИЯ 4: ИНФОРМАЦИЯ ОБ ОЧИЩЕННЫХ ДАННЫХ
# ==============================================
if st.session_state.cleaned_data:
    st.markdown("---")
    st.subheader("✅ Очищенные данные")
    
    for name, df in st.session_state.cleaned_data.items():
        with st.expander(f"📁 {name} (очищенный)"):
            st.write(f"Размер: {len(df)} строк × {len(df.columns)} колонок")
            
            # Показываем основные колонки
            st.write("**Колонки:**")
            cols_per_row = 4
            columns = list(df.columns)
            
            for i in range(0, len(columns), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(columns[i:i+cols_per_row]):
                    with cols[j]:
                        st.code(col)
            
            # Показываем предпросмотр
            st.write("**Предпросмотр:**")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Кнопка для скачивания очищенных данных
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"⬇️ Скачать очищенный {name} (CSV)",
                data=csv,
                file_name=f"очищенный_{name}.csv",
                mime="text/csv",
                key=f"download_{name}"
            )

# ==============================================
# СЕКЦИЯ 5: ИНФОРМАЦИЯ О ПРОЕКТЕ
# ==============================================
with st.sidebar:
    st.header("ℹ️ Информация о проекте")
    
    st.markdown("**Загружено файлов:**")
    st.write(f"📊 {len(st.session_state.uploaded_files)} из 4")
    
    if st.session_state.uploaded_files:
        st.markdown("**Статистика:**")
        for name, df in st.session_state.uploaded_files.items():
            st.write(f"- {name}: {len(df)} строк")
    
    st.markdown("---")
    
    st.markdown("**Очищено файлов:**")
    st.write(f"🧹 {len(st.session_state.cleaned_data)}")
    
    if st.session_state.cleaned_data:
        st.markdown("**Очищенные данные:**")
        for name, df in st.session_state.cleaned_data.items():
            st.write(f"- {name}: {len(df)} строк")
    
    st.markdown("---")
    
    # Кнопка для сброса всех данных
    if st.button("🔄 Сбросить все данные", type="secondary", use_container_width=True):
        st.session_state.uploaded_files.clear()
        st.session_state.cleaned_data.clear()
        st.rerun()
    
    st.markdown("---")
    
    st.markdown("**Следующие шаги:**")
    st.write("1. 🧹 Очистка всех файлов")
    st.write("2. 🔗 Объединение данных")
    st.write("3. 📊 Создание отчетов")

# ==============================================
# СЕКЦИЯ 6: ДЕБАГ ИНФОРМАЦИЯ (для разработки)
# ==============================================
with st.expander("🐛 Дебаг информация (только для разработки)"):
    st.write("**Session state keys:**")
    st.write(list(st.session_state.keys()))
    
    st.write("**Загруженные файлы:**")
    for key in st.session_state.uploaded_files:
        st.write(f"- {key}")
    
    st.write("**Очищенные файлы:**")
    for key in st.session_state.cleaned_data:
        st.write(f"- {key}")

