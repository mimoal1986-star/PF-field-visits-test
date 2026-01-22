import streamlit as st
import pandas as pd
import io

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

# 2. ЗАГРУЗКА АВТОКОДИФИКАЦИИ
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

# ПРОВЕРКА ВСЕХ ФАЙЛОВ
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
    
    # Кнопка для перехода к обработке
    if st.button("🚀 Перейти к обработке данных", type="primary"):
        st.success("Дашборд будет скоро добавлен!")
        
else:
    st.warning(f"⚠️ Загружено {len(st.session_state.uploaded_files)} из 4 файлов")
    missing = [f for f in ['портал', 'автокодификация', 'сервизория', 'иерархия'] 
               if f not in st.session_state.uploaded_files]
    st.write(f"Ожидаются: {', '.join(missing)}")

