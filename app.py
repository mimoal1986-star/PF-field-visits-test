# app.py
import streamlit as st
import sys
import os

# Настройка путей
current_dir = os.path.dirname(os.path.abspath(__file__))
utils_path = os.path.join(current_dir, 'utils')
if utils_path not in sys.path:
    sys.path.append(utils_path)

# Импорт загрузчика
try:
    from data_loader import data_loader
    IMPORT_SUCCESS = True
except ImportError as e:
    st.error(f"❌ Ошибка импорта: {e}")
    IMPORT_SUCCESS = False
    data_loader = None

# Настройка страницы
st.set_page_config(
    page_title="ИУ Аудиты - ПланФакт",
    page_icon="📊",
    layout="wide"
)

st.title("📊 ИУ Аудиты - Мониторинг План/Факт")
st.markdown("---")

if not IMPORT_SUCCESS:
    st.error("""
    ⚠️ Не удалось загрузить модуль data_loader.
    
    **Проверьте:**
    1. Файл `utils/data_loader.py` существует
    2. Файл `utils/__init__.py` существует
    3. Структура проекта правильная
    """)
else:
    # Кнопка для загрузки данных
    if st.button("📥 Загрузить все исходные данные", type="primary"):
        with st.spinner("Загружаю данные..."):
            all_data = data_loader.load_all_data()
            
            if all_data:
                st.success("✅ Данные успешно загружены!")
                
                # Показываем сводку
                st.subheader("📊 Сводка по данным")
                for name, df in all_data.items():
                    if df is not None:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric(f"{name} - строк", len(df))
                        with col2:
                            st.metric(f"{name} - колонок", len(df.columns))
                
                # Сохраняем в session_state
                st.session_state['raw_data'] = all_data
    else:
        st.info("👆 Нажмите кнопку выше чтобы загрузить исходные данные")
