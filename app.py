import streamlit as st
from utils.data_loader import loader

st.set_page_config(page_title="ИУ Аудиты", layout="wide")
st.title("📊 ИУ Аудиты - План/Факт")

if st.button("📥 Загрузить данные"):
    portal_data = loader.load_portal()
    if portal_data is not None:
        st.write(f"Загружено: {len(portal_data)} строк")
        st.dataframe(portal_data.head())
