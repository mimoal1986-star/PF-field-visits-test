import pandas as pd
import streamlit as st
from config import EXCEL_PATHS

class ExcelDataLoader:
    """
    Загрузчик Excel файлов для системы План-Факт
    """
    
    @st.cache_data(show_spinner="📥 Загрузка Портал данных...", ttl=3600)
    def load_portal(_self):
        """Загружает CRM данные из Массив.xlsx"""
        try:
            df = pd.read_excel(
                EXCEL_PATHS['портал'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Портал: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки портала: {e}")
            return None
    
    @st.cache_data(show_spinner="📥 Загрузка Автокодификации...", ttl=3600)
    def load_autocoding(_self):
        """Загружает справочник проектов"""
        try:
            df = pd.read_excel(
                EXCEL_PATHS['автокодификация'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Автокодификация: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки автокодификации: {e}")
            return None
    
    @st.cache_data(show_spinner="📥 Загрузка Проектов Сервизория...", ttl=3600)
    def load_service_projects(_self):
        """Загружает даты волн проектов"""
        try:
            df = pd.read_excel(
                EXCEL_PATHS['сервизория'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Проекты Сервизория: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки проектов: {e}")
            return None
    
    @st.cache_data(show_spinner="📥 Загрузка иерархии ЗОД-АСС...", ttl=3600)
    def load_hierarchy(_self):
        """Загружает иерархию руководителей"""
        try:
            df = pd.read_excel(
                EXCEL_PATHS['иерархия'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Иерархия: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки иерархии: {e}")
            return None
    
    def load_all_data(_self):
        """
        Загружает все исходные данные
        Возвращает словарь с DataFrame
        """
        st.info("📂 Загружаю все исходные данные...")
        
        data = {
            'портал': self.load_portal(),
            'автокодификация': self.load_autocoding(),
            'сервизория': self.load_service_projects(),
            'иерархия': self.load_hierarchy()
        }
        
        # Проверяем успешность загрузки
        success = all(value is not None for value in data.values())
        
        if success:
            st.success("✅ Все данные успешно загружены!")
            return data
        else:
            st.error("❌ Не удалось загрузить некоторые файлы")
            return None

# Создаем глобальный экземпляр
data_loader = ExcelDataLoader()