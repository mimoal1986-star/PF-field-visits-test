# utils/data_loader.py
import pandas as pd
import streamlit as st
from config import EXCEL_PATHS


class ExcelDataLoader:
    """
    Класс для загрузки всех исходных Excel файлов
    с кэшированием для скорости работы
    """

    @st.cache_data(show_spinner="📥 Загружаю Портал (Массив.xlsx)...", ttl=3600)
    def load_portal(_self):
        """
        Загружает данные портала (CRM) из Массив.xlsx
        """
        try:
            df = pd.read_excel(
                EXCEL_PATHS['портал'],
                dtype=str,  # Все как текст для начала
                engine='openpyxl'
            )
            st.success(f"✅ Портал загружен: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки портала: {e}")
            return None

    @st.cache_data(show_spinner="📥 Загружаю Автокодификацию...", ttl=3600)
    def load_autocoding(_self):
        """
        Загружает справочник проектов из Автокодификация.xlsx
        """
        try:
            df = pd.read_excel(
                EXCEL_PATHS['автокодификация'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Автокодификация загружена: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки автокодификации: {e}")
            return None

    @st.cache_data(show_spinner="📥 Загружаю Проекты Сервизория...", ttl=3600)
    def load_service_projects(_self):
        """
        Загружает даты волн проектов из Гугл таблица.xlsx
        """
        try:
            df = pd.read_excel(
                EXCEL_PATHS['сервизория'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Проекты Сервизория загружены: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки проектов сервизория: {e}")
            return None

    @st.cache_data(show_spinner="📥 Загружаю иерархию ЗОД-АСС...", ttl=3600)
    def load_hierarchy(_self):
        """
        Загружает иерархию руководителей из ЗОД+АСС.xlsx
        """
        try:
            df = pd.read_excel(
                EXCEL_PATHS['иерархия'],
                dtype=str,
                engine='openpyxl'
            )
            st.success(f"✅ Иерархия ЗОД-АСС загружена: {len(df)} строк, {len(df.columns)} колонок")
            return df
        except Exception as e:
            st.error(f"❌ Ошибка загрузки иерархии: {e}")
            return None

    def load_all_data(_self):
        """
        Загружает все исходные данные и возвращает словарь с DataFrame
        """
        st.info("📂 Начинаю загрузку всех исходных данных...")

        data = {
            'портал': self.load_portal(),
            'автокодификация': self.load_autocoding(),
            'сервизория': self.load_service_projects(),
            'иерархия': self.load_hierarchy()
        }

        # Проверяем успешность загрузки
        success = all(value is not None for value in data.values())

        if success:
            st.success("🎉 Все исходные данные успешно загружены!")
            return data
        else:
            st.error("⚠️ Не удалось загрузить некоторые файлы. Проверьте:")
            st.write("1. Файлы в папке data/raw/")
            st.write("2. Названия файлов в config.py")
            st.write("3. Что файлы не открыты в Excel")
            return None

    def get_data_summary(_self, data):
        """
        Возвращает сводную информацию о загруженных данных
        """
        if not data:
            return None

        summary = []
        for name, df in data.items():
            if df is not None:
                summary.append({
                    'Источник': name,
                    'Строк': len(df),
                    'Колонок': len(df.columns),
                    'Колонки (первые 5)': ', '.join(list(df.columns)[:5])
                })
        
        return pd.DataFrame(summary)


# Создаем глобальный экземпляр загрузчика
data_loader = ExcelDataLoader()
