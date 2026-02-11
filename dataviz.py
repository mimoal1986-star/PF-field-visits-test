# utils/dataviz.py
# draft 2.0
import pandas as pd
import streamlit as st

class DataVisualizer:
    
    def create_planfact_tab(self, data, hierarchy_df=None):
        """Создает вкладку ПланФакт на дату"""
        if data is None or data.empty:
            st.warning("Нет данных для отчета")
            return
        
        # Копия для модификаций
        df = data.copy()
                
        # Переименовываем существующие колонки
        df = df.rename(columns={
            'Утилизация тайминга, %': 'Потребление времени, %'
        })
        
        # Выбираем нужные колонки в правильном порядке
        display_columns = [
            'Имя клиента', 'Название проекта', 'Код проекта', 'ПО','DSM','ASM','RS','Регион',
            'Исполнение Проекта,%', 'Фокус',
            'План на дату, шт.', 'Факт на дату, шт.', '△План/Факт на дату, шт.', '%ПФ на дату',
            'План проекта, шт.', 'Прогноз на месяц, шт.',
            'Дней до конца проекта', 'Ср. план на день для 100% плана, шт.'
        ]
        
        # Оставляем только существующие
        existing_cols = [col for col in display_columns if col in df.columns]
        df_display = df[existing_cols]
        
        # KPI сверху
        col1, col2, col3 = st.columns(3)
        with col1:
            plan_total = df_display['План на дату, шт.'].sum() if 'План на дату, шт.' in df_display.columns else 0
            st.metric("📊 План на дату", f"{plan_total:,.0f} шт")
        
        with col2:
            fact_total = df_display['Факт на дату, шт.'].sum() if 'Факт на дату, шт.' in df_display.columns else 0
            st.metric("✅ Факт на дату", f"{fact_total:,.0f} шт")
        
        with col3:
            pf_percent = (fact_total / plan_total * 100) if plan_total > 0 else 0
            st.metric("🎯 План/Факт %", f"{pf_percent:.1f}%")
        
        # Отображаем таблицу
        st.markdown("""
        <style>
            div[data-testid="stDataFrame"] {
                font-size: 12px !important;
            }
            div[data-testid="stDataFrame"] table {
                table-layout: fixed !important;
                width: 100% !important;
            }
            div[data-testid="stDataFrame"] th, 
            div[data-testid="stDataFrame"] td {
                padding: 4px 6px !important;
                word-wrap: break-word !important;
                white-space: normal !important;
                max-width: 150px !important;
            }
        </style>
        """, unsafe_allow_html=True)
        
        table_height = min(800, max(300, 35 * len(df_display) + 50))
        st.dataframe(
            df_display,
            use_container_width=True,
            height=table_height,
            hide_index=True
        )

# Глобальный экземпляр
dataviz = DataVisualizer()
