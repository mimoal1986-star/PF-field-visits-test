# utils/dataviz.py
import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO

class DataVisualizer:
    
    def create_project_summary(self, df):
        """
        Агрегация данных по проектам
        Одна строка = один проект
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        # Проверяем наличие колонки 'Проект'
        project_col = 'Проект'
        if project_col not in df.columns:
            st.error(f"❌ В данных нет колонки '{project_col}'")
            return pd.DataFrame()
        
        # Список колонок для агрегации (только существующие)
        agg_columns = {
            'План проекта, шт.': 'sum',
            'План на дату, шт.': 'sum',
            'Факт на дату, шт.': 'sum',
            'Длительность': 'first',
            'Дата старта': 'first',
            'Дата финиша': 'first',
            'Клиент': 'first',
            'ПО': 'first'
        }
        
        # Оставляем только те колонки, которые есть в DataFrame
        existing_agg = {k: v for k, v in agg_columns.items() if k in df.columns}
        
        # Группируем по проекту
        project_agg = df.groupby(project_col).agg(existing_agg).reset_index()
        
        # Рассчитываем метрики
        project_agg['Исполнение проекта,%'] = 0.0
        mask_plan = project_agg['План на дату, шт.'] > 0
        if mask_plan.any():
            project_agg.loc[mask_plan, 'Исполнение проекта,%'] = (
                project_agg.loc[mask_plan, 'Факт на дату, шт.'] / 
                project_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        project_agg['△План/Факт, шт.'] = (
            project_agg['План на дату, шт.'] - project_agg['Факт на дату, шт.']
        ).round(1)
        
        project_agg['△План/Факт,%'] = 0.0
        if mask_plan.any():
            project_agg.loc[mask_plan, '△План/Факт,%'] = (
                project_agg.loc[mask_plan, '△План/Факт, шт.'] / 
                project_agg.loc[mask_plan, 'План на дату, шт.'] * 100
            ).round(1)
        
        # Дней до конца проекта
        today = pd.Timestamp.now()
        project_agg['Дней до конца проекта'] = 0
        
        if 'Дата финиша' in project_agg.columns:
            mask_finish = project_agg['Дата финиша'].notna()
            if mask_finish.any():
                days_left = (project_agg.loc[mask_finish, 'Дата финиша'] - today).dt.days
                project_agg.loc[mask_finish, 'Дней до конца проекта'] = days_left.clip(lower=0)
        
        # Сортируем по исполнению
        project_agg = project_agg.sort_values('Исполнение проекта,%', ascending=False)
        
        return project_agg
    
    def create_planfact_tab(self, data, hierarchy_df=None):
        """Создает вкладку ПланФакт на дату"""
        if data is None or data.empty:
            st.warning("⚠️ Нет данных для отчета")
            return
        
        st.subheader("📊 Сводка по проектам")
        
        # Агрегируем по проектам
        project_data = self.create_project_summary(data)
        
        if project_data.empty:
            st.warning("⚠️ Нет данных после агрегации")
            return
        
        # Колонки для отображения
        display_columns = [
            'Проект',
            'Клиент',
            'ПО',
            'Исполнение проекта,%',
            'План на дату, шт.',
            'Факт на дату, шт.',
            '△План/Факт, шт.',
            '△План/Факт,%',
            'План проекта, шт.',
            'Дней до конца проекта',
            'Длительность'
        ]
        
        # Только существующие колонки
        existing_cols = [col for col in display_columns if col in project_data.columns]
        df_display = project_data[existing_cols]
        
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
            st.metric("🎯 Выполнение плана", f"{pf_percent:.1f}%")
        
        # Таблица
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True
        )
        
        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_display.to_excel(writer, sheet_name='План_факт_проекты', index=False)
        
        st.download_button(
            label="⬇️ Скачать Excel",
            data=output.getvalue(),
            file_name=f"план_факт_проекты_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# Глобальный экземпляр
dataviz = DataVisualizer()



