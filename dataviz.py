# utils/dataviz.py
# draft 2.1
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
        
        # 2. Выбираем нужные колонки в правильном порядке
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
        
        # 3. Фильтры
        # st.sidebar.header("Фильтры")
        
        
        # 4. KPI сверху
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
        
        # 5. Добавляем строку Итого
        if not df_display.empty:
            total_row = self._calculate_totals(df_display)
            df_with_totals = pd.concat([df_display, total_row], ignore_index=True)
            
            # 6. Отображаем таблицу
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
            
            table_height = min(800, max(300, 35 * len(df_with_totals) + 50))
            st.dataframe(
                df_with_totals,
                use_container_width=True,
                height=table_height,
                hide_index=True
            )


    def _calculate_totals(self, df):
        """Создает строку Итого"""
        total_row = {}
        
        # Колонки для суммирования
        sum_columns = [
            'План на дату, шт.', 'Факт на дату, шт.', '△План/Факт на дату, шт.',
            'План проекта, шт.', 'Прогноз на месяц, шт.',
            'Кол-во визитов до 100% плана, шт.', 'Поручено'
        ]
        
        for col in df.columns:
            if col in sum_columns:
                total_row[col] = df[col].sum()
            elif col == '%ПФ на дату':
                # Для % считаем средневзвешенное по плану
                if 'План на дату, шт.' in df.columns:
                    plan = df['План на дату, шт.'].sum()
                    if plan > 0:
                        weighted_sum = (df['%ПФ на дату'] * df['План на дату, шт.']).sum()
                        total_row[col] = weighted_sum / plan
                    else:
                        total_row[col] = 0
                else:
                    total_row[col] = 0
                    
            elif col == 'Прогноз на месяц, %':
                # Для % считаем средневзвешенное по плану проекта
                if 'План проекта, шт.' in df.columns:
                    plan_total = df['План проекта, шт.'].sum()
                    if plan_total > 0:
                        weighted_sum = (df['Прогноз на месяц, %'] * df['План проекта, шт.']).sum()
                        total_row[col] = weighted_sum / plan_total
                    else:
                        total_row[col] = 0
                else:
                    total_row[col] = 0
                    
            elif col == 'Доля Поручено, %':
                if 'Кол-во визитов до 100% плана, шт.' in df.columns and 'Поручено' in df.columns:
                    need = df['Кол-во визитов до 100% плана, шт.'].sum()
                    assigned = df['Поручено'].sum()
                    total_row[col] = (assigned / need * 100) if need != 0 else 0
                else:
                    total_row[col] = 0
            else:
                total_row[col] = ''
        
        # Заполняем заголовки
        total_row['Код проекта'] = 'Итого'
        total_row['Имя клиента'] = ''
        total_row['Название проекта'] = ''
        
        return pd.DataFrame([total_row])

# Глобальный экземпляр

dataviz = DataVisualizer()












