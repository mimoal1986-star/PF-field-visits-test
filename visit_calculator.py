# utils/visit_calculator.py
# draft 1.0
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple, List
import io

# ==============================================
# СЕКЦИЯ 1: БАЗОВЫЕ ПОЛЯ 
# ==============================================

class VisitCalculator:
    """Калькулятор плана/факта визитов (воспроизводит логику 'Датасет ПФ1')"""
    
    def extract_base_data(self, field_projects_df):
        """
        Извлекает базовые данные ТОЛЬКО из полевых проектов (столбцы A-H).
        
        Параметры:
        - field_projects_df: DataFrame ТОЛЬКО полевых проектов (уже отфильтровано)
        
        Возвращает:
        - DataFrame с 8 базовыми колонками для отчета (уникальные по Названию проекта)
        """
        try:
            if field_projects_df is None or field_projects_df.empty:
                st.warning("⚠️ Нет полевых проектов для анализа")
                return pd.DataFrame()
            
            # Создаем DataFrame с нужными колонками
            result = pd.DataFrame()
            
            # Берем колонки напрямую из полевых проектов
            result['Код проекта'] = field_projects_df['Код проекта']
            result['Имя клиента'] = field_projects_df['Имя клиента']
            result['Название проекта'] = field_projects_df['Название проекта']
            result['ЗОД'] = field_projects_df['ЗОД']
            result['АСС'] = field_projects_df['АСС']
            result['ЭМ'] = field_projects_df['ЭМ']
            result['Регион short'] = field_projects_df['Регион short']
            result['Регион'] = field_projects_df['Регион']
            
            # Удаляем дубликаты по Названию проекта (как в Excel)
            result = result.drop_duplicates(subset=['Название проекта', 'Код проекта'], keep='first')
            
            st.info(f"📊 Извлечено базовых данных: {len(result)} уникальных полевых проектов")
            return result
            
        except KeyError as e:
            st.error(f"❌ В полевых проектах нет колонки: {e}")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"❌ Ошибка извлечения базовых данных: {str(e)[:100]}")
            return pd.DataFrame()

# Глобальный экземпляр
visit_calculator = VisitCalculator()

