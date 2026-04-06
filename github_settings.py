# utils/github_settings.py

import json
import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import requests
from typing import Optional, Dict, List, Tuple

class GitHubSettingsManager:
    """
    Класс для работы с настройками проектов через GitHub API
    Хранит исключенные и добавленные проекты в JSON-файле в репозитории
    """
    
    def __init__(self):
        """Инициализация с настройками из Streamlit Secrets"""
        # Проверяем наличие секретов
        if 'github' not in st.secrets:
            st.error("❌ Секреты GitHub не найдены. Добавьте их в Streamlit Secrets")
            self.available = False
            return
            
        self.token = st.secrets['github']['token']
        self.repo = st.secrets['github']['repo']
        self.branch = st.secrets['github']['branch']
        self.file_path = st.secrets['github']['settings_path']
        
        # URL для GitHub API
        self.api_url = f"https://api.github.com/repos/{self.repo}/contents/{self.file_path}"
        self.headers = {
            "Authorization": f"token {self.token}",
            "Accept": "application/vnd.github.v3+json"
        }
        
        # Текущий пользователь (для истории)
        self.current_user = st.session_state.get('username', 'unknown_user')
        self.available = True
    
    def _get_file_sha(self) -> Optional[str]:
        """Получает SHA текущего файла (нужен для обновления)"""
        try:
            response = requests.get(
                self.api_url,
                headers=self.headers,
                params={"ref": self.branch}
            )
            
            if response.status_code == 200:
                return response.json()["sha"]
            return None
        except Exception as e:
            st.error(f"Ошибка получения SHA: {e}")
            return None
    
    def _read_file(self) -> Dict:
        """Читает файл настроек с GitHub"""
        try:
            response = requests.get(
                self.api_url,
                headers=self.headers,
                params={"ref": self.branch}
            )
            
            if response.status_code == 200:
                content = response.json()
                # Декодируем из base64
                file_content = base64.b64decode(content["content"]).decode("utf-8")
                return json.loads(file_content)
            else:
                # Если файла нет, создаем структуру по умолчанию
                return self._create_default_structure()
                
        except Exception as e:
            st.warning(f"Не удалось прочитать файл настроек: {e}")
            return self._create_default_structure()
    
    def _write_file(self, data: Dict) -> bool:
        """Записывает файл настроек на GitHub"""
        try:
            # Получаем SHA для обновления
            sha = self._get_file_sha()
            
            # Конвертируем в JSON
            content = json.dumps(data, indent=2, ensure_ascii=False)
            content_bytes = content.encode("utf-8")
            content_base64 = base64.b64encode(content_bytes).decode("utf-8")
            
            # Формируем запрос
            payload = {
                "message": f"Обновление настроек проектов от {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                "content": content_base64,
                "branch": self.branch
            }
            
            if sha:
                payload["sha"] = sha
            
            response = requests.put(
                self.api_url,
                headers=self.headers,
                json=payload
            )
            
            return response.status_code in [200, 201]
            
        except Exception as e:
            st.error(f"Ошибка записи файла: {e}")
            return False
    
    def _create_default_structure(self) -> Dict:
        """Создает структуру настроек по умолчанию"""
        return {
            "excluded_projects": [],
            "included_projects": [],
            "history": [],
            "last_updated": datetime.now().isoformat(),
            "version": "1.0"
        }
    
    def _project_to_dict(self, row) -> Dict:
        """Конвертирует строку DataFrame в словарь для JSON"""
        # Принудительно преобразуем волну в строку
        wave_value = row.get('Волна', '')
        if pd.isna(wave_value):
            wave_str = ''
        else:
            wave_str = str(wave_value)  # гарантированно строка
        
        return {
            "project_name": str(row.get('Название проекта', '')),
            "wave_name": wave_str,  # используем преобразованное значение
            "project_code": str(row.get('Код проекта', '')),
            "portal": str(row.get('ПО', '')),
            "fio_om": str(row.get('ФИО ОМ', '')),
            "added_at": datetime.now().isoformat(),
            "added_by": self.current_user
        }
    
    def _dict_to_dataframe(self, projects_list: List[Dict]) -> pd.DataFrame:
        """Конвертирует список словарей в DataFrame"""
        if not projects_list:
            return pd.DataFrame(columns=['Название проекта', 'Волна', 'Код проекта', 'ПО', 'ФИО ОМ'])
        
        df = pd.DataFrame(projects_list)
        # Переименовываем колонки для отображения
        rename_map = {
            'project_name': 'Название проекта',
            'wave_name': 'Волна',
            'project_code': 'Код проекта',
            'portal': 'ПО',
            'fio_om': 'ФИО ОМ'
        }
        df = df.rename(columns=rename_map)
        return df[['Название проекта', 'Волна', 'Код проекта', 'ПО', 'ФИО ОМ']]
    
    # === ПУБЛИЧНЫЕ МЕТОДЫ ===
    
    def get_excluded_projects(self) -> pd.DataFrame:
        """Получить список исключенных проектов"""
        data = self._read_file()
        return self._dict_to_dataframe(data.get("excluded_projects", []))
    
    def get_included_projects(self) -> pd.DataFrame:
        """Получить список добавленных проектов"""
        data = self._read_file()
        return self._dict_to_dataframe(data.get("included_projects", []))
    
    def add_to_excluded(self, projects_df: pd.DataFrame) -> Tuple[bool, str]:
        """
        Добавить проекты в исключенные
        Returns: (успех, сообщение)
        """
        if projects_df.empty:
            return False, "Нет проектов для добавления"
        
        data = self._read_file()
        excluded_list = data.get("excluded_projects", [])
        history = data.get("history", [])
        
        added_count = 0
        for _, row in projects_df.iterrows():
            # Проверяем, нет ли уже такого проекта
            project_dict = self._project_to_dict(row)
            exists = any(
                p['project_name'] == project_dict['project_name'] and
                p['wave_name'] == project_dict['wave_name'] and
                p['project_code'] == project_dict['project_code']
                for p in excluded_list
            )
            
            if not exists:
                excluded_list.append(project_dict)
                added_count += 1
                
                # Добавляем в историю
                history.append({
                    "action": "add_to_excluded",
                    "project_name": project_dict['project_name'],
                    "wave_name": project_dict['wave_name'],
                    "project_code": project_dict['project_code'],
                    "timestamp": datetime.now().isoformat(),
                    "user": self.current_user
                })
        
        if added_count > 0:
            data["excluded_projects"] = excluded_list
            data["history"] = history
            data["last_updated"] = datetime.now().isoformat()
            
            if self._write_file(data):
                return True, f"✅ Добавлено {added_count} проектов в исключенные"
            else:
                return False, "❌ Ошибка при сохранении в GitHub"
        else:
            return False, "⚠️ Проекты уже есть в списке исключенных"
    
    def remove_from_excluded(self, projects_df: pd.DataFrame) -> Tuple[bool, str]:
        """Удалить проекты из исключенных"""
        if projects_df.empty:
            return False, "Нет проектов для удаления"
        
        data = self._read_file()
        excluded_list = data.get("excluded_projects", [])
        history = data.get("history", [])
        
        removed_count = 0
        for _, row in projects_df.iterrows():
            project_name = str(row.get('Название проекта', ''))
            wave_name = str(row.get('Волна', ''))
            project_code = str(row.get('Код проекта', ''))
            
            # Фильтруем список, убирая совпадения
            new_list = [
                p for p in excluded_list 
                if not (p['project_name'] == project_name and 
                       p['wave_name'] == wave_name and 
                       p['project_code'] == project_code)
            ]
            
            if len(new_list) < len(excluded_list):
                removed_count += 1
                excluded_list = new_list
                
                # Добавляем в историю
                history.append({
                    "action": "remove_from_excluded",
                    "project_name": project_name,
                    "wave_name": wave_name,
                    "project_code": project_code,
                    "timestamp": datetime.now().isoformat(),
                    "user": self.current_user
                })
        
        if removed_count > 0:
            data["excluded_projects"] = excluded_list
            data["history"] = history
            data["last_updated"] = datetime.now().isoformat()
            
            if self._write_file(data):
                return True, f"✅ Удалено {removed_count} проектов из исключенных"
            else:
                return False, "❌ Ошибка при сохранении в GitHub"
        else:
            return False, "⚠️ Проекты не найдены в списке исключенных"
    
    def add_to_included(self, projects_df: pd.DataFrame) -> Tuple[bool, str]:
        """Добавить проекты во включенные"""
        if projects_df.empty:
            return False, "Нет проектов для добавления"
        
        data = self._read_file()
        included_list = data.get("included_projects", [])
        history = data.get("history", [])
        
        added_count = 0
        for _, row in projects_df.iterrows():
            project_dict = self._project_to_dict(row)
            exists = any(
                p['project_name'] == project_dict['project_name'] and
                p['wave_name'] == project_dict['wave_name'] and
                p['project_code'] == project_dict['project_code']
                for p in included_list
            )
            
            if not exists:
                included_list.append(project_dict)
                added_count += 1
                
                history.append({
                    "action": "add_to_included",
                    "project_name": project_dict['project_name'],
                    "wave_name": project_dict['wave_name'],
                    "project_code": project_dict['project_code'],
                    "timestamp": datetime.now().isoformat(),
                    "user": self.current_user
                })
        
        if added_count > 0:
            data["included_projects"] = included_list
            data["history"] = history
            data["last_updated"] = datetime.now().isoformat()
            
            if self._write_file(data):
                return True, f"✅ Добавлено {added_count} проектов во включенные"
            else:
                return False, "❌ Ошибка при сохранении в GitHub"
        else:
            return False, "⚠️ Проекты уже есть в списке включенных"
    
    def remove_from_included(self, projects_df: pd.DataFrame) -> Tuple[bool, str]:
        """Удалить проекты из включенных"""
        if projects_df.empty:
            return False, "Нет проектов для удаления"
        
        data = self._read_file()
        included_list = data.get("included_projects", [])
        history = data.get("history", [])
        
        removed_count = 0
        for _, row in projects_df.iterrows():
            project_name = str(row.get('Название проекта', ''))
            wave_name = str(row.get('Волна', ''))
            project_code = str(row.get('Код проекта', ''))
            
            new_list = [
                p for p in included_list 
                if not (p['project_name'] == project_name and 
                       p['wave_name'] == wave_name and 
                       p['project_code'] == project_code)
            ]
            
            if len(new_list) < len(included_list):
                removed_count += 1
                included_list = new_list
                
                history.append({
                    "action": "remove_from_included",
                    "project_name": project_name,
                    "wave_name": wave_name,
                    "project_code": project_code,
                    "timestamp": datetime.now().isoformat(),
                    "user": self.current_user
                })
        
        if removed_count > 0:
            data["included_projects"] = included_list
            data["history"] = history
            data["last_updated"] = datetime.now().isoformat()
            
            if self._write_file(data):
                return True, f"✅ Удалено {removed_count} проектов из включенных"
            else:
                return False, "❌ Ошибка при сохранении в GitHub"
        else:
            return False, "⚠️ Проекты не найдены в списке включенных"
    
    def get_history(self, days: int = 30) -> pd.DataFrame:
        """Получить историю изменений за последние N дней"""
        data = self._read_file()
        history = data.get("history", [])
        
        if not history:
            return pd.DataFrame(columns=['Дата', 'Пользователь', 'Действие', 'Проект', 'Волна', 'Код'])
        
        df = pd.DataFrame(history)
        
        # Фильтр по датам
        if days > 0:
            cutoff = datetime.now() - pd.Timedelta(days=days)
            df['timestamp'] = pd.to_datetime(df['timestamp'])
            df = df[df['timestamp'] >= cutoff]
        
        if df.empty:
            return pd.DataFrame(columns=['Дата', 'Пользователь', 'Действие', 'Проект', 'Волна', 'Код'])
        
        # Форматируем для отображения
        df['Дата'] = pd.to_datetime(df['timestamp']).dt.strftime('%Y-%m-%d %H:%M')
        df['Пользователь'] = df['user']
        
        # Переводим действия на русский
        action_map = {
            'add_to_excluded': '➕ Исключил',
            'remove_from_excluded': '➖ Вернул из исключенных',
            'add_to_included': '➕ Добавил в расчет',
            'remove_from_included': '➖ Убрал из расчета',
            'clear_all': '🗑️ Очистил все настройки'
        }
        df['Действие'] = df['action'].map(action_map).fillna(df['action'])
        
        # СОЗДАЕМ колонки project_name, wave_name, project_code если их нет
        for col in ['project_name', 'wave_name', 'project_code']:
            if col not in df.columns:
                df[col] = ''
        
        # Теперь безопасно обращаемся к колонкам
        return df[['Дата', 'Пользователь', 'Действие', 'project_name', 'wave_name', 'project_code']].rename(
            columns={
                'project_name': 'Проект',
                'wave_name': 'Волна',
                'project_code': 'Код'
            }
        )
    
    def clear_all_settings(self) -> Tuple[bool, str]:
        """Очистить все настройки (для отладки)"""
        data = self._read_file()
        
        # Очищаем списки проектов
        data["excluded_projects"] = []
        data["included_projects"] = []
        
        # Добавляем запись в историю со ВСЕМИ полями (пустыми)
        data["history"].append({
            "action": "clear_all",
            "project_name": "",  # пустое, но колонка сохранится
            "wave_name": "",     # пустое, но колонка сохранится
            "project_code": "",  # пустое, но колонка сохранится
            "timestamp": datetime.now().isoformat(),
            "user": self.current_user
        })
        
        data["last_updated"] = datetime.now().isoformat()
        
        if self._write_file(data):
            return True, "✅ Все настройки очищены"
        else:
            return False, "❌ Ошибка при очистке"


class PlanAdjustmentManager:
    """Класс для работы с корректировками плана проектов"""
    
    def __init__(self, settings_manager):
        self.settings_manager = settings_manager
        self.file_path = 'plan_adjustments.json'
        self.repo = settings_manager.repo
        self.api_url = f"https://api.github.com/repos/{self.repo}/contents/{self.file_path}"
        self.headers = settings_manager.headers
        self.branch = settings_manager.branch
        self.current_user = st.session_state.get('username', 'unknown_user')

    def get_all_adjustments(self) -> dict:
        """Возвращает словарь {ключ: сумма_корректировок} за один запрос"""
        adjustments = self._read_adjustments()
        result = {}
        for adj in adjustments:
            key = (adj.get('project_name', ''), adj.get('wave_name', ''), adj.get('project_code', ''))
            result[key] = result.get(key, 0) + adj.get('adjustment_value', 0)
        return result
    
    def _get_file_sha(self):
        """Получает SHA текущего файла"""
        try:
            response = requests.get(
                self.api_url,
                headers=self.headers,
                params={"ref": self.branch}
            )
            if response.status_code == 200:
                return response.json()["sha"]
            return None
        except Exception as e:
            print(f"Ошибка получения SHA: {e}")
            return None
    
    def _read_adjustments(self):
        """Читает корректировки из GitHub"""
        try:
            response = requests.get(
                self.api_url,
                headers=self.headers,
                params={"ref": self.branch}
            )
            if response.status_code == 200:
                content = response.json()
                file_content = base64.b64decode(content["content"]).decode("utf-8")
                data = json.loads(file_content)
                return data.get("adjustments", [])
            return []
        except Exception as e:
            print(f"Ошибка чтения: {e}")
            return []
    
    def _write_adjustments(self, adjustments):
        """Записывает корректировки в GitHub"""
        try:
            sha = self._get_file_sha()
            data = {
                "adjustments": adjustments,
                "last_updated": datetime.now().isoformat()
            }
            content = json.dumps(data, indent=2, ensure_ascii=False)
            content_bytes = content.encode("utf-8")
            content_base64 = base64.b64encode(content_bytes).decode("utf-8")
            
            payload = {
                "message": f"Обновление корректировок плана от {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                "content": content_base64,
                "branch": self.branch
            }
            if sha:
                payload["sha"] = sha
            
            response = requests.put(self.api_url, headers=self.headers, json=payload)
            return response.status_code in [200, 201]
        except Exception as e:
            print(f"Ошибка записи: {e}")
            return False
    
    def get_adjustments(self, project_name=None, wave_name=None, project_code=None):
        """Получает корректировки (все или для конкретного проекта)"""
        adjustments = self._read_adjustments()
        
        if project_name and wave_name and project_code:
            # Фильтруем по проекту
            return [a for a in adjustments 
                   if a.get('project_name') == project_name 
                   and a.get('wave_name') == wave_name 
                   and a.get('project_code') == project_code]
        return adjustments
    
    def get_total_adjustment(self, project_name, wave_name, project_code):
        """Возвращает суммарную корректировку для проекта"""
        adjustments = self.get_adjustments(project_name, wave_name, project_code)
        total = sum(a.get('adjustment_value', 0) for a in adjustments)
        return total
    
    def add_adjustment(self, project_name, wave_name, project_code, adjustment_value):
        """Добавляет новую корректировку"""
        adjustments = self._read_adjustments()
        
        adjustments.append({
            "project_name": project_name,
            "wave_name": wave_name,
            "project_code": project_code,
            "adjustment_value": adjustment_value,
            "created_at": datetime.now().isoformat(),
            "created_by": self.current_user
        })
        
        if self._write_adjustments(adjustments):
            return True, f"✅ Корректировка {adjustment_value} добавлена"
        return False, "❌ Ошибка сохранения"
    
    def clear_project_adjustments(self, project_name, wave_name, project_code):
        """Очищает все корректировки для проекта"""
        adjustments = self._read_adjustments()
        new_adjustments = [a for a in adjustments 
                          if not (a.get('project_name') == project_name 
                                 and a.get('wave_name') == wave_name 
                                 and a.get('project_code') == project_code)]
        
        if self._write_adjustments(new_adjustments):
            return True, "✅ Корректировки очищены"
        return False, "❌ Ошибка сохранения"
            

# Для удобства создаем глобальный экземпляр при импорте
def get_settings_manager():
    """Возвращает экземпляр менеджера настроек"""
    if 'github_settings' not in st.session_state:
        st.session_state.github_settings = GitHubSettingsManager()
    return st.session_state.github_settings

def get_plan_adjustment_manager():
    """Возвращает экземпляр менеджера корректировок плана"""
    if 'plan_adjustment_manager' not in st.session_state:
        settings_manager = get_settings_manager()
        st.session_state.plan_adjustment_manager = PlanAdjustmentManager(settings_manager)
    return st.session_state.plan_adjustment_manager
    
