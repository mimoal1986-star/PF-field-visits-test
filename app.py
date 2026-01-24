# app.py
import streamlit as st
import pandas as pd
import sys
import os

# Настройка путей для импорта
current_dir = os.path.dirname(os.path.abspath(__file__))
utils_path = os.path.join(current_dir, 'utils')
if utils_path not in sys.path:
    sys.path.append(utils_path)

# Настройка страницы
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
if 'cleaned_data' not in st.session_state:
    st.session_state.cleaned_data = {}

# ==============================================
# СЕКЦИЯ 1: ЗАГРУЗКА ФАЙЛОВ
# ==============================================

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

# 2. ЗАГРРУЗКА АВТОКОДИФИКАЦИИ
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

# ==============================================
# СЕКЦИЯ 2: СТАТУС ЗАГРУЗКИ
# ==============================================
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
    
    # Временная кнопка без перехода
    if st.button("🚀 Начать обработку данных", type="primary"):
        st.info("Обработка данных будет реализована в следующем шаге")
        
else:
    st.warning(f"⚠️ Загружено {len(st.session_state.uploaded_files)} из 4 файлов")
    missing = [f for f in ['портал', 'автокодификация', 'сервизория', 'иерархия'] 
               if f not in st.session_state.uploaded_files]
    st.write(f"Ожидаются: {', '.join(missing)}")

# ==============================================
# СЕКЦИЯ 3: ТЕСТИРОВАНИЕ ОЧИСТКИ ДАННЫХ
# ==============================================
if len(st.session_state.uploaded_files) > 0:
    st.markdown("---")
    st.subheader("🧹 Тестирование очистки данных")
    
    # Выбор файла для очистки
    st.write("**Выберите файл для очистки:**")
    
    file_options = {
        'Гугл таблица (Проекты Сервизория)': 'сервизория',
        'Массив (Портал)': 'портал',
        'Автокодификация': 'автокодификация',
        'Иерархия ЗОД-АСС': 'иерархия'
    }
    
    available_files = {k: v for k, v in file_options.items() 
                      if v in st.session_state.uploaded_files}
    
    if available_files:
        selected_file_name = st.selectbox(
            "Выберите файл",
            options=list(available_files.keys()),
            key="file_selector"
        )
        
        selected_file_key = available_files[selected_file_name]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("🧪 Протестировать очистку", type="primary"):
                if selected_file_key in st.session_state.uploaded_files:
                    try:
                        from data_cleaner import data_cleaner
                        
                        original_df = st.session_state.uploaded_files[selected_file_key]
                        
                        # Определяем какой метод очистки использовать
                        if selected_file_key == 'сервизория':
                            with st.spinner("Очищаю Гугл таблицу..."):
                                cleaned_df = data_cleaner.clean_google(original_df)
                                process_name = "Гугл таблицы"
                        elif selected_file_key == 'портал':
                            with st.spinner("Очищаю массив (портал)..."):
                                cleaned_df = data_cleaner.clean_array(original_df)
                                process_name = "Массива"
                        else:
                            with st.spinner("Базовая очистка..."):
                                # Базовая очистка для остальных файлов
                                cleaned_df = original_df.copy()
                                process_name = "файла"
                        
                        if cleaned_df is not None:
                            # Всегда сохраняем очищенные данные (даже если не было изменений)
                            st.session_state.cleaned_data[selected_file_key] = cleaned_df
                            
                            if cleaned_df.equals(original_df):
                                st.info("ℹ️ Данные уже были чистыми, изменений не потребовалось")
                            else:
                                st.success(f"✅ Очистка {process_name} завершена!")
                                
                                # Сравнение до/после (только если были изменения)
                                with st.expander("📊 Сравнение до/после очистки", expanded=True):
                                    col_a, col_b = st.columns(2)
                                    
                                    with col_a:
                                        st.write(f"**До очистки ({selected_file_name}):**")
                                        st.write(f"Строк: {len(original_df)}")
                                        st.write(f"Колонок: {len(original_df.columns)}")
                                        st.dataframe(original_df.head(3))
                                    
                                    with col_b:
                                        st.write(f"**После очистки:**")
                                        st.write(f"Строк: {len(cleaned_df)}")
                                        st.write(f"Колонок: {len(cleaned_df.columns)}")
                                        st.dataframe(cleaned_df.head(3))

                                # Выгрузка в Excel для сверки (только если были изменения)
                                st.markdown("---")
                                st.subheader("📥 Выгрузка для сверки")
                                
                                if selected_file_key == 'портал':  # Массив
                                    excel_file = data_cleaner.export_array_to_excel(cleaned_df)
                                    file_name = "очищенный_массив.xlsx"
                                    help_text = "Файл содержит 3 вкладки: ОЧИЩЕННЫЙ МАССИВ, СТРОКИ С Н Д, НУЛИ В ДАТАХ"
                                else:  # Гугл таблица и другие
                                    excel_file = data_cleaner.export_to_excel(
                                        original_df, 
                                        cleaned_df,
                                        filename=f"очищенный_{selected_file_key}"
                                    )
                                    file_name = f"очищенный_{selected_file_key}.xlsx"
                                    help_text = "Файл содержит 3 вкладки: ОРИГИНАЛ, ОЧИЩЕННЫЙ, СРАВНЕНИЕ"
                                
                                if excel_file:
                                    st.download_button(
                                        label=f"⬇️ Скачать Excel с сравнением ({selected_file_name})",
                                        data=excel_file,
                                        file_name=file_name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        help=help_text
                                    )
                                    
                                    st.info(f"""
                                    **Файл содержит 3 вкладки:**
                                    1. 📋 **ОРИГИНАЛ** - исходные данные
                                    2. ✅ **ОЧИЩЕННЫЙ** - после всех преобразований  
                                    3. 🔍 **СРАВНЕНИЕ** - статистика изменений
                                    """)
                                
                                # Показываем основные изменения
                                with st.expander("🔍 Детали изменений"):
                                    changes = []
                                    
                                    # Изменение количества строк
                                    if len(original_df) != len(cleaned_df):
                                        changes.append(f"Строк: {len(original_df)} → {len(cleaned_df)}")
                                    
                                    # Изменение количества колонок
                                    if len(original_df.columns) != len(cleaned_df.columns):
                                        changes.append(f"Колонок: {len(original_df.columns)} → {len(cleaned_df.columns)}")
                                    
                                    # Добавленные колонки
                                    added_cols = set(cleaned_df.columns) - set(original_df.columns)
                                    if added_cols:
                                        changes.append(f"Добавлены колонки: {', '.join(added_cols)}")
                                    
                                    # Удаленные колонки
                                    removed_cols = set(original_df.columns) - set(cleaned_df.columns)
                                    if removed_cols:
                                        changes.append(f"Удалены колонки: {', '.join(removed_cols)}")
                                    
                                    if changes:
                                        for change in changes:
                                            st.write(f"- {change}")
                                    else:
                                        st.write("Структура данных не изменилась")
                        else:
                            st.error("❌ Очистка не удалась")
                            
                    except ImportError as e:
                        st.error(f"❌ Не удалось импортировать data_cleaner: {e}")
                        st.info("Убедитесь что файл utils/data_cleaner.py существует")
                    except Exception as e:
                        st.error(f"❌ Ошибка при очистке: {e}")
                else:
                    st.error(f"Файл '{selected_file_key}' не найден в загруженных данных")
        
        with col2:
            if st.button("🧹 Очистить ВСЕ файлы", type="secondary"):
                try:
                    from data_cleaner import data_cleaner
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    files_to_clean = list(st.session_state.uploaded_files.keys())
                    total_files = len(files_to_clean)
                    
                    for i, file_key in enumerate(files_to_clean):
                        status_text.text(f"Очищаю {file_key}... ({i+1}/{total_files})")
                        
                        original_df = st.session_state.uploaded_files[file_key]
                        
                        # Пока только для Гугл таблицы есть полная очистка
                        if file_key == 'сервизория':
                            cleaned_df = data_cleaner.clean_google(original_df)
                        else:
                            cleaned_df = original_df.copy()
                        
                        if cleaned_df is not None:
                            st.session_state.cleaned_data[file_key] = cleaned_df
                        
                        progress_bar.progress((i + 1) / total_files)
                    
                    status_text.text("✅ Все файлы очищены!")
                    st.success(f"Очищено {total_files} файлов")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Ошибка при массовой очистке: {e}")
        
        with col3:
            if st.button("🗑️ Очистить ВСЕ сохраненные данные", type="secondary"):
                st.session_state.cleaned_data.clear()
                st.success("✅ Все очищенные данные удалены из session_state")
                st.rerun()
    
    else:
        st.warning("⚠️ Нет загруженных файлов для очистки")

# ==============================================
# СЕКЦИЯ 4: ИНФОРМАЦИЯ ОБ ОЧИЩЕННЫХ ДАННЫХ
# ==============================================

if st.session_state.cleaned_data:
    st.markdown("---")
    
    st.subheader("🎯 Обогащение Массива кодами проектов")

    st.write("🐛 **ДИАГНОСТИКА:** cleaned_data существует")
    st.write(f"- Ключи: {list(st.session_state.cleaned_data.keys())}")
    st.write(f"- 'портал' есть? {'портал' in st.session_state.cleaned_data}")
    st.write(f"- 'сервизория' есть? {'сервизория' in st.session_state.cleaned_data}")
    
    # Проверяем, есть ли оба необходимых файла
    if 'портал' in st.session_state.cleaned_data and 'сервизория' in st.session_state.cleaned_data:
        st.success("✅ Условие выполнено! Кнопка должна быть видна.")
        
        # 🔴 ДОБАВЬТЕ ФЛАГ ДЛЯ СОХРАНЕНИЯ СОСТОЯНИЯ МЕЖДУ ПЕРЕЗАПУСКАМИ
        if 'enrichment_completed' not in st.session_state:
            st.session_state.enrichment_completed = False
        
        # Если обогащение еще не выполнялось - показываем кнопку
        if not st.session_state.enrichment_completed:
            if st.button("🔍 Обогатить Массив кодами проектов", type="primary"):
                st.write("🔍 **1. Кнопка нажата!**")

                try:
                    from data_cleaner import data_cleaner
                    st.write("✅ 2. data_cleaner импортирован")
                    
                    array_df = st.session_state.cleaned_data['портал']
                    projects_df = st.session_state.cleaned_data['сервизория']
                    st.write(f"✅ 3. Данные получены")
                    
                    st.write("🔍 4. Вызываю функцию enrich_array_with_project_codes...")
                    
                    with st.spinner("Ищу и заполняю коды проектов..."):
                        try:
                            # ВЫЗОВ ФУНКЦИИ
                            enriched_array, discrepancy_df, stats = data_cleaner.enrich_array_with_project_codes(
                                array_df, projects_df
                            )
                            
                            st.write(f"✅ 5. Функция выполнилась успешно!")
                            st.write(f"   - Размер enriched_array: {len(enriched_array)}")
                            st.write(f"   - Размер discrepancy_df: {len(discrepancy_df)}")
                            st.write(f"   - Статистика: {stats}")
                            
                            # Сохраняем обогащенный массив обратно
                            st.session_state.cleaned_data['портал'] = enriched_array
                            
                            # Сохраняем расхождения
                            if not discrepancy_df.empty:
                                st.session_state['array_discrepancies'] = discrepancy_df
                                st.session_state['discrepancy_stats'] = stats
                                
                                # Создаем Excel файл с расхождениями
                                st.write("🔍 6. Создаю Excel файл...")
                                excel_file = data_cleaner.export_discrepancies_to_excel(discrepancy_df)
                                
                                if excel_file:
                                    st.success(f"✅ Файл создан! {len(excel_file.getvalue())} байт")
                                    st.download_button(
                                        label=f"⬇️ Скачать 'Расхождение Массив.xlsx' ({len(discrepancy_df)} строк)",
                                        data=excel_file,
                                        file_name="Расхождение_Массив.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="download_discrepancies"
                                    )
                                else:
                                    st.warning("⚠️ Excel файл не создан")
                            
                            st.success(f"✅ Обогащение завершено! Заполнено {stats.get('filled', 0)} кодов.")
                            
                            # 🔴 ОТМЕЧАЕМ ЧТО ОБОГАЩЕНИЕ ВЫПОЛНЕНО
                            st.session_state.enrichment_completed = True
                            
                            # 🔴 УБИРАЕМ st.rerun() - он вызывает бесконечный цикл
                            # Вместо этого даем пользователю выбор
                            st.markdown("---")
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("🔄 Продолжить работу", type="primary"):
                                    # Просто продолжаем показ результатов
                                    pass
                            with col2:
                                if st.button("🗑️ Сбросить и начать заново", type="secondary"):
                                    st.session_state.enrichment_completed = False
                                    st.rerun()
                            
                        except Exception as func_error:
                            st.error(f"❌ ОШИБКА ВНУТРИ ФУНКЦИИ: {str(func_error)}")
                            import traceback
                            st.text("📋 **ТРАССИРОВКА ОШИБКИ:**")
                            st.code(traceback.format_exc())
                            
                except Exception as e:
                    st.error(f"❌ ОБЩАЯ ОШИБКА: {str(e)}")
                    import traceback
                    st.text("📋 **ПОЛНАЯ ТРАССИРОВКА:**")
                    st.code(traceback.format_exc())
        else:
            # 🔴 ЕСЛИ ОБОГАЩЕНИЕ УЖЕ ВЫПОЛНЕНО - ПОКАЗЫВАЕМ РЕЗУЛЬТАТЫ И КНОПКУ ДЛЯ ПОВТОРА
            st.success("✅ Обогащение уже выполнено!")
            if st.button("🔄 Выполнить обогащение снова", type="secondary"):
                st.session_state.enrichment_completed = False
                st.rerun()
                
    else:
        st.info("Для обогащения нужны оба файла: 'портал' (Массив) и 'сервизория' (Проекты Сервизория)")
    
    # 🔴 УБИРАЕМ st.stop() - он блокирует дальнейший код
    # st.stop()  # УДАЛИТЬ ЭТУ СТРОКУ
    
    st.markdown("---")
    st.subheader("✅ Очищенные данные")
    
    for name, df in st.session_state.cleaned_data.items():
        with st.expander(f"📁 {name} (очищенный)"):
            st.write(f"Размер: {len(df)} строк × {len(df.columns)} колонок")
            
            # Показываем основные колонки
            st.write("**Колонки:**")
            cols_per_row = 4
            columns = list(df.columns)
            
            for i in range(0, len(columns), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(columns[i:i+cols_per_row]):
                    with cols[j]:
                        st.code(col)
            
            # Показываем предпросмотр
            st.write("**Предпросмотр:**")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Кнопка для скачивания очищенных данных
            if name == 'портал':  # Массив - Excel с вкладками
                try:
                    from data_cleaner import data_cleaner
                    excel_file = data_cleaner.export_array_to_excel(df)
                    if excel_file:
                        st.download_button(
                            label=f"⬇️ Скачать очищенный {name} (Excel с вкладками)",
                            data=excel_file,
                            file_name=f"очищенный_массив.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="3 вкладки: 📊 Данные, ⚠️ Строки с Н Д, 📅 Нули в датах"
                        )
                    else:
                        # Fallback на CSV если Excel не создался
                        csv = df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label=f"⬇️ Скачать очищенный {name} (CSV)",
                            data=csv,
                            file_name=f"очищенный_{name}.csv",
                            mime="text/csv"
                        )
                except Exception as e:
                    st.error(f"Ошибка экспорта в Excel: {e}")
                    # Fallback на CSV
                    csv = df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"⬇️ Скачать очищенный {name} (CSV)",
                        data=csv,
                        file_name=f"очищенный_{name}.csv",
                        mime="text/csv"
                    )
            else:  # Остальные файлы - CSV
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"⬇️ Скачать очищенный {name} (CSV)",
                    data=csv,
                    file_name=f"очищенный_{name}.csv",
                    mime="text/csv"
                )

# ==============================================
# СЕКЦИЯ 5: ИНФОРМАЦИЯ О ПРОЕКТЕ
# ==============================================
with st.sidebar:
    st.header("ℹ️ Информация о проекте")
    
    st.markdown("**Загружено файлов:**")
    st.write(f"📊 {len(st.session_state.uploaded_files)} из 4")
    
    if st.session_state.uploaded_files:
        st.markdown("**Статистика:**")
        for name, df in st.session_state.uploaded_files.items():
            st.write(f"- {name}: {len(df)} строк")
    
    st.markdown("---")
    
    st.markdown("**Очищено файлов:**")
    st.write(f"🧹 {len(st.session_state.cleaned_data)}")
    
    if st.session_state.cleaned_data:
        st.markdown("**Очищенные данные:**")
        for name, df in st.session_state.cleaned_data.items():
            st.write(f"- {name}: {len(df)} строк")
    
    st.markdown("---")
    
    # Кнопка для сброса всех данных
    if st.button("🔄 Сбросить все данные", type="secondary", use_container_width=True):
        st.session_state.uploaded_files.clear()
        st.session_state.cleaned_data.clear()
        st.rerun()
    
    st.markdown("---")
    
    st.markdown("**Следующие шаги:**")
    st.write("1. 🧹 Очистка всех файлов")
    st.write("2. 🔗 Объединение данных")
    st.write("3. 📊 Создание отчетов")
















