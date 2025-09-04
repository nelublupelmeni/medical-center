# -*- coding: utf-8 -*-
"""
Домашнее задание №02. Отчётный скрипт.

Функции:
- Загрузка данных из Excel
- Сохранение справочников в Pickle
- Генерация текстовых и графических отчётов
- Текстовый пользовательский интерфейс

Тема: Медицинский центр
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def doctors_by_department(data_frame, dept_name=None):
    """
    

    Parameters
    ----------
    data_frame :  pandas.DataFrame
        исходные данные.
    dept_name : str, optional
       Название отделения. Если None, возвращаются все врачи.

    Returns
    -------
    pandas.DataFrame or str
        Таблица с колонками ФАМ, СПЕЦИАЛЬНОСТЬ, НАЗВ_ОТД или сообщение об ошибке.

    """
    if data_frame.empty:
        return "Нет данных в исходной таблице" #проверка на наличие данных в таблице
    try:
        if dept_name:
            selection = data_frame['НАЗВ_ОТД'] == dept_name
            if not selection.any():
                return "Нет данных для указанного отделения"
        else:
            selection = np.ones(len(data_frame), dtype=bool)
        result = data_frame.loc[selection, ['ФАМ', 'СПЕЦИАЛЬНОСТЬ', 'НАЗВ_ОТД']]
        result = result.drop_duplicates()
        return result if not result.empty else "Список пуст"
    except KeyError as error:
        return f"Ошибка формирования отчета: отсутствует столбец {error}"


def appointments_by_date(data_frame, date_start, date_end):
    """
    

    Parameters
    ----------
    data_frame :  pandas.DataFrame
        исходные данные.
   date_start : str
       Начальная дата в формате DD-MM-YYYY.
   date_end : str
       Конечная дата в формате DD-MM-YYYY.

    Returns
    -------
    pandas.DataFrame or str
        Таблица с колонками ДАТА_НАЗН, ФИО_ПАЦ, ФАМ, ДИАГНОЗ или сообщение об ошибке.

    """
    if data_frame.empty:
        return "Нет данных в исходной таблице" 
    try:
        start = pd.to_datetime(date_start, format='%d-%m-%Y')
        end = pd.to_datetime(date_end, format='%d-%m-%Y')
        mask = (data_frame['ДАТА_НАЗН'] >= start) & (data_frame['ДАТА_НАЗН'] <= end)
        result = data_frame.loc[mask, ['ДАТА_НАЗН', 'ФИО_ПАЦ', 'ФАМ', 'ДИАГНОЗ']]
        return result if not result.empty else "Назначения не найдены"
    except ValueError:
        return "Неверный формат даты. Используйте DD-MM-YYYY"
    except KeyError as error:
        return f"Ошибка формирования отчета: отсутствует столбец {error}"


def create_pivot_table(data_frame):
    """
    Сводная таблица диагнозов по отделениям.

    Parameters
    ----------
    data_frame :  pandas.DataFrame
    исходные данные.

    Returns
    -------
    pandas.DataFrame or str
       Сводная таблица или сообщение об ошибке.

    """
    if data_frame.empty:
        return "Нет данных в исходной таблице" 
    try:
        pivot = pd.pivot_table(
            data_frame,
            values='Н_НАЗН',
            index='НАЗВ_ОТД',
            columns='ДИАГНОЗ',
            aggfunc='count',
            fill_value=0
        )
        return pivot
    except (KeyError, ValueError) as error:
        return f"Ошибка создания сводной таблицы: {error}"


def plot_doctors_by_specialty(doctors_df):
    """
    Cтолбчатая диаграмма распределения врачей по специальностям.

    Parameters
    ----------
    doctors_df : pandas.DataFrame
        данные о врачах и их специальностях

    Returns
    -------
    False или True
    """
    try:
        if doctors_df.empty:
            print("Нет данных для построения диаграммы распределения врачей по специальностям") #проверка на данные
            return False

        specialties, counts = np.unique(doctors_df['СПЕЦИАЛЬНОСТЬ'], return_counts=True)
        plt.figure(figsize=(8, 6))
        plt.bar(specialties, counts, color='skyblue')
        plt.title('Врачи по специальностям')
        plt.xticks(rotation=45)
        plt.grid(True)
        plt.tight_layout()
        plt.savefig('doctors_by_specialty.png')
        plt.close()
        return True
    except (KeyError, ValueError) as error:
        print(f"Ошибка построения графика1: {error}")
        return False


def plot_appointments_by_month(data_frame):
    """
    Линейный график назначений по месяцам.

    Parameters
    ----------
    data_frame :  pandas.DataFrame
    исходные данные.

    Returns
    -------
    False или True

    """
    try:
        monthly = data_frame.groupby('МЕСЯЦ').size()
        if monthly.empty or monthly.count() < 2:
            print("Нет данных для построения графика назначений по месяцам") #проверка на данные
            return False
        plt.figure(figsize=(8, 6))
        monthly.plot(kind='line', marker='o', grid=True)
        plt.title('Назначения по месяцам')
        plt.tight_layout()
        plt.savefig('appointments_by_month.png')
        plt.close()
        return True
    except (KeyError, TypeError) as error:
        print(f"Ошибка построения графика 2: {error}")
        return False


def plot_diagnoses_distribution(data_frame):
    """
    Круговая диаграмма распределения диагнозов.

    Parameters
    ----------
    data_frame :  pandas.DataFrame
    исходные данные.

    Returns
    -------
    False или True

    """
    try:

        diagnoses, counts = np.unique(data_frame['ДИАГНОЗ'].dropna(), return_counts=True)
        if diagnoses.shape[0] == 0:
            print("Недостаточно данных для построения диаграммы распределения диагнозов") #проверка на данные
            return False
        plt.figure(figsize=(8, 6))
        plt.pie(counts, labels=diagnoses, autopct='%1.1f%%', startangle=90, shadow=True)
        plt.title('Распределение диагнозов')
        plt.ylabel('')
        plt.tight_layout()
        plt.savefig('diagnoses_distribution.png')
        plt.close()
        return True
    except (KeyError, ValueError) as error:
        print(f"Ошибка построения графика 3: {error}")
        return False


def user_interface(data_frame, departments_df, doctors_df):
    """
    Запускает консольный интерфейс пользователя.

    Parameters
    ----------
    data_frame :  pandas.DataFrame
    departments_df : TYPE
        DESCRIPTION.
    doctors_df : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    try:
        departments = departments_df['НАЗВ_ОТД'].unique()
        print("Доступные отделения:", *departments)
    except KeyError as error:
        print(f"Ошибка получения списка отделений: отсутствует столбец {error}")
        return

    while True:
        print("\nМеню:")
        print("1. Врачи по отделению")
        print("2. Назначения по дате")
        print("3. Сводная таблица")
        print("4. Создать графики")
        print("0. Выход")
        choice = input("Выберите опцию: ")

        if choice == '1':
            dept = input("Введите отделение (оставьте пустым для всех): ")
            result = doctors_by_department(data_frame, dept if dept else None)
            print(result)
        elif choice == '2':
            start = input("Начальная дата (DD-MM-YYYY): ")
            end = input("Конечная дата (DD-MM-YYYY): ")
            result = appointments_by_date(data_frame, start, end)
            print(result)
        elif choice == '3':
            result = create_pivot_table(data_frame)
            print(result)
        elif choice == '4':
            fg_doctors_by_specialty = plot_doctors_by_specialty(doctors_df)
            fg_appointments_by_month = plot_appointments_by_month(data_frame)
            fg_diagnoses_distribution = plot_diagnoses_distribution(data_frame)
            if fg_doctors_by_specialty * fg_appointments_by_month * fg_diagnoses_distribution:
                print("Графики сохранены в PNG-файлы")
            elif fg_diagnoses_distribution + fg_appointments_by_month + fg_doctors_by_specialty > 1:
                print("Сохранено часть графиков") #доп выводы
            else:
                print("Ни один график не сохранён")

        elif choice == '0':
            break
        else:
            print("Неверная опция. Попробуйте снова.")


def load_and_save_data(data_dir, excel_file):
    """
    Загружает данные из Excel и сохраняет листы в Pickle-файлы.

    Parameters
    ----------
    data_dir : str
        Путь к директории для сохранения Pickle-файлов
    excel_file : str
        Путь к Excel-файлу.
    Returns
    -------
    bool
        True, если данные успешно сохранены, False в случае ошибки.

    """
    try:
        excel_data = pd.ExcelFile(excel_file)
        print("Найдены листы Excel:", excel_data.sheet_names)
    except FileNotFoundError:
        print("Ошибка: Файл medical_center_db.xlsx не найден")
        print("Убедитесь, что файл находится в директории проекта")
        return False

    sheet_mapping = {
        'departaments': 'DEPARTMENTS',
        'doctors': 'DOCTORS',
        'patients': 'PATIENTS',
        'diagnosis': 'DIAGNOSES',
        'appointment': 'APPOINTMENTS'
    }
    for sheet_name in excel_data.sheet_names:
        if sheet_name == 'model':
            continue
        try:
            var_name = sheet_mapping.get(sheet_name, sheet_name)
            data_f = excel_data.parse(sheet_name=sheet_name)
            data_f.to_pickle(f"{data_dir}/{var_name}.pick")
            print(f"Лист '{sheet_name}' сохранен как {var_name}.pick")
        except ValueError as error:
            print(f"Ошибка обработки листа {sheet_name}: {error}")
            return False
    return True


def merge_data(data_dir):
    """
    Загружает Pickle-файлы и объединяет данные в один DataFrame.

    Parameters
    ----------
    data_dir : str
        Путь к директории с Pickle-файлами.

    Returns
    -------
    pandas.DataFrame
       Объединённая таблица
    pandas.DataFrame
        Данные об отделениях
   pandas.DataFrame
        Данные о врачах

    """
    try:
        departments_df = pd.read_pickle(f"{data_dir}/DEPARTMENTS.pick")
        doctors_df = pd.read_pickle(f"{data_dir}/DOCTORS.pick")
        patients_df = pd.read_pickle(f"{data_dir}/PATIENTS.pick")
        diagnoses_df = pd.read_pickle(f"{data_dir}/DIAGNOSES.pick")
        appointments_df = pd.read_pickle(f"{data_dir}/APPOINTMENTS.pick")
    except FileNotFoundError as error:
        print(f"Ошибка загрузки Pickle-файлов: {error}")
        print("Убедитесь, что .pick файлы находятся в директории data")
        return None, None, None

    try:
        merged_df = appointments_df.merge(doctors_df, on='Н_ВРАЧ', how='left')
        merged_df = merged_df.merge(patients_df, on='Н_ПАЦ', how='left')
        merged_df = merged_df.merge(diagnoses_df, on='Н_ДИАГ', how='left')
        merged_df = merged_df.merge(departments_df, on='Н_ОТД', how='left')
        if merged_df['ДАТА_НАЗН'].dtype != 'datetime64[ns]':
            merged_df['ДАТА_НАЗН'] = pd.to_datetime(
                merged_df['ДАТА_НАЗН'], format='%d-%m-%Y', errors='coerce'
            )
        merged_df['МЕСЯЦ'] = merged_df['ДАТА_НАЗН'].dt.to_period('M')
        return merged_df, departments_df, doctors_df
    except (KeyError, ValueError) as error:
        print(f"Ошибка объединения таблиц: {error}")
        return None, None, None


def main():
    """
    Основная функция для выполнения скрипта.
    """
    data_dir = "./data"
    try:
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
    except OSError as error:
        print(f"Ошибка создания директории {data_dir}: {error}")
        return

    excel_file = "./medical_center_db.xlsx"
    if not load_and_save_data(data_dir, excel_file):
        return

    merged_df, departments_df, doctors_df = merge_data(data_dir)
    if merged_df is None:
        return

    user_interface(merged_df, departments_df, doctors_df)


if __name__ == "__main__":
    main()
    