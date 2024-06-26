import datetime
import numpy as np
import pandas as pd
import statsmodels.api as sm
import os
import re


def data_preparation(file, sheet, name_column_time, name_column_for_predict, name_column_factors,
                     date_format='%d.%m.%Y'):
    """
    Подготавливает данные из файла Excel.

    Parameters:
    - file: str, путь к файлу Excel
    - sheet: str, название листа в файле Excel
    - name_column_time: str, название столбца с временными метками
    - name_column_factors: list, список названий столбцов-факторов

    Returns:
    - result_df: DataFrame, подготовленные данные
    """
    result_data = []

    # Чтение данных из Excel файла
    data = pd.read_excel(file, sheet, engine='openpyxl')

    # Преобразование времени
    data_time = data[name_column_time]
    result_data.append(data_time)

    # Преобразование остальных факторов
    for factor in name_column_factors:
        factor_data = data[factor]
        result_data.append(factor_data)

    # Собираем все данные в один DataFrame
    result_df = pd.concat(result_data, axis=1)

    # Проверяем, есть ли в результирующих данных столбец, который мы планируем прогнозировать
    if name_column_for_predict not in result_df.columns:
        # Преобразование времени
        result_df.append(data[name_column_for_predict])

    # Удаление строк, где значение в столбце name_column_time отличается от дд.мм.гггг
    result_df[name_column_time] = pd.to_datetime(result_df[name_column_time], format=date_format, errors='coerce')
    result_df = result_df.dropna(subset=[name_column_time])

    result_df.reset_index(drop=True, inplace=True)  # Переиндексируем DataFrame

    return result_df


def create_lags(data, columns, lag_count, need_create_lag_for_predictable, chosen_column_for_predict):
    """
    Создает лаги для указанных столбцов данных.

    Parameters:
    - data: DataFrame, исходные данные
    - columns: список, столбцы, для которых нужно создать лаги
    - lag_count: int, количество лагов, которые необходимо создать

    Returns:
    - DataFrame, обновленные данные с добавленными лагами
    """

    columns_for_created_lags = columns

    # Проверка необходимости создания лага для предсказываемых значений
    if not need_create_lag_for_predictable:
        columns_for_created_lags.remove(chosen_column_for_predict)

    # Проходим по каждому столбцу
    for column in columns_for_created_lags:
        # Создаем лаги для каждого столбца
        for lag in range(1, lag_count + 1):
            # Формируем имя нового столбца с учетом лага
            lag_column_name = f'{column}_lag_{lag}'
            # Добавляем новый столбец с лагом
            data[lag_column_name] = data[column].shift(lag)

    # Заменяем пропущенные значения в данных на 0
    data = data.fillna(0)

    return data


def create_model(data, column_for_predict, column_factors, lag_count=1):
    """
    Создает и обучает модель на основе данных.

    Parameters:
    - data: DataFrame, исходные данные
    - column_for_predict: str, название столбца, который мы хотим предсказать
    - column_factors: список, столбцы-факторы для модели
    - lag_count: int, количество лагов для каждого фактора

    Returns:
    - model: обученная модель
    """
    # Создаем список для хранения факторов и их лагов
    all_factors = []

    # Добавляем факторы
    all_factors.extend(column_factors)

    # Добавляем колонны с лагами
    for column in column_factors:
        for lag in range(1, lag_count + 1):
            lag_column_name = f'{column}_lag_{lag}'
            all_factors.append(lag_column_name)

    # Удаляем столбец, который мы хотим предсказать
    if column_for_predict in all_factors:
        all_factors.remove(column_for_predict)

    # Выводим в консоль для проверки
    # print(all_factors)
    # print(column_for_predict)

    # Создание матрицы X и вектора Y
    X = data[all_factors]
    X = sm.add_constant(X)
    y = data[column_for_predict]
    # Оценка параметров с использованием МНК
    model = sm.OLS(y, X).fit()

    return model


def separation_data(data, percent):
    """
    Разделяет данные на обучающую и тестовую выборки.

    Parameters:
    - data: DataFrame. Данные для разделения.
    - percent: int. Процент данных, используемых для обучения (от 1 до 100).

    Returns:
    - Tuple. Две части данных: обучающая и тестовая выборки.
    """
    # Определите границу разделения данных
    split_index = int(len(data) * (percent / 100))

    # Разделите данные
    data_learn = data[:split_index]
    data_test = data[split_index:]

    return data_learn, data_test


def calculate_mape(df, actual_column):
    """
    Рассчитывает коэффициент MAPE (Mean Absolute Percentage Error) для прогноза и добавляет его в DataFrame.

    Parameters:
    - df: DataFrame. Исходные данные.
    - actual_column: str. Название столбца с фактическими значениями.
    - forecast_column: str. Название столбца с прогнозными значениями.

    Returns:
    - float. Среднее значение коэффициента MAPE.
    """

    forecast_column = f'Прогноз {actual_column}'

    # Рассчитываем MAPE для каждой строки в DataFrame
    df['MAPE'] = np.where(df[actual_column] != 0,
                          abs((df[actual_column] - df[forecast_column]) / df[actual_column]) * 100, 0)

    # Рассчитываем среднее значение MAPE
    average_mape = df['MAPE'].mean()

    return average_mape


def learn_on_params(data, params_train, len_dataset_learn, chosen_column_for_predict):
    """
    Прогнозирует значения на основе обученных параметров для заданного процента тестовых данных.

    Parameters:
    - data: DataFrame. Данные для прогноза.
    - params_train: dict. Обученные параметры, включая 'const' и другие параметры модели.
    - len_dataset_learn: int. Процент данных, используемых для обучения (от 1 до 100).

    Returns:
    - DataFrame. Данные с добавленным столбцом 'Прогноз' и меткой 'Тип данных'.
    """
    split_index = int(len(data) * (len_dataset_learn / 100))
    # Инициализация переменных
    forecast_values = []

    # Извлечение выбранных параметров из params_train
    const_param = params_train.get('const', 0)
    selected_params = list(params_train.keys())[1:]  # Исключаем 'const'

    # Прогнозирование для каждой строки в тестовой выборке
    for index, row in data.iterrows():
        # Прогнозирование с использованием выбранных параметров
        forecast = const_param + sum(params_train[param] * row[param] for param in selected_params)
        forecast_values.append(forecast)

    # Добавление прогнозных значений в DataFrame
    data[f'Прогноз {chosen_column_for_predict}'] = forecast_values

    # Присвоение значения для обучающей выборки
    data.loc[0:split_index - 1, 'Тип данных'] = 'Обучающая'

    # Присвоение значения для тестовой выборки
    data.loc[split_index:, 'Тип данных'] = 'Тестовая'

    # Преобразование столбца 'Тип данных' к строковому типу
    data['Тип данных'] = data['Тип данных'].astype(str)

    return data


def create_predict_one_day(data, params_train, chosen_column_for_predict):
    # Создаем копию последней строки
    new_row = data.iloc[-1].copy()

    # Извлекаем выбранные параметры из params_train
    const_param = params_train.get('const', 0)
    selected_params = list(params_train.keys())[1:]  # Исключаем 'const'

    # Прогнозирование для последней строки
    forecast = const_param + sum(params_train[param] * new_row[param] for param in selected_params)

    # Добавляем значение прогноза в столбец
    new_column_name = f'Прогноз {chosen_column_for_predict}'
    data[new_column_name].iloc[-1] = forecast

    # Маркируем тип данных как 'Прогноз 1-го дня'
    data['Тип данных'].iloc[-1] = 'Прогноз 1-го дня'

    # Маркировка столбцов, которые испльзкуются в качестве фактора в модели *название*'_P'
    data = rename_columns_with_suffix(data, selected_params)

    return data


# def create_predict_one_day(data, params_train, chosen_column_for_predict):
#     df = data.copy()  # Создаем копию DataFrame
#     df.loc[df.index[-1] + 1] = df.iloc[-1]  # Добавляем последнюю строку в конец
#     df.reset_index(drop=True, inplace=True)  # Переиндексируем DataFrame
#
#     # Извлекаем выбранные параметры из params_train
#     const_param = params_train.get('const', 0)
#     selected_params = list(params_train.keys())[1:]  # Исключаем 'const'
#
#     # Прогнозирование для последней строки
#     forecast = const_param + sum(params_train[param] * df.at[df.index[-1], param] for param in selected_params)
#
#     df[f'Прогноз {chosen_column_for_predict}'].iloc[-1] = forecast
#
#     # Маркировка столбцов, которые испльзкуются в качестве фактора в модели *название*'_P'
#     df = rename_columns_with_suffix(df, selected_params)
#
#     return df


def rename_columns_with_suffix(df, chosen_column):
    column_mapping = {column: column + "_P" for column in chosen_column}
    df.rename(columns=column_mapping, inplace=True)
    return df


def process_sheet_name(sheet_name):
    # Проверка на пустоту
    if not sheet_name:
        return "PredictionSheetName"

    # Проверка на длину
    if len(sheet_name) > 31:
        return "PredictionSheetName"

    # Замена запрещенных символов на "_"
    forbidden_characters = r'/\?*:[]'
    sheet_name = ''.join('_' if char in forbidden_characters else char for char in sheet_name)

    # Замена начала или конца на "_", если не были удалены апострофы
    sheet_name = sheet_name.strip("'") + '_'

    # Замена зарезервированного слова "History"
    if sheet_name.lower() == 'History':
        return 'PredictionSheetName'

    # Замена апострофов между символами на "_"
    sheet_name = sheet_name.replace("''", '_')

    # Прошли все проверки, возвращаем обработанное имя листа
    return sheet_name


def process_string_filename(input_string):
    # Заменяем указанные символы на нижнее подчеркивание
    invalid_chars = r'[\ / : * ? " < > | + . , _]'
    processed_string = re.sub(invalid_chars, '_', input_string)

    # Заменяем два подряд нижних подчеркивания на одно
    processed_string = re.sub(r'_{2,}', '_', processed_string)

    return processed_string


def write_to_excel(data, output_file, output_directory=None, file_format="xlsx", sheet="Prediction"):
    """
    Записывает данные в файл Excel или CSV.

    Parameters:
    - data: DataFrame. Данные для записи.
    - output_file: str. Имя файла (без расширения).
    - output_directory: str, optional. Директория для сохранения файла. Если не указана, используется текущая директория.
    - file_format: str, optional. Формат файла: "xlsx" (по умолчанию) или "csv".

    Returns:
    - result_write: str. Сообщение о результате операции записи.
    """

    try:
        # Обработка имени файла
        if not output_file.startswith('Prediction '):
            output_file = 'Prediction ' + output_file
        output_file = process_string_filename(output_file)

        # Обработка имени листа
        if not sheet.startswith('Prediction '):
            sheet = 'Prediction ' + sheet
        sheet = process_sheet_name(sheet)

        # Если не указана директория, используем текущую
        current_directory = output_directory or os.getcwd()

        # Создаем папку, если ее нет
        os.makedirs(current_directory, exist_ok=True)

        # Формируем имя файла с использованием префикса и текущей даты/времени
        current_datetime = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        file_name = f'{output_file}_{current_datetime}'

        # Выбираем формат файла
        file_extension = 'xlsx' if file_format.lower() == 'xlsx' else 'csv'

        # Составляем полный путь к файлу в текущей директории с учетом формата
        full_path = os.path.join(current_directory, f'{file_name}.{file_extension}')
        # Записываем данные в файл
        if file_extension == 'xlsx':
            data.to_excel(full_path, sheet_name=sheet, index=False)
        elif file_extension == 'csv':
            data.to_csv(full_path, sheet_name=sheet, index=False)

        result_write = dict()
        result_write['Result'] = True
        result_write['Path'] = full_path
    except Exception as e:
        result_write['Result'] = False

    return result_write
