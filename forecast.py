import os
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
import datetime
import statsmodels
import statsmodels.api as sm
from statsmodels.tsa.stattools import adfuller
from colorama import Fore, Style
from sklearn.metrics import mean_squared_error
from scipy.interpolate import Rbf


def insert_into_cell_excel(title, sheet_name, cell_coordinates, cell_value):
    # Укажите путь к вашему файлу Excel
    excel_file_path = f'ADL/{title}.xlsx'

    # Открываем файл Excel
    workbook = openpyxl.load_workbook(excel_file_path)

    # Выбираем лист, на котором вы хотите внести изменения (укажите имя листа)
    sheet = workbook[sheet_name]

    # Записываем значение в ячейку
    sheet[cell_coordinates] = cell_value

    # Сохраняем изменения
    workbook.save(excel_file_path)

    # Закрываем файл Excel
    workbook.close()


def write_mape_rmse(mape, rmse, file_path):
    with open(file_path, 'a+') as file:
        file.write(f'\n\n\n MAPE - {mape}')
        file.write(f'\n\n\n RMSE - {rmse}')


def write_mape(mape, file_path):
    with open(file_path, 'a+') as file:
        file.write(f'\n\n\n MAPE - {mape}')


def write_parameters_to_file(model, file_path):
    params = model.params
    with open(file_path, 'w') as file:
        for key, value in params.items():
            if key == 'const':
                new_key = "a0"
            elif key == 'Энергопотребление_lag_1':
                new_key = "a1"
            elif key == 'Объем_работы':
                new_key = "b0"
            else:
                new_key = key  # если ключ не соответствует ни одному из вышеперечисленных
            file.write(f'{new_key}: {value}\n')
        file.write(f'\n\n\n{model.summary()}')


def calculate_mape(actual, forecast):
    # Предотвращение деления на ноль
    mask = actual != 0

    # Рассчитываем абсолютные процентные ошибки
    absolute_percentage_errors = np.abs((actual - forecast) / actual) * 100

    # Исключаем нулевые значения (предотвращение деления на ноль)
    absolute_percentage_errors = absolute_percentage_errors[mask]

    # Рассчитываем среднее значение абсолютных процентных ошибок
    mape = np.mean(absolute_percentage_errors)

    return mape


def percentage_error(data):
    column_name = 'MAPE'
    array = np.zeros((len(data) + 1, 1))
    df = pd.DataFrame(array, columns=[column_name])

    for j in range(1, len(data) + 1):
        subset = data.loc[j:j, ['Энергопотребление', 'Прогноз']].copy()
        if subset['Энергопотребление'].values[0] != 0:
            df.loc[j, column_name] = abs((subset['Энергопотребление'].values[0] - subset['Прогноз'].values[0]) /
                                         subset['Энергопотребление'].values[0])
        else:
            df.loc[j, column_name] = None
    return df


def MAPE_statsmodels(actual, predict):
    mape = statsmodels.tools.eval_measures.meanabs(actual, predict, axis=0)
    return mape


def predict_on_params(data_main, params_train, title, len_dataset_learn):
    # Инициализация переменных
    forecast_values = []
    # Прогнозирование для каждой строки в тестовой выборке
    for index, row in data_main.iterrows():
        forecast = params_train['const'] + params_train['Энергопотребление_lag_1'] * row['Энергопотребление'] + \
                   params_train['Объем_работы'] * row['Объем_работы']
        forecast_values.append(forecast)
    # Добавление прогнозных значений в DataFrame
    data_main.loc[:, 'Прогноз'] = forecast_values
    # Создание метки, где обучающая выборка, где тестовая
    data_main.loc[0:len_dataset_learn - 1, 'Тип данных'] = 'Обучающая'

    # Присвоение значения для оставшейся части данных (test_data)
    data_main.loc[len_dataset_learn:, 'Тип данных'] = 'Тестовая'
    fig = go.Figure()
    # Добавляем линию к графику
    fig.add_trace(go.Scatter(y=data_main['Энергопотребление'], mode='lines', name='Фактическое значение'))
    fig.add_trace(go.Scatter(y=data_main['Прогноз'], mode='lines', name='Прогноз'))

    # Настройки макета графика
    fig.update_layout(title=title,
                      xaxis_title='Ось X',
                      yaxis_title='Ось Y')

    fig.write_image(f"adl/{title}.png", width=1200)

    # Отображаем график
    fig.show()
    return data_main


def estimate_parameters(data):
    # Преобразование столбцов в числовой тип данных
    data.loc[:, 'Энергопотребление'] = pd.to_numeric(data['Энергопотребление'], errors='coerce')
    data.loc[:, 'Объем_работы'] = pd.to_numeric(data['Объем_работы'], errors='coerce')

    # Создание лагов энергопотребления
    data['Энергопотребление_lag_1'] = data['Энергопотребление'].shift(1)
    data = data.fillna(0)

    # Создание матрицы X и вектора Y
    X = data[['Энергопотребление_lag_1', 'Объем_работы']]
    X = sm.add_constant(X)
    y = data['Энергопотребление']

    # Оценка параметров с использованием МНК
    model = sm.OLS(y, X).fit()

    # Вывод статистической информации о модели
    # print(model.summary())

    return model


def separation_data(data):
    # Разделение на обучающую и тестовую выборку
    len_learn_data = int((2 * len(data) - 1) / 3)
    len_test_data = int((1 * len(data) - 1) / 3)

    data_train = data[0:len_learn_data]
    data_test = data[len_learn_data:len_learn_data + len_test_data]

    return data_train, data_test


def create_path_output_files():
    # Создание директории сохранения файлов
    if not os.path.exists("adl"):
        os.mkdir("adl")
        print("Created 'adl' directory.")


def main_func(data, title):
    title_type_learn = 'Обучающая'
    title_type_test = 'Тестовая'

    # Разделяем данные на 2 выборки в соотношении 1\3 и 2\3
    data_learn = separation_data(data)[0]
    data_test = separation_data(data)[1]

    # Находим коэффициенты a0, a1, b0 на обучающей выборке
    model = estimate_parameters(data_learn)
    write_parameters_to_file(model, f"ADL/{title}.txt")
    print(f"Оцененные параметры для data_learn:", model.params)

    # Создание прогноза на обучающей выборке по коэффициентам
    data_learn_with_predict = predict_on_params(data, model.params, title, len(data_learn)).fillna(0)

    # Расчёт RMSE и MAPE
    # rmse = np.sqrt(mean_squared_error(data['Энергопотребление'], data['Прогноз']))
    mape = percentage_error(data)
    data.loc[:, 'MAPE'] = mape
    avg_mape = np.mean(data['MAPE'])
    avg_mape_learn = np.mean(data['MAPE'][0:len(data_learn) - 1])
    avg_mape_test = np.mean(data['MAPE'][len(data_learn):])

    # Запись в файлы расчитанной информации
    write_mape(avg_mape, f"ADL/{title}.txt")
    # print(f"{Fore.GREEN}\n\n\n Среднее MAPE для {title} - |{avg_mape}|\n\n\n RMSE - |{rmse}|{Style.RESET_ALL}")
    print(f"{Fore.GREEN}\n\n\n Среднее MAPE для {title} - |{avg_mape}|{Style.RESET_ALL}")
    print(
        f"{Fore.GREEN}\n\n Среднее MAPE обучающей выборки для {title} - |{avg_mape_learn}|\n\n\n Среднее MAPE тестовой выборки для {title} - |{avg_mape_test}|\n\n\n{Style.RESET_ALL}")
    data.to_excel(f'ADL/{title}.xlsx', index=False)

    # Запись MAPE
    sheet_name = 'Sheet1'
    insert_into_cell_excel(title, sheet_name, 'G1', 'Среднее значение MAPE обучающей выборки')
    insert_into_cell_excel(title, sheet_name, 'H1', 'Среднее значение MAPE тестовой выборки')

    insert_into_cell_excel(title, sheet_name, 'G2', avg_mape_learn)
    insert_into_cell_excel(title, sheet_name, 'H2', avg_mape_test)


def main_import_file(file_path, sheet_name):
    create_path_output_files()
    # excel_file = 'Бестранспортная вскрыша_Сутки.xlsx'
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    data.drop(columns=['T'], inplace=True)
    data.drop(columns=['t'], inplace=True)
    data_esh_19 = data.copy()
    data_esh_19.drop(columns=['Объем работы, м3'], inplace=True)
    data_esh_19.drop(columns=['Энергопотребление, кВт*ч'], inplace=True)
    data_esh_29 = data_esh_19.copy()
    data_esh_29.drop(columns=['Объем работы, м3.1'], inplace=True)
    data_esh_29.drop(columns=['Энергопотребление, кВт*ч.1'], inplace=True)
    data_esh_19.drop(columns=['Объем работы, м3.2'], inplace=True)
    data_esh_19.drop(columns=['Энергопотребление, кВт*ч.2'], inplace=True)
    data_esh_19 = data_esh_19.rename(
        columns={"Объем работы, м3.1": "Объем_работы", "Энергопотребление, кВт*ч.1": "Энергопотребление"})
    data_esh_29 = data_esh_29.rename(
        columns={"Объем работы, м3.2": "Объем_работы", "Энергопотребление, кВт*ч.2": "Энергопотребление"})
    data_esh_29 = data_esh_29.dropna()
    data_esh_19 = data_esh_19.dropna()
    data_esh_29['Энергопотребление'] = pd.to_numeric(data_esh_29['Энергопотребление'], errors='coerce')
    data_esh_29['Объем_работы'] = pd.to_numeric(data_esh_29['Объем_работы'], errors='coerce')
    data_esh_19['Энергопотребление'] = pd.to_numeric(data_esh_19['Энергопотребление'], errors='coerce')
    data_esh_19['Объем_работы'] = pd.to_numeric(data_esh_19['Объем_работы'], errors='coerce')
    data_esh_29 = data_esh_29.replace([np.inf, -np.inf], np.nan).dropna()
    data_esh_19 = data_esh_19.replace([np.inf, -np.inf], np.nan).dropna()
    main_func(data_esh_19, 'ЭШ-20-90 №19')
    main_func(data_esh_29, 'ЭШ-20-90 №29')
