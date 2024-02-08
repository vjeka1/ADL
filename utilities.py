import numpy as np
import pandas as pd


def data_preparation(file, sheet, name_column_time, name_column_factors):
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

    # Удаление строк, где значение в столбце name_column_time отличается от дд.мм.гггг
    result_df[name_column_time] = pd.to_datetime(result_df[name_column_time], format='%d.%m.%Y', errors='coerce')
    result_df = result_df.dropna(subset=[name_column_time])

    # Вывод полученного DataFrame (можно удалить)
    print(result_df)

    return result_df


def create_lags(data, columns, lag_count):
    for column in columns:
        for lag in range(1, lag_count + 1):
            lag_column_name = f'{column}_lag_{lag}'
            data[lag_column_name] = data[column].shift(lag)
    data = data.fillna(0)
    print(data)
    return data
