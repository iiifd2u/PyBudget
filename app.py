import os
import pandas as pd
from typing import List, Union, Optional
import argparse

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

days_dict = {"январь": 31,
             "февраль": 28,
             "март": 31,
             "апрель": 30,
             "май": 31,
             "июнь": 30,
             "июль": 31,
             "август": 31,
             "сентябрь": 30,
             "октябрь": 31,
             "ноябрь": 30,
             "декабрь": 31}


def create_df_from_excel(xl: pd.ExcelFile, sheet: str, columns_to_rename: Optional[List[str]]):
    # iloc - для индексов [колонки, столбцы]
    df1 = xl.parse(sheet)

    if columns_to_rename is not None:
        old_keys = list(df1.keys())
        new_keys = list(df1.iloc[0, :])

        columns = {o: n for o, n in zip(old_keys, new_keys)}
        data1 = df1.iloc[1:, :].rename(columns=columns)
        for column in columns_to_rename:
            data1[column] = data1[column].astype("int64")

        return data1
    else:
        return df1


def write_to_excel(df: pd.DataFrame, path: str, startrow:int, startcol:int):

    indices = ['Кафе', 'Продукты', 'Транспорт',
               'Спорт', 'Туризм', 'Аптека',
               'Сумма по еде в месяц', 'Итог по еде в день',
               'Сумма всего', 'В день', 'Доход', 'Разница']

    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Отчётность', startcol=startcol, startrow=startrow)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-F", "--file", help="local excel file")
    args = parser.parse_args()
    file = args.file


    month = os.path.basename(file).split('_')[0].lower()
    print(month)

    xl = pd.ExcelFile(file)

    sheets = xl.sheet_names

    df_income = create_df_from_excel(xl, sheets[0],
                                     columns_to_rename=["Сумма в валюте учета", "Сумма в валюте счета"])  # ДОХОДЫ
    # print(df_income)

    df_expenses = create_df_from_excel(xl, sheets[1],
                                       columns_to_rename=["Сумма в валюте учета", "Сумма в валюте счета"])  # РАСХОДЫ
    # print(df_expenses)

    df_remittance = create_df_from_excel(xl, sheets[2],
                                         columns_to_rename=["Сумма в исходящей валюте счета",
                                                            "Сумма во входящей валюте счета"])  # РАСХОДЫ

    # print(df_minus.loc[:, 'Категория'])
    # exit()

    # Продукты
    food_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Продукты"].sum()
    # Кафе
    cafe_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Кафе"].sum()
    # Транспорт
    transport_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Транспорт"].sum()
    # Аптека
    pharmacy_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Аптека"].sum()
    # Спорт
    sport_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[
        (df_expenses.loc[:, 'Категория'] == "Спорт") | (df_expenses.loc[:, 'Категория'] == "Тренировки")].sum()
    # Туризм
    tourism_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Туризм"].sum()

    # Внутренние переводы
    remittance_price = df_remittance.loc[:, "Сумма в исходящей валюте счета"].sum()
    # Вся сумма:
    all_price = df_expenses.loc[:, "Сумма в валюте учета"].sum()

    #Доход
    all_income = df_income.loc[:, "Сумма в валюте учета"].sum()


    all_food_month = food_price+ cafe_price

    all_food_per_day = all_food_month / days_dict[month]


    all_price_per_day = (all_price-remittance_price)/days_dict[month]

    indices = ['Кафе', 'Продукты', 'Транспорт',
               'Спорт', 'Туризм', 'Аптека', 'Внутренние переводы',
               'Сумма по еде в месяц', 'Итог по еде в день',
               'Сумма всего', 'В день', 'Доход', 'Разница']

    prices = [cafe_price, food_price, transport_price,
              sport_price, tourism_price, pharmacy_price, remittance_price,
              all_food_month, all_food_per_day, all_price-remittance_price,
              all_price_per_day, all_income, all_income - all_price]

    # df_RPM = pd.DataFrame(data=prices, index=indices, columns=[month])
    # print(pd.Series(index=indices, data=prices))

    out_path = os.path.join("output_file", "records_2024.xlsx")

    xl = pd.ExcelFile(out_path)
    sheets = xl.sheet_names
    df_RPM = create_df_from_excel(xl, sheet=sheets[0], columns_to_rename=None).iloc[2:, 1:]
    months = df_RPM.iloc[0, :].tolist()
    print(months)
    print("------------")
    columns = {o: n for o, n in zip(list(df_RPM.keys()), months)}
    index = {o+1: n for o, n in zip(list(df_RPM.index), indices)}
    print(index)
    # exit()
    df_RPM = df_RPM.iloc[1:, :].rename(columns=columns, index = index)

    print("From excel:", df_RPM)

    if month in months:
        key = months.index(month)
        df_RPM[month] = df_RPM[month] + pd.Series(index=indices, data=prices)
        # write_to_excel(df_RPM, out_path, startrow = 10, startcol = key)
    else:
        key = len(months)
        print("from 2")
        print(key)
        # df_RPM[month] = pd.Series(index=indices, data=prices)
        # write_to_excel(df_RPM, out_path, startrow=3, startcol=key)

    print("To excel:", df_RPM)







# print(df1.loc['Сумма в валюте учета'])
