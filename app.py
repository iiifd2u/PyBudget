import os
import pandas as pd
from typing import List
import argparse

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)



days_dict = {"январь":31,
             "февраль":28,
             "март":31,
             "апрель":30,
             "май":31,
             "июнь":30,
             "июль":31,
             "август":31,
             "сентябрь":30,
             "октябрь":31,
             "ноябрь":30,
             "декабрь":31}

def create_df_from_excel(xl:pd.ExcelFile, sheet:str, columns_to_rename:List[str]):

    # iloc - для индексов [колонки, столбцы]
    df1 = xl.parse(sheet)
    old_keys = list(df1.keys())
    new_keys = list(df1.iloc[0, :])

    columns = {o: n for o, n in zip(old_keys, new_keys)}
    data1 = df1.iloc[1:, :].rename(columns=columns)

    data1[columns_to_rename[0]] = data1[columns_to_rename[0]].astype("int64")
    data1[columns_to_rename[1]] = data1[columns_to_rename[1]].astype("int64")

    return data1



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
                                     columns_to_rename=["Сумма в валюте учета", "Сумма в валюте счета"]) # ДОХОДЫ
    print(df_income)

    df_expenses = create_df_from_excel(xl, sheets[1],
                                       columns_to_rename=["Сумма в валюте учета", "Сумма в валюте счета"]) # РАСХОДЫ
    print(df_expenses)

    df_remittance = create_df_from_excel(xl, sheets[2],
                                         columns_to_rename=["Сумма в исходящей валюте счета", "Сумма во входящей валюте счета"]) # РАСХОДЫ
    print(df_remittance)

    # print(df_minus.loc[:, 'Категория'])
    # exit()

    # Продукты
    food_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Продукты"]
    # Кафе
    cafe_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Кафе"]
    # Транспорт
    transport_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Транспорт"]
    # Аптека
    pharmacy_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Аптека"]
    # Спорт
    sport_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[(df_expenses.loc[:, 'Категория'] == "Спорт") | (df_expenses.loc[:, 'Категория'] == "Тренировки")]
    # Туризм
    tourism_price = df_expenses.loc[:, "Сумма в валюте учета"].loc[df_expenses.loc[:, 'Категория'] == "Туризм"]

    print(food_price.sum())
    print(cafe_price.sum())

    all_food_month = food_price.sum()+cafe_price.sum()
    print(all_food_month)

    all_food_per_day = all_food_month/days_dict[month]
    print(all_food_per_day)

    print(sport_price.sum())


# print(df1.loc['Сумма в валюте учета'])