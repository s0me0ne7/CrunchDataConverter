from pathlib import Path

import pandas as pd

from utils.io import read_excel
from config.configurations import RetailerConfig


mvm_retailer_config = RetailerConfig(
    company_name="МВМ",
    table_display_name="MVM_data",
    excel_table_style_name="TableStyleLight10",
    start_summation_row="E",
    table_header_style="Headline 1",
)


def get_mvm_data(data: pd.DataFrame) -> pd.DataFrame:
    # add non-NA value in the columns with item names so that this is not deleting during dropna action
    data.iloc[0, 6] = "city"
    data.iloc[1, 6] = "code"

    data = data.drop(columns=data.iloc[:, 0:5].columns).dropna(subset=["Unnamed: 6"])
    cols_to_drop = data.iloc[:, 2:28].columns
    data = data.drop(columns=cols_to_drop)
    data = data.rename(columns={"Unnamed: 5": "model_code", "Unnamed: 6": "model"})

    data.loc[0] = data.loc[0].ffill()
    data.loc[1] = data.loc[1].ffill()
    data = data.reset_index(drop=True)

    # Encode model_code into model for uniqueness during transpose
    data.loc[3:, "model"] = (
        data.loc[3:, "model_code"].astype(str) + "|||" + data.loc[3:, "model"].astype(str)
    )

    data = data.drop(columns="model_code")
    data = data.set_index(["model"]).T.reset_index(drop=True).fillna(0)

    # melting
    data = data.melt(id_vars=["city", "code", "Наименование"])

    # Split model back into model_code and model name
    data[["model_code", "model"]] = data["model"].str.split("|||", n=1, expand=True)

    data = (
        data.groupby(["model_code", "model", "city", "code", "Наименование"], as_index=False)
        .agg({"value": "sum"})
    )
    data = (
        data.set_index(["model_code", "model", "city", "code", "Наименование"])
        .unstack(4)
        .reset_index()
    )
    # Flatten multi-level columns after unstack
    data.columns = [col[1] if col[1] else col[0] for col in data.columns]

    col_names = [
        "Артикул",
        "Наименование",
        "Город",
        "Код магазина",
        "Остатки, шт",
        "Остатки, руб",
        "Продажи, руб",
        "Продажи, шт",
    ]
    data.columns = col_names

    # filter non zero values and re-arrange columns
    data = data.loc[
        lambda r: (r["Остатки, шт"] != 0)
        | (r["Остатки, руб"] != 0)
        | (r["Продажи, руб"] != 0)
        | (r["Продажи, шт"] != 0)
    ]
    data = data.iloc[:, [0, 1, 2, 3, 7, 6, 4, 5]].reset_index(drop=True)

    return data


def run_mvm(path: Path | str, header=2):
    """
    Вспомогательная функция для учета нового формата в исходных таблицах МВМ.
    Сначала запускается первоначальный вариант функции, и в случае ошибки
    запускается вариант, заточенный под новую структуру (просто добавились
    пустые строки в начале файла)

    Parameters
    ----------
    path : Path | str
        Путь к файлу для загрузки данных
    header : int
        Номер строки, где начинается заголовок таблицы
        По умолчанию 2 (новый формат отчета)
    """
    try:
        data = read_excel(path, header=header)
    except KeyError as e:
        print(e)
        data = read_excel(path, header=0)

    # Обрезаем фильтровые строки: ищем первую строку с данными в per-store колонках
    # (col 6 = NaN, но колонки 33+ содержат данные — это строка с городами)
    for i in range(len(data)):
        if pd.isna(data.iloc[i, 6]) and data.iloc[i, 33:].notna().any():
            data = data.iloc[i:].reset_index(drop=True)
            break

    return get_mvm_data(data=data)
