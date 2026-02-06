import pandas as pd
import numpy as np
from pathlib import Path

from utils.io import read_excel
from config.configurations import RetailerConfig

dns_retailer_config_old = RetailerConfig(
    company_name="ДНС",
    table_display_name="DNS_data",
    excel_table_style_name="TableStyleLight9",
    start_summation_row="F",
    table_header_style="Headline 1",
)


#  Наименования столбцов при обработке
column_names = ["model", "shop", "shop_code", "sales", "stock_pcs", "stock_trans"]

# Наименования полей на выходе
colnames_mapper = {
    "model_code": "Код модели",
    "model": "Наименование",
    "manuf_code": "Артикул",
    "shop": "Магазин",
    "shop_code": "Код магазина",
    "sales": "Продажи, шт",
    "stock_pcs": "Остатки, шт",
    "stock_trans": "Остатки, в пути",
}


def get_dns_data_old_format(data: pd.DataFrame) -> pd.DataFrame:
    """
    Обработка данных от ДНС.

    Параметры:
    ----------
    data : pd.DataFrame
        Входные данные для обработки.
    colnames : list[str]
        Наименования полей таблицы, используемые в процессе обработки
    ru_colnames_mapper : dict[str, str]
        Словарь для переименования полей финальной таблицы
    """
    data = data.drop(columns=["Итого", "Unnamed: 4", "Unnamed: 5"])
    data.iloc[0:2, 1] = data.iloc[0:2, 1].fillna("donotdel")
    data = data.dropna(subset=["Unnamed: 1"]).rename(
        columns={
            "Unnamed: 1": "model",
            "Изделие": "model_code",
            "Unnamed: 2": "manuf_code",
        }
    )
    data = (
        data.set_index(["model", "model_code", "manuf_code"])
        .T.reset_index()
        .rename(columns={"index": "shop"})
    )
    data.iloc[:, 0] = (
        data.iloc[:, 0]
        .apply(lambda x: np.nan if x.startswith("Unnamed") else x)
        .ffill()
    )
    data.iloc[:, 1] = data.iloc[:, 1].ffill()
    data = data.fillna(0)

    cols_to_id = data.columns[:3].to_list()

    # melt down!
    data = data.melt(id_vars=cols_to_id)

    tmp_cols = [
        "shop",
        "shop_code",
        "stat",
        "model",
        "model_code",
        "manuf_code",
        "value",
    ]
    data.columns = tmp_cols

    data = (
        data.set_index(
            ["manuf_code", "model", "model_code", "shop", "shop_code", "stat"]
        )
        .unstack(5)
        .droplevel(axis=1, level=0)
        .reset_index()
    )

    final_column_names = [
        "Код модели",
        "Наименование",
        "Артикул",
        "Магазин",
        "Код магазина",
        "Продажи, шт",
        "Остатки, шт",
        "Остатки в пути",
    ]

    data.columns = final_column_names
    data = data.loc[
        lambda r: (r["Продажи, шт"] != 0)
        | (r["Остатки, шт"] != 0)
        | (r["Остатки в пути"] != 0)
    ].reset_index(drop=True)

    return data


def run_dns_old(path: Path | str, **kwargs) -> pd.DataFrame:
    data = read_excel(path)
    return get_dns_data_old_format(data)
