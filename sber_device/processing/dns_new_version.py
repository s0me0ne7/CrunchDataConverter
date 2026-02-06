from pathlib import Path

import numpy as np
import pandas as pd
from utils.helpers import get_col_name_for_value
from config.configurations import RetailerConfig
from utils.io import read_excel

dns_retailer_config_new = RetailerConfig(
    company_name="ДНС",
    table_display_name="DNS_data",
    excel_table_style_name="TableStyleLight9",
    start_summation_row="E",
    table_header_style="Headline 1",
)


LIST_OF_VALID_COL_NAMES = [
    "Код",
    "Товар",
    "КодПроизводителя",
    "Кол-во",
    "Себестоимость без НДС",
    "Ост. Сумма без НДС",
    "Ост. пути",
    "Ост. кол-во",
    "Продажа без НДС",
]

temp_colnames = ["Артикул", "Наименование", "Код модели", "Магазин", "metric", "value"]

final_column_names = [
    "Код модели",
    "Наименование",
    "Артикул",
    "Магазин",
    "Продажи, шт",
    "Остатки, шт",
    "Остатки в пути",
]


def get_dns_data_new_format(data: pd.DataFrame) -> pd.DataFrame:
    sku_field_name = get_col_name_for_value(data, required_value="Товар")[0]
    data = data.dropna(subset=[sku_field_name]).dropna(how="all", axis=1)

    data = data.T.reset_index().map(
        lambda x: np.nan if ((isinstance(x, str)) and x.startswith("Unnamed")) else x
    )
    data = data.assign(index=data["index"].ffill())

    data = (
        data.loc[data[0].isin(LIST_OF_VALID_COL_NAMES)]
        .loc[data["index"] != "Итого"]
        .fillna(0)
        .set_index(["index", 0])
        .T.rename_axis([None, None], axis=1)
    )

    data = data.melt(id_vars=data.columns[:3].tolist())
    data.columns = temp_colnames

    data = (
        data.set_index(temp_colnames[:5])
        .unstack()
        .droplevel(level=0, axis=1)
        .reset_index()
        .rename_axis(None, axis=1)
        .rename(
            columns={
                "Кол-во": "Продажи, шт",
                "Ост. кол-во": "Остатки, шт",
                "Ост. пути": "Остатки в пути",
            }
        )
        .reindex(final_column_names, axis=1)
    )

    return data


def run_dns_new(path: Path | str, **kwargs) -> pd.DataFrame:
    data = read_excel(path)
    return get_dns_data_new_format(data)
