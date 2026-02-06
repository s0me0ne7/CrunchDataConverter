from pathlib import Path
from dataclasses import dataclass, InitVar, field
from typing import ClassVar

import numpy as np
import pandas as pd

from .dns_new_version import get_col_name_for_value
from config.configurations import RetailerConfig

dns_retailer_config_extended = RetailerConfig(
    company_name="ДНС",
    table_display_name="DNS_data_extended",
    excel_table_style_name="TableStyleLight9",
    start_summation_row="F",
    table_header_style="Headline 1",
)


@dataclass
class ExtendedReport:
    df_path: InitVar[Path | str]
    df: pd.DataFrame | None = None
    sku_field_name: str | None = field(init=False, default=None)

    LIST_OF_VALID_COLNAMES: ClassVar[list[str]] = [
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

    TMP_COLNAMES: ClassVar[list[str]] = [
        "Артикул",
        "Наименование",
        "Код модели",
        "Магазин",
        "Код магазина",
        "metric",
        "value",
    ]

    FINAL_COLNAMES: ClassVar[list[str]] = [
        "Код модели",
        "Наименование",
        "Артикул",
        "Магазин",
        "Код магазина",
        "Продажи, шт",
        "Остатки, шт",
        "Остатки в пути",
        "Себестоимость без НДС",
        "Ост. Сумма без НДС",
        "Продажа без НДС",
    ]

    def __post_init__(self, df_path: Path | str) -> None:
        self.df = pd.read_excel(df_path, header=0)
        if self.df is None:
            raise ValueError("Нет данных для обработки")
        self.sku_field_name = get_col_name_for_value(self.df, required_value="Товар")[0]

    def _fill_nan_values(self) -> None:
        """
        Заполнение пропущенных значений и подготовка таблицы для трансформации в отчет.

        Returns
        -------
        None
        """
        self.df.iloc[[0], :] = self.df.iloc[[0], :].infer_objects().ffill(axis=1)
        self.df = self.df.dropna(subset=[self.sku_field_name]).dropna(how="all", axis=1)
        self.df = self.df.T.reset_index().map(
            lambda x: np.nan
            if ((isinstance(x, str)) and x.startswith("Unnamed"))
            else x
        )
        self.df = self.df.assign(index=self.df["index"].ffill())
        self.df = (
            self.df.loc[self.df.loc[:, 1].isin(self.LIST_OF_VALID_COLNAMES)]
            .loc[self.df["index"] != "Итого"]
            .fillna(0.0)
            .set_index(["index", 0, 1])
            .T.rename_axis([None, None, None], axis=1)
        )

    def create_report(self) -> None:
        """
        Сборка отчета и сохранение его в атрибуте класса df.

        Returns
        -------
        None
        """
        self._fill_nan_values()
        self.df = self.df.melt(id_vars=self.df.columns[:3].tolist())
        self.df.columns = self.TMP_COLNAMES
        self.df = (
            self.df.set_index(self.TMP_COLNAMES[:6])
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
            .reindex(self.FINAL_COLNAMES, axis=1)
        )
        self.df.loc[:, ["Продажи, шт", "Остатки, шт", "Остатки в пути"]] = self.df.loc[
            :, ["Продажи, шт", "Остатки, шт", "Остатки в пути"]
        ].astype(int)


def run_dns_extended(path: Path | str) -> pd.DataFrame:
    """
    Обработка нового отчета с расширенным количеством полей (июль 2025)

    Parameters
    ----------
    path : Path
        Путь к файлу с отчетом.

    Return
    ------
    pd.DataFrame
        Отчет в формате pandas.DataFrame.
    """
    extended_processor = ExtendedReport(path)
    extended_processor.create_report()
    return extended_processor.df
