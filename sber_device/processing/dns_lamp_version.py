from pathlib import Path
from dataclasses import dataclass, InitVar, field

import numpy as np
import pandas as pd

from .dns_new_version import get_col_name_for_value
from .column_lists import ExtendedReportColumnLists
from config.configurations import RetailerConfig

dns_retailer_config_lamp_version = RetailerConfig(
    company_name="ДНС",
    table_display_name="DNS_data_extended",
    excel_table_style_name="TableStyleLight9",
    start_summation_row="F",
    table_header_style="Headline 1",
)


@dataclass
class ExtendedReportLampVersion(ExtendedReportColumnLists):
    df_path: InitVar[Path | str]
    df: pd.DataFrame | None = field(init=False, default=None)
    sku_field_name: str | None = field(init=False, default=None)

    def __post_init__(self, df_path: Path | str) -> None:
        self.df = pd.read_excel(df_path, header=0)
        if self.df is None:
            raise ValueError("Нет данных для обарботки")
        self.df.iloc[[0], :] = self.df.iloc[[0], :].infer_objects().ffill(axis=1)
        self.df.iloc[[1], :] = self.df.iloc[1, :].infer_objects().bfill()
        self.sku_field_name = get_col_name_for_value(self.df, required_value="Товар")[0]

    def _fill_nan_values(self) -> None:
        """
        Заполнение пропущенных значений и подготовка таблицы для трансформации в отчет.

        Returns
        -------
        None
        """
        self.df = self.df.dropna(subset=[self.sku_field_name]).dropna(how="all", axis=1)
        self.df = self.df.T.reset_index().map(
            lambda x: np.nan if ((isinstance(x, str)) and x.startswith("Unnamed")) else x
        )
        self.df = self.df.assign(index=self.df["index"].ffill())
        self.df = (
            self.df.loc[self.df.loc[:, 1].isin(self.LIST_OF_VALID_COLNAMES)]
            .loc[self.df["index"] != "Итого"]
            .set_index(["index", 0, 1])
            .T.rename_axis([None, None, None], axis=1)
            .dropna(how="all", axis=1)
            .infer_objects()
            .fillna(0.0)
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
            .fillna(0.0)
            .reindex(self.FINAL_COLNAMES, axis=1)
        )
        self.df.loc[:,  ["Продажи, шт", "Остатки, шт", "Остатки в пути"]] = (
            self.df.loc[:,  ["Продажи, шт", "Остатки, шт", "Остатки в пути"]].map(int)
        )
        self.df.loc[:, ["Код модели"]] = self.df.loc[:, ["Код модели"]].map(
            lambda x: "" if x == 0.0 else x
        )


def run_dns_lamp_version(path: Path | str) -> pd.DataFrame:
    """
    Обработка нового отчета с расширенным количеством полей (в версии "Лампы") (июль 2025)

    Parameters
    ----------
    path : Path
        Путь к файлу с отчетом.

    Return
    ------
    pd.DataFrame
        Отчет в формате pandas.DataFrame.
    """
    extended_processor = ExtendedReportLampVersion(path)
    extended_processor.create_report()
    return extended_processor.df