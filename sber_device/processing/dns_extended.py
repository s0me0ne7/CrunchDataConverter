from pathlib import Path
from dataclasses import dataclass, InitVar, field
from typing import ClassVar

import numpy as np
import pandas as pd

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
    """
    Обработчик расширенного отчёта ДНС.

    Структура входного файла:
    - Row 0: Названия групп (Изделие, Итого, названия магазинов)
    - Row 1: Метки полей (Код, Товар, КодПроизводителя) + коды магазинов
    - Row 2: Названия метрик (повторяются для каждого магазина)
    - Row 3: Строка итогов по категории
    - Row 4+: Данные товаров

    Колонки:
    - Col 0: Код товара
    - Col 2: Название товара
    - Col 4: Артикул
    - Col 5+: Блоки по 6 колонок для каждого магазина (метрики)
    """

    df_path: InitVar[Path | str]
    df: pd.DataFrame | None = field(init=False, default=None)
    raw_df: pd.DataFrame | None = field(init=False, default=None)

    # Названия метрик в исходном файле (порядок важен - соответствует порядку колонок)
    METRICS: ClassVar[list[str]] = [
        "Кол-во",
        "Себестоимость без НДС",
        "Ост. Сумма без НДС",
        "Ост. пути",
        "Ост. кол-во",
        "Продажа без НДС",
    ]

    # Маппинг метрик на выходные названия
    METRIC_MAPPING: ClassVar[dict[str, str]] = {
        "Кол-во": "Продажи, шт",
        "Ост. кол-во": "Остатки, шт",
        "Ост. пути": "Остатки в пути",
    }

    # Целочисленные поля
    INTEGER_FIELDS: ClassVar[list[str]] = [
        "Продажи, шт",
        "Остатки, шт",
        "Остатки в пути",
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
        # Читаем без заголовка чтобы сохранить структуру
        self.raw_df = pd.read_excel(df_path, header=None, dtype=object)
        if self.raw_df is None or self.raw_df.empty:
            raise ValueError("Нет данных для обработки")

    def _find_product_columns(self) -> dict[str, int]:
        """
        Найти колонки с идентификаторами товара по значениям в Row 1.

        Returns
        -------
        dict[str, int]
            Словарь {название_поля: индекс_колонки}
        """
        row1 = self.raw_df.iloc[1]
        product_cols = {}

        for col_idx, value in enumerate(row1):
            if value == "Код":
                product_cols["Код"] = col_idx
            elif value == "Товар":
                product_cols["Товар"] = col_idx
            elif value == "КодПроизводителя":
                product_cols["КодПроизводителя"] = col_idx

        required = {"Код", "Товар", "КодПроизводителя"}
        missing = required - set(product_cols.keys())
        if missing:
            raise KeyError(f"Не найдены обязательные поля в строке 1: {missing}")

        return product_cols

    def _find_shop_blocks(self) -> list[dict]:
        """
        Найти блоки колонок для каждого магазина.

        Returns
        -------
        list[dict]
            Список словарей с информацией о магазинах:
            [{'name': 'Магазин', 'code': '5645', 'start_col': 11}, ...]
        """
        row0 = self.raw_df.iloc[0]  # Названия магазинов
        row1 = self.raw_df.iloc[1]  # Коды магазинов
        row2 = self.raw_df.iloc[2]  # Метрики

        shops = []
        metrics_count = len(self.METRICS)

        # Найти первую колонку с метриками (где Row 2 == "Кол-во")
        first_metric_col = None
        for col_idx, value in enumerate(row2):
            if value == "Кол-во":
                first_metric_col = col_idx
                break

        if first_metric_col is None:
            raise KeyError("Не найдена колонка с метрикой 'Кол-во'")

        # Пропускаем блок "Итого" (первый блок метрик)
        current_col = first_metric_col + metrics_count

        while current_col < len(row0):
            shop_name = row0.iloc[current_col]
            shop_code = row1.iloc[current_col]

            # Проверяем что это действительно магазин (есть название)
            if pd.notna(shop_name) and shop_name != "Итого":
                shops.append({
                    "name": str(shop_name),
                    "code": str(shop_code) if pd.notna(shop_code) else "",
                    "start_col": current_col,
                })

            current_col += metrics_count

        return shops

    def _get_product_data(self, product_cols: dict[str, int]) -> pd.DataFrame:
        """
        Извлечь данные товаров (без метрик).

        Parameters
        ----------
        product_cols : dict[str, int]
            Индексы колонок с идентификаторами товара.

        Returns
        -------
        pd.DataFrame
            DataFrame с колонками [Код модели, Наименование, Артикул]
        """
        # Данные начинаются с Row 4 (Row 3 - это итоги по категории)
        data_start_row = 4

        # Маппинг: Код модели = КодПроизводителя, Артикул = Код (внутренний код ДНС)
        products = pd.DataFrame({
            "Код модели": self.raw_df.iloc[data_start_row:, product_cols["КодПроизводителя"]],
            "Наименование": self.raw_df.iloc[data_start_row:, product_cols["Товар"]],
            "Артикул": self.raw_df.iloc[data_start_row:, product_cols["Код"]],
        }).reset_index(drop=True)

        return products

    def create_report(self) -> None:
        """
        Сборка отчёта.
        """
        product_cols = self._find_product_columns()
        shops = self._find_shop_blocks()
        products = self._get_product_data(product_cols)

        if not shops:
            raise ValueError("Не найдено ни одного магазина в данных")

        data_start_row = 4
        metrics_count = len(self.METRICS)
        all_rows = []

        # Для каждого товара и каждого магазина создаём строку
        for prod_idx in range(len(products)):
            raw_row_idx = data_start_row + prod_idx

            for shop in shops:
                row_data = {
                    "Код модели": products.iloc[prod_idx]["Код модели"],
                    "Наименование": products.iloc[prod_idx]["Наименование"],
                    "Артикул": products.iloc[prod_idx]["Артикул"],
                    "Магазин": shop["name"],
                    "Код магазина": shop["code"],
                }

                # Извлекаем метрики для этого магазина
                for metric_idx, metric_name in enumerate(self.METRICS):
                    col_idx = shop["start_col"] + metric_idx
                    value = self.raw_df.iloc[raw_row_idx, col_idx]

                    # Применяем маппинг названий
                    output_name = self.METRIC_MAPPING.get(metric_name, metric_name)
                    row_data[output_name] = value if pd.notna(value) else 0

                all_rows.append(row_data)

        self.df = pd.DataFrame(all_rows)

        # Преобразование типов
        for field in self.INTEGER_FIELDS:
            if field in self.df.columns:
                self.df[field] = pd.to_numeric(self.df[field], errors="coerce").fillna(0).astype(int)

        # Числовые поля
        numeric_fields = ["Себестоимость без НДС", "Ост. Сумма без НДС", "Продажа без НДС"]
        for field in numeric_fields:
            if field in self.df.columns:
                self.df[field] = pd.to_numeric(self.df[field], errors="coerce").fillna(0)

        # Упорядочиваем колонки
        self.df = self.df.reindex(columns=self.FINAL_COLNAMES)

        # Сортировка по Код модели, затем по Магазин (как в старой версии)
        self.df = self.df.sort_values(
            by=["Код модели", "Магазин"],
            key=lambda x: x.astype(str)
        ).reset_index(drop=True)


def run_dns_extended(path: Path | str) -> pd.DataFrame:
    """
    Обработка расширенного отчёта ДНС.

    Parameters
    ----------
    path : Path | str
        Путь к файлу с отчётом.

    Returns
    -------
    pd.DataFrame
        Отчёт в формате pandas.DataFrame.
    """
    processor = ExtendedReport(path)
    processor.create_report()
    return processor.df