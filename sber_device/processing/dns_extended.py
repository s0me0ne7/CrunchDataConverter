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
    - Row 0: Названия групп (Изделие, Итого, названия магазинов — объединённые ячейки)
    - Row 1: Метки полей (Код, Товар, КодПроизводителя) + коды магазинов
    - Row 2: Названия метрик (повторяются для каждого магазина)
    - Row 3: Строка итогов по категории
    - Row 4+: Данные товаров

    Блоки магазинов имеют переменный размер (объединённые ячейки),
    поэтому поиск магазинов выполняется динамически по Row 0.
    """

    df_path: InitVar[Path | str]
    df: pd.DataFrame | None = field(init=False, default=None)
    raw_df: pd.DataFrame | None = field(init=False, default=None)

    # Маппинг метрик на выходные названия
    METRIC_MAPPING: ClassVar[dict[str, str]] = {
        "Кол-во": "Продажи, шт",
        "Ост. кол-во": "Остатки, шт",
        "Ост. пути": "Остатки в пути",
    }

    # Целочисленные поля
    INTEGER_FIELDS: ClassVar[set[str]] = {
        "Продажи, шт",
        "Остатки, шт",
        "Остатки в пути",
    }

    # Базовые колонки (всегда присутствуют)
    BASE_COLNAMES: ClassVar[list[str]] = [
        "Код модели",
        "Наименование",
        "Артикул",
        "Магазин",
        "Код магазина",
    ]

    # Порядок метрик в выходном файле (максимальный набор)
    METRIC_ORDER: ClassVar[list[str]] = [
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

        Магазины определяются динамически по непустым ячейкам в Row 0.
        Размер блока (количество метрик) может отличаться от запуска к запуску.
        Для каждого магазина находим фактические колонки метрик из Row 2.

        Returns
        -------
        list[dict]
            [{'name': str, 'code': str, 'metrics': {metric_name: col_idx}}]
        """
        row0 = self.raw_df.iloc[0]
        row1 = self.raw_df.iloc[1]
        row2 = self.raw_df.iloc[2]
        total_cols = len(row0)

        # Собираем все позиции непустых ячеек в Row 0 — это начала блоков
        block_starts = []
        for col_idx in range(total_cols):
            val = row0.iloc[col_idx]
            if pd.notna(val):
                block_starts.append((col_idx, str(val)))

        # Отбираем только магазины (пропускаем «Изделие» и «Итого»)
        shops = []
        for i, (start_col, name) in enumerate(block_starts):
            if name in ("Изделие", "Итого"):
                continue

            end_col = block_starts[i + 1][0] if i + 1 < len(block_starts) else total_cols

            shop_code = str(row1.iloc[start_col]) if pd.notna(row1.iloc[start_col]) else ""

            # Находим метрики внутри блока по Row 2
            metrics = {}
            for col in range(start_col, end_col):
                metric_name = row2.iloc[col]
                if pd.notna(metric_name):
                    metrics[str(metric_name)] = col

            shops.append({
                "name": name,
                "code": shop_code,
                "metrics": metrics,
            })

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

                # Извлекаем метрики по фактическим позициям колонок
                for metric_name, col_idx in shop["metrics"].items():
                    value = self.raw_df.iloc[raw_row_idx, col_idx]
                    output_name = self.METRIC_MAPPING.get(metric_name, metric_name)
                    row_data[output_name] = value if pd.notna(value) else 0

                all_rows.append(row_data)

        self.df = pd.DataFrame(all_rows)

        # Фильтруем строки-итоги (где Артикул не является числом)
        self.df = self.df[
            pd.to_numeric(self.df["Артикул"], errors="coerce").notna()
        ].reset_index(drop=True)

        # Преобразование Артикул и Код магазина в числовые
        self.df["Артикул"] = pd.to_numeric(self.df["Артикул"], errors="coerce").astype(int)
        self.df["Код магазина"] = pd.to_numeric(self.df["Код магазина"], errors="coerce").fillna(0).astype(int)

        # Преобразование типов
        for col in self.df.columns:
            if col in self.BASE_COLNAMES:
                continue
            if col in self.INTEGER_FIELDS:
                self.df[col] = pd.to_numeric(self.df[col], errors="coerce").fillna(0).astype(int)
            else:
                self.df[col] = pd.to_numeric(self.df[col], errors="coerce").fillna(0)

        # Упорядочиваем колонки: базовые + метрики в порядке METRIC_ORDER
        found_metrics = [m for m in self.METRIC_ORDER if m in self.df.columns]
        self.df = self.df.reindex(columns=self.BASE_COLNAMES + found_metrics)

        # Сортировка по Артикул, затем по Магазин
        self.df = self.df.sort_values(
            by=["Артикул", "Магазин"]
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