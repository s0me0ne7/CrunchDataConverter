from collections import namedtuple
from dataclasses import dataclass, InitVar, asdict


@dataclass
class RetailerConfig:
    """
    Параметры отчета для выбранной компании.

    Parameters
    ----------
    company_name : str
        Наименование компании
    table_display_name : str
        Название таблицы, как оно отображается на странице Excel
    excel_table_style_name : str
        Название встроенного стиля таблицы в Excel
    start_summation_row : str
        Имя столбца, с которого начинаются ячейки, где суммируются результаты.
    table_header_style : str
        Название стиля ячейки в Excel. По умолчанию "Headline 1", что соответствует стилю
        заголовка первого уровня.
    """

    company_name: str
    table_display_name: str
    excel_table_style_name: str
    start_summation_row: str
    table_header_style: str = "Headline 1"


@dataclass
class ReportConfig:
    """
    Параметры формирования отчета

    Parameters
    ----------
    report_period : str
        Период отчета в формате строки для вывода в заголовке страницы.
    category : str
        Название категории. Оно же используется для вкладки в книге Excel
    """

    report_period: str
    category: str


@dataclass(slots=True)
class Config:
    retailer_config: InitVar[RetailerConfig]
    report_config: InitVar[ReportConfig]
    parameters: namedtuple = None

    def __post_init__(
        self, retailer_config: RetailerConfig, report_config: ReportConfig
    ):
        """
        Формирует параметры конфигурации в виде именованного кортежа.
        """
        ReturnTuple = namedtuple(
            "ReturnTuple",
            list((asdict(retailer_config) | asdict(report_config)).keys()),
        )
        self.parameters = ReturnTuple(
            **(asdict(retailer_config) | asdict(report_config))
        )
