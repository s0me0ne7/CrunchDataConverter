import re

# Шаблон поиска названий месяцев
month_pattern = re.compile(r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec", re.I)

# Словарь сопоставления названий месяцев по-английски и по-русски
month_mapper = dict(
    Jan="янв",
    Feb="фев",
    Mar="мар",
    Apr="апр",
    May="май",
    Jun="июн",
    Jul="июл",
    Aug="авг",
    Sep="сен",
    Oct="окт",
    Nov="ноя",
    Dec="дек",
)


def replacer(s: str) -> str:
    """
    Функция, которая преобразует месяцы с английского языка на русский.

    Parameters
    ----------
    s : str
        Строка с датой, в которой есть название месяца.

    Returns
    -------
    str
        Строка с преобразованным названием месяца.
    """
    match = month_pattern.search(s)

    if match is not None:
        month = match.group(0)
        s = s.replace(month, month_mapper.get(month, month))
    return s
