import pandas as pd


def get_col_name_for_value(
    df: pd.DataFrame, required_value: str = "Товар"
) -> list[str]:
    """
    Позволяет получить название полей в таблице, в которых записано
    искомое значение.

    Parameters
    ----------
    df : pd.DataFrame
        Таблица с данными для поиска.
    required_value : str, optional
        Значение, для которого требуется найти имя поля.

    Returns
    -------
    list[str]
        Список наименований полей, где встречается искомое значение.
    """
    return df.columns[df.eq(required_value).any()].tolist()
