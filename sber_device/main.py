import os
import time

import streamlit as st
from config.configurations import Config, ReportConfig
from processing.dns_extended import dns_retailer_config_extended, run_dns_extended
from processing.mvm import mvm_retailer_config, run_mvm
from utils.io import DEFAULT_SAVE_PATH, Writer
from utils.mappings import replacer

st.set_page_config(page_title="СберДевайс", page_icon="sber_logo.png", layout="wide")

st.title("Сбер")
st.divider()

st.header("Порядок работы")
st.markdown("""
1. Загрузить файл, используя кнопку на панели слева или перетащив файл на поле загрузки.
2. Выбрать сеть из выпадающего списка.
3. Выбрать дату отчета.
4. Нажать кнопку "Обработать загруженный файл". На странице отобразится таблица с обработанными данными.
5. Если все верно, можно загрузить файл в формате Excel, нажав на кнопку "Скачать файл".
""")

# Товарные категории
categories = ["Медиаплееры", "СберДевайс", "Аксессуары", "Телевизоры", "Лампы"]

sidebar = st.sidebar
sidebar.header("Загрузка файла и ввод параметров")

uploaded = sidebar.file_uploader(
    "Загрузить файл", type=["xlsx", "xls"], accept_multiple_files=False
)


if uploaded is not None:
    # Выбор названия сети
    retailer = sidebar.selectbox(
        "Выбрать сеть",
        options=["ДНС", "МВМ"],
        help="Выбор сети",
        index=None,
        placeholder="Выбрать сеть",
    )

    # Выбор периода отчета
    date_input = sidebar.date_input("Выбрать период", help="Выбор даты отчета")
    period = replacer(date_input.strftime("%d %b %Y")) if date_input else ""

    # Выбор товарной категории отчета
    selected_category = sidebar.selectbox(
        "Выбрать категорию", options=categories, help="Выбор категории"
    )

    # Кнопка запуска обработки загруженного файла
    process_button = sidebar.button(
        "Обработать загруженный файл", icon=":material/manufacturing:"
    )

    @st.fragment
    def process():
        report_config = ReportConfig(report_period=period, category=selected_category)

        match retailer:
            case "МВМ":
                result = run_mvm(uploaded)
                config = Config(mvm_retailer_config, report_config)
            case "ДНС":
                result = run_dns_extended(uploaded)
                config = Config(dns_retailer_config_extended, report_config)
            case _:
                raise ValueError("Неизвестное название сети.")
        return result, config

    st.divider()

    if process_button:
        fname = "current_result.xlsx"

        if fname in os.listdir(DEFAULT_SAVE_PATH):
            os.remove(DEFAULT_SAVE_PATH.joinpath(fname))

        with st.spinner("Работаю...", show_time=True):
            result, config = process()
            writer = Writer(config)
            writer.export_to_xls(result, fname)
            badge = st.badge("Готово!", color="green", icon=":material/check:")

        if result is not None:
            st.subheader("Обработанные данные", divider=True)
            st.dataframe(result, use_container_width=True, hide_index=True)
            time.sleep(2)
            badge.empty()
        else:
            st.write("Ошибка обработки. В таблице нет данных.")

        if result is not None:
            output_fname = f"{selected_category}_{retailer}_{period}.xlsx"

            with open(DEFAULT_SAVE_PATH.joinpath(fname), "rb") as f:
                file_object = f.read()
                download_button = st.download_button(
                    "Скачать файл",
                    data=file_object,
                    file_name=output_fname,
                    on_click=lambda: os.remove(DEFAULT_SAVE_PATH.joinpath(fname)),
                    help="Скачать результат в формате .xlsx",
                    icon=":material/download:",
                )
