import time
import pandas as pd
from playwright.sync_api import sync_playwright


def run() -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.moex.com/")
    page.get_by_label("Открыть бургер-меню").click()  # ???
    page.get_by_role("banner").get_by_role("link", name="Срочный рынок").click()
    page.get_by_role("button", name="Согласен", exact=True).click()
    page.get_by_role("link", name="Индикативные курсы").click()
    page.get_by_label("Показать").click()
    with page.expect_download() as download_info:
        page.get_by_role("link", name="Получить данные в XML").click()
    download = download_info.value
    download.save_as("./data1.xml")
    time.sleep(10)

    page.locator("#app div").filter(has_text="Выберите валюты: USD/RUB").locator("use").click()
    page.get_by_role("link", name="JPY/RUB").click()
    page.get_by_label("Показать").click()
    with page.expect_download() as download1_info:
        page.get_by_role("link", name="Получить данные в XML").click()
    download1 = download1_info.value
    download1.save_as("./data2.xml")

    # Читаем скачанные файлы
    df1 = pd.read_xml("./data1.xml")
    df2 = pd.read_xml("./data2.xml")

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run()
