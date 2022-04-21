import os
from os.path import basename
from pathlib import Path

import pandas as pd
import xlwings as xw

# global options and vars
os.environ['NO_PROXY'] = 'moex.com'
FIREFOX_BINARY = 'path_to_firefox_bin'
EMAIL_PASS = 'my_pass'


def send_mail(send_from: str, send_to: list, subject: str, text: str, files: list = None):
    """
    Sends the mail to its destination; used email module with Yandex mail server

    :return void
    """

    from email import encoders
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import formatdate
    from smtplib import SMTP_SSL

    assert isinstance(send_to, list)

    mail_server = 'smtp.yandex.ru'
    mail_server_port = 465

    s = SMTP_SSL(mail_server, mail_server_port)
    s.ehlo(send_from)
    s.login(send_from, EMAIL_PASS)

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for attachment in files:
        if not os.path.isfile(attachment):
            assert f'File does not exist: {f}. Skipping...'

        with open(attachment, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())

        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={basename(attachment)}',
        )
        msg.attach(part)

    s.sendmail(send_from, send_to, msg.as_string())
    s.close()


def get_string_declension(row_number: int):
    """
    Returns the correct declension of the noun

    :return out: string
    """
    words = ['строк', 'строки', 'строка']

    remainder = row_number % 100
    if 11 <= remainder <= 19:
        out = words[0]
    else:
        i = remainder % 10
        if i == 1:
            out = words[2]
        elif i in [2, 3, 4]:
            out = words[1]
        else:
            out = words[0]
    return out


def get_moex_data():
    """
    Working with the Selenium to obtain ruble exchange rate data from moex.com

    :return usd_data_values: list, eur_data_values: list
    """

    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.firefox.service import Service
    from webdriver_manager.firefox import GeckoDriverManager
    from selenium.webdriver.firefox.options import Options
    from selenium.webdriver.support.ui import Select
    from selenium.common.exceptions import NoSuchElementException

    options = Options()
    options.binary_location = FIREFOX_BINARY
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    driver.maximize_window()
    driver.implicitly_wait(0.5)
    driver.get("http://www.moex.com")

    elem = driver.find_element(by=By.XPATH, value="//button[@class='header-menu__link is-button js-menu-dropdown-button']")
    elem.click()
    elem = driver.find_element(by=By.LINK_TEXT, value="Срочный рынок")
    elem.click()
    try:
        elem = driver.find_element(by=By.LINK_TEXT, value="Согласен")
        elem.click()
    except NoSuchElementException:
        pass
    elem = driver.find_element(by=By.LINK_TEXT, value="Индикативные курсы")
    ActionChains(driver).move_to_element(elem).click().perform()
    elem.click()

    # work with table for USD/RUB
    table_id = driver.find_element(by=By.CLASS_NAME, value='tablels')
    rows = table_id.find_elements(by=By.TAG_NAME, value="tr")

    usd_data_values = []

    for row in rows[2::]:
        usd_data_values.append(row.text.replace(',', '.').split(' '))

    # work with table for EUR/RUB
    elem = Select(driver.find_element(by=By.ID, value="ctl00_PageContent_CurrencySelect"))
    elem.select_by_value("EUR_RUB")

    table_id = driver.find_element(by=By.TAG_NAME, value="table")
    rows = table_id.find_elements(by=By.TAG_NAME, value="tr")

    eur_data_values = []

    if not rows[0].text:
        from_dates_id = Select(driver.find_element(by=By.ID, value="d1year"))
        from_dates_id.select_by_value("2021")
        till_dates_id = Select(driver.find_element(by=By.ID, value="d2year"))
        till_dates_id.select_by_value("2021")
        show_bt = driver.find_element(by=By.CLASS_NAME, value='button80')
        ActionChains(driver).move_to_element(show_bt).click().perform()

        table_id = driver.find_element(by=By.CLASS_NAME, value='tablels')
        rows = table_id.find_elements(by=By.TAG_NAME, value="tr")

        for row in rows[2::]:
            eur_data_values.append(row.text.replace(',', '.').split(' '))

        print("No data for EUR/RUB! Used 2021 year")
    else:
        for row in rows[2::]:
            eur_data_values.append(row.text.split(' '))

    driver.close()

    return usd_data_values, eur_data_values


def work_with_excel(usd_data_values: list, eur_data_values: list):
    """
    Saves the ruble exchange rate data to a Excel file,
    formatting by width and setting the financial rate for numeric values

    :return last_row: int
    """

    # MOEX data
    column_names = ['Дата', 'Значение курса промежуточного клиринга', 'Время промежуточного клиринга',
                    'Значение курса основного клиринга', 'Время основного клиринга']

    # Format data
    usd_data = [row for row in usd_data_values if '-' not in row]
    eur_data = [row for row in eur_data_values if '-' not in row]

    usd_data_len = len(usd_data)
    eur_data_len = len(eur_data)
    min_len = min(usd_data_len, eur_data_len)
    max_size_array = usd_data if usd_data_len != min_len else eur_data
    while usd_data_len != eur_data_len:
        del max_size_array[-1]

    df_usd_rub = pd.DataFrame(usd_data, columns=column_names)
    df_eur_rub = pd.DataFrame(eur_data, columns=column_names)

    for column in ['Значение курса промежуточного клиринга', 'Значение курса основного клиринга']:
        df_usd_rub[column] = pd.to_numeric(df_usd_rub[column])
        df_eur_rub[column] = pd.to_numeric(df_eur_rub[column])

    change_eur_usd = pd.DataFrame({'Изменение': df_eur_rub['Значение курса основного клиринга'] / df_usd_rub[
        'Значение курса основного клиринга']})

    # Work with Excel
    wb = xw.Book()
    sheet = wb.sheets[0]
    sheet.range("A1").options(index=False).value = df_usd_rub
    sheet.range("F1").options(index=False).value = df_eur_rub
    sheet.range("K1").options(index=False).value = change_eur_usd

    sheet.autofit()
    last_row = sheet.range(1, 1).end('down').row

    for index in ["B", "D", "G", "I", "K"]:
        # set currency format
        sheet.range(f"{index}:{index}").number_format = "#,##0.0000 [$₽-ru-RU]"

    # sum check
    sheet.range(f"B{last_row + 1}").value = f"=SUM(B2:B{last_row})"
    sum_check = sum(df_usd_rub["Значение курса промежуточного клиринга"])
    if sum_check == sheet.range(f"B{last_row + 1}").value:
        print("SUM check successfully!")
    sheet.range(f"B{last_row + 1}").api.Delete()

    wb.save('moex.xlsx')
    wb.close()

    return last_row


if __name__ == "__main__":
    usd_data_values, eur_data_values = get_moex_data()
    last_row = work_with_excel(usd_data_values, eur_data_values)
    # send Excel file to email
    string_out = get_string_declension(last_row)
    text = f'Moex data in Excel file: В документе содержится: {last_row} {string_out}'
    send_mail(send_from='mail@data.ru', send_to=['mail@data.ru'],
              subject='Moex data', text=text, files=[Path('moex.xlsx')])
    print(f"Email with Excel file successfully sent!")
