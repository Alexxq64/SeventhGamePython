from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


def click_detail_button(driver):
    """
    Находит и нажимает на кнопку с селектором на странице матча.

    Аргументы:
        driver: Объект WebDriver для управления браузером.
    """
    try:
        # Ожидаем, пока кнопка станет кликабельной
        detail_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "#detail > div.filterOver.filterOver--indent > div > a:nth-child(3) > button"))
        )
        # Кликаем по кнопке
        detail_button.click()
        print("Кнопка нажата успешно.")
    except Exception as e:
        print(f"Ошибка при нажатии на кнопку: {e}")


def open_match_page(match_url):
    """
    Переходит по ссылке матча и нажимает на кнопку.

    Аргументы:
        match_url (str): Ссылка на страницу матча.
    """
    # Настройка WebDriver
    chrome_options = Options()
    # Открываем браузер в видимом режиме
    # (удаляем строку headless для видимости)
    chrome_service = Service("C:\\Users\\User\\Desktop\\Настройки\\chromedriver.exe")
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    try:
        # Переходим на страницу матча
        driver.get(match_url)
        time.sleep(3)  # Ждём загрузки страницы

        # Нажимаем на нужную кнопку
        click_detail_button(driver)
        time.sleep(3)  # Ожидаем действия после нажатия

        print("Оставляем браузер открытым. Закройте его вручную, когда закончите.")
        input("Нажмите Enter, чтобы завершить программу...")

    except Exception as e:
        print(f"Ошибка: {e}")
    finally:
        # Браузер больше не закрывается автоматически
        pass


# Основная функция для тестирования
if __name__ == "__main__":
    match_url = "https://www.livesport.com/game/A5bhEEyG/#/game-summary"
    open_match_page(match_url)
