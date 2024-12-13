from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time


def extract_data_from_page(driver):
    """
    Извлекает информацию о счёте всех геймов с текущей страницы (point-by-point),
    а также информацию о подающем игроке в первом гейме первого сета.

    Аргументы:
        driver: WebDriver.

    Возвращает:
        data (list): Список всех данных о счёте геймов (без лишних символов).
        server_info (str): Информация о подающем игроке в первом гейме.
    """
    game_data = []
    server_info = "Неизвестно"  # Если не удастся определить подающего игрока

    try:
        # Перебор с шагом 2 (nth-child от 2 до 26)
        for i in range(2, 27, 2):
            try:
                # Формируем селектор для данных о счёте
                score_selector = f"#detail > div.matchHistoryRowWrapper > div:nth-child({i}) > div.matchHistoryRow__scoreBox"
                score_element = driver.find_element(By.CSS_SELECTOR, score_selector)
                game_data.append(score_element.text.replace("\n", ""))

                # Если мы находимся в первом гейме первого сета, извлекаем информацию о подающем
                if i == 2:  # Первый гейм первого сета
                    # Селекторы для подающих
                    serve_selector_left = f"#detail > div.matchHistoryRowWrapper > div:nth-child({i}) > div.matchHistoryRow__servis.matchHistoryRow__home > div > svg"
                    serve_selector_right = f"#detail > div.matchHistoryRowWrapper > div:nth-child({i}) > div.matchHistoryRow__servis.matchHistoryRow__away > div > svg"

                    try:
                        # Проверяем, если есть SVG в левой части (подает Игрок 1)
                        if driver.find_element(By.CSS_SELECTOR, serve_selector_left):
                            server_info = "Игрок 1 подает"
                    except:
                        try:
                            # Проверяем, если есть SVG в правой части (подает Игрок 2)
                            if driver.find_element(By.CSS_SELECTOR, serve_selector_right):
                                server_info = "Игрок 2 подает"
                        except:
                            pass

            except Exception:
                break

        return game_data, server_info
    except Exception as e:
        print(f"Ошибка при извлечении данных: {e}")
        return [], "Неизвестно"


def switch_tabs_and_collect_data(driver):
    """
    Последовательно переключается между вкладками и собирает информацию с каждой.

    Аргументы:
        driver: WebDriver.

    Возвращает:
        all_data (dict): Словарь с данными по каждой вкладке.
    """
    all_data = {}
    try:
        tab_buttons = driver.find_elements(
            By.CSS_SELECTOR,
            "#detail > div.subFilterOver.subFilterOver--indent > div > a > button"
        )

        print(f"Найдено вкладок: {len(tab_buttons)}")

        for i, button in enumerate(tab_buttons):
            try:
                print(f"Переход на вкладку {i}...")
                button.click()
                time.sleep(2)

                # Извлекаем данные с текущей страницы
                data, server_info = extract_data_from_page(driver)
                all_data[f"point-by-point/{i}"] = data
                if i == 0:  # Только в первом наборе данных сохраняем информацию о подающем
                    all_data["server_info"] = server_info

            except Exception as e:
                print(f"Ошибка при переходе на вкладку {i}: {e}")
    except Exception as e:
        print(f"Ошибка при работе с вкладками: {e}")

    return all_data


def process_match_page(match_url):
    """
    Переходит на страницу матча, собирает данные со всех вкладок point-by-point.

    Аргументы:
        match_url (str): URL страницы матча.
    """
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_service = Service("C:\\Users\\User\\Desktop\\Настройки\\chromedriver.exe")

    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    try:
        driver.get(match_url)
        time.sleep(3)

        try:
            button_selector = "#detail > div.filterOver.filterOver--indent > div > a:nth-child(3) > button"
            button = driver.find_element(By.CSS_SELECTOR, button_selector)
            button.click()
            time.sleep(2)
        except Exception as e:
            print(f"Не удалось нажать кнопку 'point-by-point': {e}")
            return

        all_data = switch_tabs_and_collect_data(driver)

        # Выводим информацию о подающем игроке первой
        if "server_info" in all_data:
            print(f"Информация о подающем игроке: {all_data['server_info']}")

        # Выводим данные по остальным вкладкам
        print("Собранные данные со всех вкладок:")
        for tab, data in all_data.items():
            if tab != "server_info":  # Пропускаем вывод информации о подающем игроке, так как она уже выведена
                print(f"{tab}: {data}")

    finally:
        driver.quit()


# Основная функция https://www.livesport.com/game/YJwpuC0Q/#/game-summary
if __name__ == "__main__":
    # match_url = "https://www.livesport.com/game/A5bhEEyG/#/game-summary"
    # process_match_page(match_url)
    match_url = "https://www.livesport.com/game/YJwpuC0Q/#/game-summary"
    process_match_page(match_url)
