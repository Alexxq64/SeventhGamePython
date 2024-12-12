from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import os
import time

def extract_matches_with_selenium_to_excel(tournament_url):
    """
    Использует Selenium для извлечения матчей с сайта Livesport и сохраняет данные в Excel файл.

    Параметры:
        tournament_url (str): URL страницы турнира.
    """
    # Настраиваем Selenium
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Открытие в фоновом режиме
    chrome_service = Service("C:\\Users\\User\\Desktop\\Настройки\\chromedriver.exe")

    # Инициализация браузера
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    # Путь к Excel файлу
    file_path = "D:\\Hobby\\SeventhGamePython\\7thGame.xlsx"

    try:
        # Открываем страницу
        driver.get(tournament_url)

        # Небольшая пауза для загрузки данных
        time.sleep(5)

        # Находим элементы матчей
        matches = driver.find_elements(By.CSS_SELECTOR, "div.event__match")

        # Загружаем файл Excel или создаём новый
        wb = load_workbook(file_path) if os.path.exists(file_path) else Workbook()
        ws = wb.create_sheet("MatchLinks") if "MatchLinks" not in wb.sheetnames else wb["MatchLinks"]

        # Заголовки
        ws.cell(1, 1, "Ссылка на матч")
        ws.cell(1, 2, "Игрок 1")
        ws.cell(1, 3, "Игрок 2")

        row = 2

        # Обрабатываем каждый элемент матча
        for match in matches:
            try:
                # Находим ссылку на матч
                link_element = match.find_element(By.CSS_SELECTOR, "a.eventRowLink")
                match_url = link_element.get_attribute("href")

                # Извлекаем игроков
                home_player = match.find_element(By.CSS_SELECTOR, "div.event__participant--home").text
                away_player = match.find_element(By.CSS_SELECTOR, "div.event__participant--away").text

                # Записываем данные в Excel
                ws.cell(row, 1, match_url)
                ws.cell(row, 1).hyperlink = match_url
                ws.cell(row, 2, home_player)
                ws.cell(row, 3, away_player)

                row += 1
            except Exception as e:
                print(f"Ошибка обработки матча: {e}")

        # Автоподгонка ширины столбцов
        for col in range(1, 4):
            column = get_column_letter(col)
            max_length = max(len(str(ws[f'{column}{r}'].value)) for r in range(1, ws.max_row + 1))
            ws.column_dimensions[column].width = max_length + 2

        # Сохраняем файл Excel
        wb.save(file_path)

        # Выводим сообщение об успехе
        print(f"Ссылки на матчи и данные игроков сохранены в Excel файл по пути: {file_path}")

    finally:
        # Закрываем браузер
        driver.quit()

# Основная функция
def main():
    tournament_url = "https://www.livesport.com/tennis/atp-singles/adelaide/results/"
    extract_matches_with_selenium_to_excel(tournament_url)

if __name__ == "__main__":
    main()
