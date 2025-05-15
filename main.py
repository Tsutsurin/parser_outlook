import win32com.client
from openpyxl import Workbook
import os
import re
from datetime import datetime, timedelta

def export_emails_from_folder(folder_name, excel_output_path, mailbox_name):
    # Подключаемся к Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Получаем доступ к почтовому ящику
    try:
        mailbox = outlook.Folders[mailbox_name]
    except Exception as e:
        print(f"❌ Почтовый ящик '{mailbox_name}' не найден: {e}")
        return

    # Ищем папку
    def find_folder(parent_folder, target_name):
        if parent_folder.Name == target_name:
            return parent_folder
        for folder in parent_folder.Folders:
            found = find_folder(folder, target_name)
            if found:
                return found
        return None

    target_folder = find_folder(mailbox, folder_name)

    if not target_folder:
        print(f"❌ Папка '{folder_name}' не найдена в почтовом ящике '{mailbox_name}'!")
        return

    # Создаем Excel-файл
    wb = Workbook()
    ws = wb.active
    ws.title = "Темы писем"

    # Заголовки столбцов Excel
    ws.append(["№", "Слово 1", "INC-номер", "Текст в кавычках", "Дата регистрации (DD.MM.YYYY)", "Дата регистрации (текст)", "Дата получения письма"])

    # Собираем все письма из папки
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True)  # Сортировка по дате (новые сверху)

    for idx, message in enumerate(messages, start=1):
        subject = message.Subject
        body = message.Body  # Получаем тело письма

        # Разбираем тему письма
        match_subject = re.match(r"(\w+)\s+([A-Z]+-\d+)\s+\"(.*)\"", subject)
        if match_subject:
            word1, inc_number, quoted_text = match_subject.groups()
        else:
            word1, inc_number, quoted_text = "N/A", "N/A", "N/A"

        # Ищем дату регистрации в формате DD.MM.YYYY HH:MM:SS
        match_date1 = re.search(r"Дата регистрации\s*(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})", body)
        date_reg1 = match_date1.group(1) if match_date1 else "N/A"

        # Ищем дату регистрации в текстовом формате (15 May 2025, 12:24:05 UTC) и преобразуем
        match_date2 = re.search(r"(\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4},\s+\d{2}:\d{2}:\d{2})\s+UTC", body)
        date_reg2 = "N/A"
        if match_date2:
            try:
                # Преобразуем текстовую дату в объект datetime
                date_obj = datetime.strptime(match_date2.group(1), "%d %b %Y, %H:%M:%S")
                # Добавляем 3 часа
                date_obj = date_obj + timedelta(hours=3)
                # Форматируем в нужный формат
                date_reg2 = date_obj.strftime("%d.%m.%Y %H:%M:%S")
            except ValueError:
                date_reg2 = "Некорректная дата"


        date_received = message.ReceivedTime.strftime("%d.%m.%Y %H:%M:%S") if hasattr(message, 'ReceivedTime') else "N/A"

        # Добавляем 3 часа к датам регистрации, если они найдены
        if date_reg1 != "N/A":
            try:
                date_obj = datetime.strptime(date_reg1, "%d.%m.%Y %H:%M:%S")
                date_obj = date_obj + timedelta(hours=3)
                date_reg1 = date_obj.strftime("%d.%m.%Y %H:%M:%S")
            except ValueError:
                date_reg1 = "Некорректная дата"


        ws.append([idx, word1, inc_number, quoted_text, date_reg1, date_reg2, date_received])

    # Сохраняем Excel
    wb.save(excel_output_path)
    print(f"✅ Готово! Сохранено в: {excel_output_path}")


def get_parameters_from_file(config_file="config.txt"):
    """
    Читает параметры из текстового файла.
    Формат файла:
    folder_name=ИмяПапки
    excel_path=ИмяФайла.xlsx
    mailbox_name=адрес@почты.ru
    """
    params = {}
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if "=" in line:
                    key, value = line.split("=", 1)
                    params[key.strip()] = value.strip()
    except FileNotFoundError:
        print(f"❌ Файл конфигурации '{config_file}' не найден.  Используйте config.txt для настройки.")
        return None
    except Exception as e:
        print(f"❌ Ошибка при чтении файла конфигурации: {e}")
        return None
    return params


if __name__ == "__main__":
    # Получаем параметры из файла конфигурации
    params = get_parameters_from_file()

    if params:
        folder_name = params.get("folder_name")
        excel_path = params.get("excel_path")
        mailbox_name = params.get("mailbox_name")

        if not all([folder_name, excel_path, mailbox_name]):
            print("❌ Не все параметры указаны в config.txt.")
        else:
            export_emails_from_folder(folder_name, excel_path, mailbox_name)
