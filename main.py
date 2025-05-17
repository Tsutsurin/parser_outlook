import time
from utils import OutlookHandler, ExcelHandler, parse_email_data, get_parameters_from_file


def main():
    params = get_parameters_from_file()

    if not params:
        print('Работа программы завершена из-за ошибок в файле конфигурации.')
        return

    folder_name = params.get('folder_name')
    excel_path = params.get('excel_path')
    mailbox_name = params.get('mailbox_name')

    if not all([folder_name, excel_path, mailbox_name]):
        print('❌ Не все параметры указаны в config.txt.')
        return

    outlook_handler = OutlookHandler()
    try:
        mailbox = outlook_handler.get_mailbox(mailbox_name)
        target_folder = outlook_handler.find_folder(mailbox, folder_name)
        if not target_folder:
            print(f'❌ Папка \'{folder_name}\' не найдена в почтовом ящике \'{mailbox_name}\'!')
            return
    except ValueError as e:
        print(e)
        return

    excel_handler = ExcelHandler(excel_path)
    processed_message_ids = set()

    while True:
        try:
            messages = target_folder.Items
            messages.Sort('[ReceivedTime]', True)

            new_rows = []

            for message in messages:
                if message.EntryID not in processed_message_ids:
                    word1, inc_number, quoted_text, date_reg1, date_reg2, date_received = parse_email_data(message)

                    new_row = [excel_handler.worksheet.max_row, word1, inc_number, quoted_text, date_reg1, date_reg2, date_received]
                    new_rows.append(new_row)
                    processed_message_ids.add(message.EntryID)

            for row in reversed(new_rows):
                excel_handler.add_data_to_excel(row)

            excel_handler.save_workbook()
            print(f'✅ Готово! Сохранено в: {excel_path}')

        except Exception as e:
            print(f'❌ Ошибка во время работы: {e}')

        print('⏳ Ожидание следующего запуска...')
        time.sleep(3600)


if __name__ == '__main__':
    main()
