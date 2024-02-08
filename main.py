from twilio.rest import Client
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

def main():
    # ------------------------------------------------------------------------------------------------------------------
    def minmax_list_parsing(filename):
        # Открываем файл xlsx
        wb = openpyxl.load_workbook(filename)
        # Выбираем активный лист
        sheet = wb.active

        # Список для хранения всех данных
        all_data = []
        # Список для хранения данных без подмассивов с нулевыми значениями второго и третьего индексов
        data_without_zeros = []

        # Начинаем считывание с ячейки C2 и двигаемся вниз до первой пустой ячейки
        row = 2
        while True:
            # Получаем значения из столбцов C, I и J
            nomenclature = sheet[f'C{row}'].value
            value_i = sheet[f'I{row}'].value
            value_j = sheet[f'J{row}'].value

            # Если значение номенклатуры пустое, завершаем чтение
            if nomenclature is None:
                break

            # Добавляем значения в список всех данных в виде вложенного списка
            all_data.append([nomenclature, str(value_i), str(value_j)])

            # Проверяем, если массив в массиве имеет значения 0 во втором и третьем индексах, не добавляем его
            if str(value_i) != '0' or str(value_j) != '0':
                data_without_zeros.append([nomenclature, str(value_i), str(value_j)])

            row += 1

        # Закрываем книгу после чтения
        wb.close()

        return all_data, data_without_zeros

    # Пример использования
    filename = 'minmax-list.xlsx'
    all_data, data_without_zeros = minmax_list_parsing(filename)
    # print("Весь массив:")
    # print(all_data)
    # print("\nМассив без нулевых значений второго и третьего индексов:")
    # print(data_without_zeros)

    # ------------------------------------------------------------------------------------------------------------------
    def ved_parsing(filename_ved):
        # Открываем файл xlsx
        wb = openpyxl.load_workbook(filename_ved)
        # Выбираем активный лист
        sheet = wb.active

        # Список для хранения данных
        data_ved = []

        # Начинаем считывание с ячейки D11 и двигаемся вниз до первой пустой ячейки в столбце D
        row = 11
        while True:
            # Получаем значения из столбцов D и P
            nomenclature_cell = sheet[f'D{row}']
            value_p_cell = sheet[f'P{row}']

            # Проверяем, что значения ячеек не являются None
            if nomenclature_cell.value is not None and value_p_cell.value is not None:
                # Убираем пробелы в начале и конце строки
                nomenclature = nomenclature_cell.value.strip()
                value_p = str(value_p_cell.value).strip()

                # Удаляем запятые и точки в конце строки
                nomenclature = nomenclature.rstrip(',. ')

                # Добавляем значения в список данных в виде вложенного списка
                data_ved.append([nomenclature, value_p])

            # Если значение номенклатуры пустое, завершаем чтение
            if nomenclature_cell.value is None:
                break

            row += 1

        # Закрываем книгу после чтения
        wb.close()

        return data_ved

    # Пример использования
    filename_ved = 'ved.xlsx'
    data_ved = ved_parsing(filename_ved)
    # print(data_ved)

    # ------------------------------------------------------------------------------------------------------------------
    # Two arrays
    array1 = data_without_zeros
    array2 = data_ved

    # Extracting the first elements from each array
    names_array1 = [item[0] for item in array1]
    names_array2 = [item[0] for item in array2]

    output_messages = ""

    # Counter for numbering items
    count = 0

    for i, name in enumerate(names_array1):
        if name in names_array2:
            value_array1 = int(array1[i][2])
            value_array2 = int(array2[names_array2.index(name)][1])
            value_array3 = int(array1[i][1])

            if value_array1 > value_array2:
                count += 1
                output_messages += f"\n{count}. Нужно заказать: {name}, текущее кол-во на складе: {value_array2}\n"
                output_messages += f"        по программе мин-макс должно быть: минимум {value_array1}, заказать до {value_array3}\n"
        else:
            value_array1 = int(array1[i][2])
            value_array3 = int(array1[i][1])

            count += 1
            output_messages += f"\n{count}. На складе отсутствует: {name}\n"
            output_messages += f"        по программе мин-макс должно быть: минимум {value_array1}, заказать до {value_array3}\n"

    print(output_messages)

    # ------------------------------------------------------------------------------------------------------------------
    # Email configuration
    sender_email = "order.notifier@mail.ru"
    receiver_emails = ["erm.zhumagulov@gmail.com", "ermek.zhumagulov@pepsico.com"]    # "vitaliy.bobylev@pepsico.com"
    password = "82w8yvf24ikVpD657pej"  # Use your Mail.ru email password

    # Create message container
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = ', '.join(receiver_emails)  # Join the list of recipients with commas
    message['Subject'] = "Статусы по запасам"

    # Email body
    body = output_messages
    message.attach(MIMEText(body, 'plain'))

    try:
        # Connect to the Mail.ru SMTP server
        smtp_server = smtplib.SMTP_SSL('smtp.mail.ru', 465)
        smtp_server.login(sender_email, password)

        # Send the email to all recipients
        smtp_server.sendmail(sender_email, receiver_emails, message.as_string())

        # Close the connection to the SMTP server
        smtp_server.quit()

        print("All emails sent successfully!")

    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()