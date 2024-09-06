from flask import Flask, render_template, request, send_file, jsonify
import psycopg2
from openpyxl import load_workbook
from datetime import datetime
import pythoncom
import win32com.client as win32  # Добавлено для конвертации Excel в PDF
#from io import BytesIO
#from fpdf import FPDF
#import os
app = Flask(__name__)

def db_connection():
    with psycopg2.connect(
        host='localhost',
        database='form_tth',
        user='postgres',
        password='123456'
    ) as conn:
        return conn

def get_laboratories():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM laboratory")
    laboratories = cursor.fetchall()
    cursor.close()
    conn.close()
    return [{'id': row[0], 'name': row[1]} for row in laboratories]

def get_delivery_data():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT recipient, inn, razgruzka FROM delivery")
    recipients = cursor.fetchall()
    cursor.close()
    conn.close()
    return [{'recipient': row[0], 'inn': row[1], 'razgruzka': row[2]} for row in recipients]

# Функция для получения данных о прицепах
def get_trailer_data():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, trailer_number, section1, section2, section3, section4, section5, section6, section7 FROM trailers")
    trailers = cursor.fetchall()
    cursor.close()
    conn.close()
    return [{'id': row[0], 'number': row[1], 'sections': row[2:]} for row in trailers]

@app.route('/get_initials/<int:driver_id>', methods=['GET'])
def get_initials(driver_id):
    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT initials FROM drivers WHERE id = %s", (driver_id,))
    initials = cursor.fetchone()

    cursor.close()
    conn.close()

    if initials:
        return jsonify({'initials': initials[0]})
    else:
        return jsonify({'initials': ''})


@app.route('/', methods=['GET', 'POST'])
def index():
    drivers = []
    transports = []
    senders = []
    laboratories = get_laboratories()
    delivery_data = get_delivery_data()


    # Получаем список водителей, транспорта и отправителей из базы данных
    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, full_name FROM drivers")
    drivers = cursor.fetchall()

    cursor.execute("SELECT transport_number FROM transport")
    transports = cursor.fetchall()

    cursor.execute("SELECT id, name FROM senders")
    senders = cursor.fetchall()

    cursor.close()
    conn.close()

    if request.method == 'POST':
        driver_id = request.form.get('drivers', 'не указан')
        transport_number = request.form.get('transport', 'не указан')
        sender_id = request.form.get('senders', 'не указан')
        address_id = request.form.get('addresses', 'не указан')
        trailer_id = request.form.get('trailer', 'не указан')
        laboratory = request.form.get('laboratory', 'не указан')  # Получаем выбранного лаборанта
        raw_material =request.form.get('raw_material', 'не указан')
        delivery_method = request.form.get('delivery_method', 'не указан')


        # Получаем номер ттн, дату и  разбиваем на число, месяц и год
        laboratory = request.form.get('laboratory')
        ttn = request.form.get('ttn')  # Номер ТТН
        physical_weight = request.form.get('physical_weight')
        date_input = request.form.get('date')  # Получаем дату в формате YYYY-MM-DD
        inn = request.form.get('inn')
        razgruzka = request.form.get('razgruzka')
        if date_input:
            date_obj = datetime.strptime(date_input, '%Y-%m-%d')
            day = f"{date_obj.day:02}"  # Форматирование с ведущим нулем
            month = f"{date_obj.month:02}"  # Форматирование с ведущим нулем
            year = date_obj.year
        else:
            day = month = year = None

        # Получаем веса из секций
        section_weights = []
        for i in range(1, 8):
            weight = request.form.get(f'section_weight_{i}')
            section_weights.append(weight)

        if driver_id and transport_number and sender_id and address_id and trailer_id:
            # Получаем инициалы и полное имя водителя
            conn = db_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT initials, full_name FROM drivers WHERE id=%s", (driver_id,))
            driver_info = cursor.fetchone()  # Получаем инициалы и полное имя водителя

            cursor.execute("SELECT brand FROM transport WHERE transport_number=%s", (transport_number,))
            brand = cursor.fetchone()

            cursor.execute("SELECT name FROM senders WHERE id=%s", (sender_id,))
            sender_name = cursor.fetchone()

            cursor.execute("SELECT address FROM addresses WHERE id=%s", (address_id,))
            address = cursor.fetchone()

            cursor.execute("SELECT trailer_number FROM trailers WHERE id=%s", (trailer_id,))
            trailer_number = cursor.fetchone()[0]  # Получаем номер прицепа

            cursor.close()
            conn.close()

            if driver_info and brand and sender_name and address:
                # Загружаем существующий Excel файл
                wb = load_workbook(r'C:\Users\oleg.d\PycharmProjects\New_project\Excel_Project\template.xlsx')
                ws_main = wb.active  # Первый лист
                #ws_second = wb['стр2']  # Замените на фактическое имя второго листа

                # Вставляем данные в нужные ячейки первого листа
                ws_main["G13"] = driver_info[1]  # Полное имя водителя
                ws_main["AN51"] = driver_info[0]  # Инициалы водителя
                ws_main["CY49"] = driver_info[0]  # Инициалы водителя
                ws_main["K7"] = sender_name[0]  # Имя отправителя
                ws_main["K18"] = sender_name[0]  # Имя отправителя
                ws_main["K21"] = address[0]  # Адрес

                ws_main["H30"] = raw_material
                ws_main["AX13"] = delivery_method
                ws_main["K23"] = request.form.get('recipient')  # Получатель
                ws_main["H15"] = inn  # ИНН
                ws_main["K25"] = razgruzka  # Адрес разгрузки
                ws_main["BG3"] = ttn  # Номер ТТН в ячейку FM6
                ws_main["BD8"] = trailer_number  # Номер прицепа
                ws_main["AN49"] = laboratory  # Номер прицепа
                current_row = 35  # Начальная строка для заполнения
                for i, weight in enumerate(section_weights):
                    if weight:  # Если вес секции существует
                        ws_main[f"A{current_row}"] = i + 1  # Номер любой секции
                        ws_main[f"D{current_row}"] = weight  # Заполнение веса секции
                        current_row += 1  # Переход к следующему ряду

                # Очищаем оставшиеся ячейки (A38-A41) если они не заполнены
                for j in range(current_row, 42):
                    ws_main[f"A{j}"] = None
                    ws_main[f"D{j}"] = None
                    ws_main["BO35"] = physical_weight # сумма веса с секций

                ws_main["K8"] = brand[0]  # Заполняем ячейку CO4 (марка)
                ws_main["AM8"] = transport_number  # Заполняем ячейку EL4 (номер)
                # Вставка даты в ячейки первого листа
                ws_main["Y6"] = day
                ws_main["AE6"] = month
                ws_main["AU6"] = year

                current_row = 35  # Начальная строка для заполнения

                # Массив атрибутов для секций
                attributes = [
                    'fat_content', 'protein_content', 'acidity', 'temperature',
                    'density', 'cell_content', 'purity_group', 'heat_resistance', 'grade'
                ]

                # Список для заполненных секций
                filled_sections = []

                # Заполнение ячеек для каждой секции
                for index in range(len(section_weights)):
                    weight = section_weights[index]
                    # Проверка, заполнена ли секция
                    if weight:  # Если вес секции существует
                        filled_sections.append(index)  # Сохраняем номер секции
                        ws_main[f"A{current_row}"] = index + 1  # Номер секции
                        ws_main[f"D{current_row}"] = weight  # Заполнение веса секции

                        # Заполнение атрибутов
                        for attr in attributes:
                            value = request.form.get(f'{attr}_{index + 1}', '')
                            if attr == 'fat_content':
                                ws_main[f'J{current_row}'] = value  # Массовая доля жира %
                            elif attr == 'protein_content':
                                ws_main[f'P{current_row}'] = value  # Массовая доля белка %
                            elif attr == 'acidity':
                                ws_main[f'V{current_row}'] = value  # Кислотность °Т
                            elif attr == 'temperature':
                                ws_main[f'AB{current_row}'] = value  # Температура °С
                            elif attr == 'density':
                                ws_main[f'AH{current_row}'] = value  # Плотность кг/м3
                            elif attr == 'cell_content':
                                ws_main[f'AN{current_row}'] = value  # Содер. Самат. Клеток, тыс/см3
                            elif attr == 'purity_group':
                                ws_main[f'AT{current_row}'] = value  # Группа чистоты
                            elif attr == 'heat_resistance':
                                ws_main[f'AZ{current_row}'] = value  # Термоустойчивочть, группа
                            elif attr == 'grade':
                                ws_main[f'BF{current_row}'] = value  # Сорт

                        current_row += 1  # Переход к следующему ряду

                # Очищаем оставшиеся ячейки, если они не заполнены
                for j in range(current_row, 42):
                    ws_main[f"A{j}"] = None
                    ws_main[f"D{j}"] = None

                # Заполнение ячеек на втором листе
                #ws_second["CO4"] = brand[0]  # Заполняем ячейку CO4 (марка)
                #ws_second["EL4"] = transport_number  # Заполняем ячейку EL4 (номер)

                excel_path = r'C:\Users\oleg.d\PycharmProjects\New_project\Excel_Project\updated_drivers_info.xlsx'
                wb.save(excel_path)

                # Генерация PDF из заполненного Excel
                pdf_path = convert_excel_to_pdf(excel_path)

                return send_file(pdf_path, as_attachment=False, download_name='document.pdf', mimetype='application/pdf')
    # Передаем данные на шаблон
    return render_template(
        'index.html',
        drivers=drivers,
        transports=[t[0] for t in transports],
        senders=senders,
        laboratories=laboratories,  # Передаем данные лаборантов
        recipients=delivery_data
    )

def convert_excel_to_pdf(excel_path):
    pythoncom.CoInitialize()  # Инициализация COM
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(excel_path)

    pdf_path = excel_path.replace('.xlsx', '.pdf')
    wb.ExportAsFixedFormat(0, pdf_path)  # 0 означает xlTypePDF
    wb.Close(False)
    excel.Application.Quit()

    return pdf_path
#Новый маршрут для обработки POST-запроса, который будет получать данные из формы и записывать их в базу данных.
@app.route('/submit-data', methods=['POST'])
def submit_data():
    driver_full_name = request.form.get('driver_full_name')
    driver_initials = request.form.get('driver_initials')

    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute("""
            INSERT INTO drivers (full_name, initials) VALUES (%s, %s) RETURNING id
        """, (driver_full_name, driver_initials))
    driver_id = cursor.fetchone()[0]  # Получаем ID водителя
    conn.commit()  # Сохраняем изменения
    cursor.close()
    conn.close()

    return "Данные успешно сохранены!"  # Можно заменить на redirect на нужную страницу
@app.route('/get_addresses/<int:sender_id>', methods=['GET'])
def get_addresses(sender_id):
    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, address FROM addresses WHERE sender_id = %s", (sender_id,))
    addresses = cursor.fetchall()

    cursor.close()
    conn.close()
    return jsonify(addresses)

@app.route('/trailers', methods=['GET'])
def trailers():
    trailer_data = get_trailer_data()
    return jsonify(trailer_data)

@app.route('/data-entry', methods=['GET'])
def data_entry():
    # Здесь вы можете определить логику, которую хотите использовать на странице ввода данных.
    return render_template('data_entry.html')  # Создайте новый шаблон для этой страницы


if __name__ == '__main__':
    app.run(debug=True)