from flask import Flask, render_template, request, send_file, jsonify
import psycopg2
from openpyxl import load_workbook
from datetime import datetime
import pythoncom
import re
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

def get_addresses():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, address FROM addresses")
    addresses = cursor.fetchall()
    cursor.close()
    conn.close()
    return [{'id': row[0], 'address': row[1]} for row in addresses]

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
    addresses = get_addresses()

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
        driver_id = request.form.get('drivers', None)
        transport_number = request.form.get('transport', None)
        sender_id = request.form.get('senders', None)
        trailer_id = request.form.get('trailer', None)
        raw_material = request.form.get('raw_material', None)
        vladelec = request.form.get('vladelec', None)
        delivery_method = request.form.get('delivery_method', None)
        addresses = request.form.get('addresses', None)
        laboratory = request.form.get('laboratory')
        output_format = request.form.get('output_format', 'excel')  # Получаем выбранный формат

        # Получаем номер ттн, дату и  разбиваем на число, месяц и год
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
        section_weights = [request.form.get(f'section_weight_{i}', None) for i in range(1, 8)]

        conn = db_connection()
        cursor = conn.cursor()

        if driver_id:
            cursor.execute("SELECT initials, full_name FROM drivers WHERE id=%s", (driver_id,))
            driver_info = cursor.fetchone()
        else:
            driver_info = (None, None)

        if transport_number:
            cursor.execute("SELECT brand FROM transport WHERE transport_number=%s", (transport_number,))
            brand = cursor.fetchone()
        else:
            brand = (None,)

        if sender_id:
            cursor.execute("SELECT name FROM senders WHERE id=%s", (sender_id,))
            sender_name = cursor.fetchone()
        else:
            sender_name = (None,)

        if trailer_id:
            cursor.execute("SELECT trailer_number FROM trailers WHERE id=%s", (trailer_id,))
            trailer_number = cursor.fetchone()[0]
        else:
            trailer_number = None

        cursor.close()
        conn.close()

        wb = load_workbook(r'C:\Users\oleg.d\PycharmProjects\New_project\Excel_Project\template.xlsx')
        ws_main = wb.active  # Первый лист
        #ws_second = wb['стр2']  # Замените на фактическое имя второго листа

        # Вставляем данные в нужные ячейки первого листа
        ws_main["G13"] = driver_info[1]  # Полное имя водителя
        ws_main["AN51"] = driver_info[0]  # Инициалы водителя
        ws_main["CY49"] = driver_info[0]  # Инициалы водителя
        ws_main["K7"] = sender_name[0]  # Имя отправителя
        ws_main["K18"] = sender_name[0]  # Имя отправителя
        ws_main["K21"] = addresses
        ws_main["H30"] = raw_material
        ws_main["W10"] = vladelec
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
                ws_main[f"BL{current_row}"] = i + 1  # Запись номера секции в ячейку BL
                current_row += 1  # Переход к следующему ряду

        # Очищаем оставшиеся ячейки (A38-A41) если они не заполнены
        for j in range(current_row, 42):
            ws_main[f"A{j}"] = None
            ws_main[f"D{j}"] = None
            ws_main["E43"] = physical_weight # сумма веса с секций

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

        if output_format == 'pdf':
            # Генерация PDF из заполненного Excel
            pdf_path = convert_excel_to_pdf(excel_path)
            return jsonify({'pdf_url': pdf_path.split('\\')[-1]})  # Отправляем имя PDF
        elif output_format == 'excel':
            return jsonify({'excel_url': 'updated_drivers_info.xlsx'})  # Отправляем имя Excel
# Передаем данные на шаблон
    return render_template(
        'index.html',
        drivers=drivers,
        transports=[t[0] for t in transports],
        senders=senders,
        laboratories=laboratories,  # Передаем данные лаборантов
        recipients=delivery_data,
        addresses=addresses,
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
@app.route('/files/<path:filename>', methods=['GET'])
def uploaded_file(filename):
    return send_file(r'C:\Users\oleg.d\PycharmProjects\New_project\Excel_Project\\' + filename, as_attachment=True)

@app.route('/submit-driver-data', methods=['POST'])
def submit_data():
    driver_full_name = request.form.get('driver_full_name')
    # Приводим полное имя к формату "С Заглавной Буквы" для каждого слова
    driver_full_name = driver_full_name.title()

    # Генерация инициалов
    names = driver_full_name.split()
    if len(names) >= 3:
        driver_initials = f"{names[0]} {names[1][0].upper()}.{names[2][0].upper()}."
    else:
        return jsonify({"message": "Недостаточно данных для генерации инициалов."}), 400

    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute("""
            INSERT INTO drivers (full_name, initials) VALUES (%s, %s) RETURNING id
        """, (driver_full_name, driver_initials))
    driver_id = cursor.fetchone()[0]  # Получаем ID водителя
    conn.commit()  # Сохраняем изменения
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные водителя успешно сохранены! Нажми назад и обнови предыдущую страницу!"})  # Возвращаем JSON-ответ

@app.route('/submit-address' , methods=['POST'])
def submit_address():
    address = request.form.get('address')
    if not address:
        return jsonify({"message": "Не указаны все обязательные данные для лаборанта."}), 400
    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute('INSERT INTO addresses (address) VALUES (%s)', (address,))
    conn.commit()
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные о пункте погрузки успешно сохранены! Нажми назад и обнови предыдущую страницу!"})

@app.route('/submit-senders' , methods=['POST'])
def submit_senders():
    name = request.form.get('senders')
    if not name:
        return  jsonify({"message": "Не указаны все обязательные данные для лаборанта."}), 400
    conn = db_connection()
    cursor = conn.cursor()

    cursor.execute('INSERT INTO senders (name) VALUES (%s)', (name,))
    conn.commit()
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные о грузоотправителе успешно сохранены! Нажми назад и обнови предыдущую страницу!"})

@app.route('/submit-transport-data', methods=['POST'])
def submit_transport_data():
    transport_number = request.form.get('transport_number', '').strip()  # Удаляем пробелы
    brand = request.form.get('brand')
    # Проверка на наличие каждого из обязательных полей
    if not transport_number or not brand:
        return jsonify({"message": "Не указаны все обязательные данные для транспорта."}), 400

    # Приведение данных к верхнему регистру
    transport_number = transport_number.upper()

    try:
        conn = db_connection()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO transport (transport_number, brand) VALUES (%s, %s)', (transport_number, brand))
        conn.commit()
    except Exception as e:
        return jsonify({"message": f"Ошибка при сохранении данных: {str(e)}"}), 500

    finally:
        cursor.close()
        conn.close()

    return jsonify({"message": "Данные транспорта успешно сохранены! Нажми назад и обнови предыдущую страницу!"})

@app.route('/submit-laboratory', methods=['POST'])
def submit_laboratory_data():
    laboratory_name = request.form.get('laboratory')
    # Проверка на наличие обязательного поля
    if not laboratory_name:
        return jsonify({"message": "Не указаны все обязательные данные для лаборанта."}), 400

    conn = db_connection()
    cursor = conn.cursor()

    # Вставка данных в таблицу laboratory
    cursor.execute('INSERT INTO laboratory (name) VALUES (%s)', (laboratory_name,))
    conn.commit()
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные лаборанта успешно сохранены! Нажми назад и обнови предыдущую страницу!"})


@app.route('//submit-delivery', methods=['POST'])
def submit_delivery():
    recipient = request.form.get('recipient')
    inn = request.form.get('inn')
    razgruzka = request.form.get('razgruzka')
    # Проверка на наличие каждого из обязательных полей
    if not recipient or not inn or not razgruzka:
        return jsonify({"message": "Не указаны все обязательные данные для транспорта."}), 400

    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute('INSERT INTO delivery (recipient, inn, razgruzka) VALUES (%s, %s, %s)', (recipient, inn, razgruzka))
    conn.commit()
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные о грузополучателе успешно сохранены! Нажми назад и обнови предыдущую страницу!"})

@app.route('/submit-trailer-data', methods=['POST'])
def submit_trailer_data():
    trailer_number = request.form.get('trailer_number')
    section1 = request.form.get('section1') or None
    section2 = request.form.get('section2') or None
    section3 = request.form.get('section3') or None
    section4 = request.form.get('section4') or None
    section5 = request.form.get('section5') or None
    section6 = request.form.get('section6') or None
    section7 = request.form.get('section7') or None

    # Преобразуем все данные к верхнему регистру
    trailer_number = trailer_number.upper() if trailer_number else None
    section1 = section1.upper() if section1 else None
    section2 = section2.upper() if section2 else None
    section3 = section3.upper() if section3 else None
    section4 = section4.upper() if section4 else None
    section5 = section5.upper() if section5 else None
    section6 = section6.upper() if section6 else None
    section7 = section7.upper() if section7 else None

    conn = db_connection()
    cursor = conn.cursor()

    # Выполняем вставку в таблицу trailers
    cursor.execute("""
        INSERT INTO trailers (trailer_number, section1, section2, section3, section4, section5, section6, section7)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """, (trailer_number, section1, section2, section3, section4, section5, section6, section7))

    conn.commit()
    cursor.close()
    conn.close()

    return jsonify({"message": "Данные о прицепе успешно сохранены!"})

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