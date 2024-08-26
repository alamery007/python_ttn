from flask import Flask, render_template, request, send_file, jsonify
import psycopg2
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
from fpdf import FPDF

app = Flask(__name__)

def db_connection():
    conn = psycopg2.connect(
        host='localhost',
        database='form_tth',
        user='postgres',
        password='123456'
    )
    return conn
def get_laboratories():
    conn = db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM laboratory")
    laboratories = cursor.fetchall()
    cursor.close()
    conn.close()
    return [{'id': row[0], 'name': row[1]} for row in laboratories]

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

@app.route('/', methods=['GET', 'POST'])
def index():
    drivers = []
    transports = []
    senders = []
    laboratories = get_laboratories()

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
        driver_id = request.form.get('drivers')
        transport_number = request.form.get('transport')
        sender_id = request.form.get('senders')
        address_id = request.form.get('addresses')
        trailer_id = request.form.get('trailer')
        laboratory = request.form.get('laboratory')  # Получаем выбранного лаборанта

        # Получаем номер ттн, дату и  разбиваем на число, месяц и год
        laboratory = request.form.get('laboratory')
        ttn = request.form.get('ttn')  # Номер ТТН
        series = request.form.get('series')
        physical_weight = request.form.get('physical_weight')
        date_input = request.form.get('date')  # Получаем дату в формате YYYY-MM-DD
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
                ws_second = wb['стр2']  # Замените на фактическое имя второго листа

                # Вставляем данные в нужные ячейки первого листа
                ws_main["L18"] = driver_info[1]  # Полное имя водителя
                ws_main["AF18"] = driver_info[0]  # Инициалы водителя
                ws_main["L20"] = sender_name[0]  # Имя отправителя
                ws_main["AF20"] = address[0]  # Адрес

                ws_main["FM6"] = ttn  # Номер ТТН в ячейку FM6
                ws_main["DR6"] = series
                ws_main["A22"] = trailer_number  # Номер прицепа
                ws_main["FI36"] = laboratory  # Номер прицепа
                # Заполнение секций в соответствующие ячейки первого листа
                ws_main["L25"] = section_weights[0] if section_weights[0] else None
                ws_main["AF25"] = section_weights[1] if section_weights[1] else None
                ws_main["AS25"] = section_weights[2] if section_weights[2] else None
                ws_main["BE25"] = section_weights[3] if section_weights[3] else None
                ws_main["BU25"] = section_weights[4] if section_weights[4] else None
                ws_main["DH25"] = section_weights[5] if section_weights[5] else None
                ws_main["DS25"] = section_weights[6] if section_weights[6] else None
                ws_main["CR29"] = physical_weight # сумма веса с секций
                # Вставка даты в ячейки первого листа
                ws_main["FM7"] = day
                ws_main["FS7"] = month
                ws_main["FZ7"] = year

                # Заполнение ячеек на втором листе
                ws_second["CO4"] = brand[0]  # Заполняем ячейку CO4 (марка)
                ws_second["EL4"] = transport_number  # Заполняем ячейку EL4 (номер)
                # Сохранение файла в BytesIO для отправки
                output = BytesIO()
                wb.save(output)
                output.seek(0)

                return send_file(output, as_attachment=True, download_name='updated_drivers_info.xlsx',
                                 mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Передаем данные на шаблон
    return render_template(
        'index.html',
        drivers=drivers,
        transports=[t[0] for t in transports],
        senders=senders,
        laboratories=laboratories  # Передаем данные лаборантов
    )

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

if __name__ == '__main__':
    app.run(debug=True)
