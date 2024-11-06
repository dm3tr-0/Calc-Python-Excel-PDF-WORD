from flask import Flask, render_template, request, jsonify, send_file, session
from settings import * #тут зашитые значения
import math
import os
from io import BytesIO
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate
from reportlab.lib import colors

app = Flask(__name__)
app.secret_key = os.urandom(24) #ключ защищающий сессию пользователя

@app.route('/')
def index():
    return render_template('index.html')

#временная функция для улучшения читабельночти вывода результатов
def format_dicts_side_by_side(dicts):
    """Форматирует словари, подгоняя их под ключи 180h, 168h, 79h, 180h_pr, 180h_night"""


    column_keys = ['180h', '168h', '79h', '180h_pr', '180h_night']


    table = {key: [] for key in column_keys}


    for d in dicts:
        for key, value in d.items():
            if key in column_keys:
                table[key].append(value)
            elif key == '180-79h':
                table['180h'].append(value)
                table['168h'].append(value)
                table['79h'].append(value)
            elif key == '180h_pr_night':
                table['180h_pr'].append(value)
                table['180h_night'].append(value)
            else:
                for k in column_keys:
                    table[k].append('')

    return table
    #rows = []
    #for key in column_keys:
    #    row = f"{key}: " + " | ".join(str(v) for v in table[key])
    #    rows.append(row)

    #return "<br>".join(rows)

#факт расчет
def fact(total_files, day_files, night_files, day_pr_files, machines_180h, machines_168h, machines_79h, machines_180h_night):
    # факт среднее количество файлов в месяц для 180-79ч 180ч пр/вых 180чночь
    fact_mid_files_month = {
        '180-79h': day_files,
        '180h_pr': day_pr_files,
        '180h_night': night_files
    }

    # факт кол-во машин для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_mashine_count = {
        '180h': machines_180h,
        '168h': machines_168h,
        '79h': machines_79h,
        '180h_pr': machines_180h_night,  # надо уточнить берем ли мы тут значение из F17(см калькулятор.xls)
        '180h_night': machines_180h_night
    }

    # факт максимальное количество файлов для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_max_files = {
        '180h': round(fact_mashine_count['180h'] * max_files_month['180h']),
        '168h': round(fact_mashine_count['168h'] * max_files_month['168h']),
        '79h': round(fact_mashine_count['79h'] * max_files_month['79h']),
        '180h_pr': round(fact_mashine_count['180h_pr'] * max_files_month['180h_pr']),
        '180h_night': round(fact_mashine_count['180h_night'] * max_files_month['180h_night'])
    }

    # факт разница нагрузки
    fact_diff_stress = {
        '180-79h': round(
            fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - fact_mid_files_month['180-79h']),
        '180h_pr': round(fact_max_files['180h_pr'] - fact_mid_files_month['180h_pr']),
        '180h_night': round(fact_max_files['180h_night'] - fact_mid_files_month['180h_night'])
    }

    # факт нагрузка в %
    fact_percent_stress = {
        '180-79h': round(100 * fact_mid_files_month['180-79h'] / (
                    fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'])),
        '180h_pr': round(100 * fact_mid_files_month['180h_pr'] / fact_max_files['180h_pr']),
        '180h_night': round(100 * fact_mid_files_month['180h_night'] / fact_max_files['180h_night'])
    }

    # факт нехватка машин
    fact_mashine_loss = {
        '180h': 'black',
        '168h': 0,
        '79h': -1,
        '180h_pr_night': 0
    }
    if fact_percent_stress['180-79h'] < 86:
        fact_mashine_loss['168h'] = 0
    else:
         fact_mashine_loss['168h'] = round( ((fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'])  * (fact_percent_stress['180-79h'] - 86)) / 86 / max_files_month['168h'])





    if fact_mid_files_month['180-79h'] / (fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] + fact_mashine_loss['168h'] * max_files_month['168h']) < 0.86:
        fact_mashine_loss['79h'] = 0
    else:
        fact_mashine_loss['79h'] = math.ceil((fact_mid_files_month['180-79h'] / (fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] + fact_mashine_loss['168h'] * max_files_month['168h']) - 0.86) * (fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] + fact_mashine_loss['168h'] * max_files_month['168h']) / 86 / max_files_month['79h'])






    if fact_diff_stress['180h_night'] < 0:
        fact_mashine_loss['180h_pr_night'] = round(fact_max_files['180h_night'] / max_files_month['180h_night'] * 1.8)
    elif round((fact_percent_stress['180h_night'] + fact_percent_stress['180h_pr']) / 2) < 80:
        fact_mashine_loss['180h_pr_night'] = 0
    else:
        fact_mashine_loss['180h_pr_night'] = round(max_files_month['180h_night'] / fact_max_files['180h_night'])





    data[0] = format_dicts_side_by_side([fact_mid_files_month, fact_mashine_count, fact_max_files, fact_diff_stress, fact_percent_stress, fact_mashine_loss])

    fact_result = (
        #"Факт: <br><br>"
        #f"{format_dicts_side_by_side([fact_mid_files_month, fact_mashine_count, fact_max_files, fact_diff_stress, fact_percent_stress, fact_mashine_loss])}<br><br>"
        f"Фактическая нехватка машин<br>"
        f"для 168ч: {fact_mashine_loss['168h']}<br>"
        f"для 79ч: {fact_mashine_loss['79h']}<br>"
        f"для 180ч пр/ночь: {fact_mashine_loss['180h_pr_night']}<br><br>"
    )


    return fact_result
#план расчет
def plan(total_files, day_files, night_files, day_pr_files, machines_180h, machines_168h, machines_79h, machines_180h_night, new_users):
    # факт среднее количество файлов в месяц для 180-79ч 180ч пр/вых 180чночь
    fact_mid_files_month = {
        '180-79h': day_files,
        '180h_pr': day_pr_files,
        '180h_night': night_files
    }

    # факт кол-во машин для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_mashine_count = {
        '180h': machines_180h,
        '168h': machines_168h,
        '79h': machines_79h,
        '180h_pr': machines_180h_night,  # надо уточнить берем ли мы тут значение из F17(см калькулятор.xls)
        '180h_night': machines_180h_night
    }

    # факт максимальное количество файлов для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_max_files = {
        '180h': round(fact_mashine_count['180h'] * max_files_month['180h']),
        '168h': round(fact_mashine_count['168h'] * max_files_month['168h']),
        '79h': round(fact_mashine_count['79h'] * max_files_month['79h']),
        '180h_pr': round(fact_mashine_count['180h_pr'] * max_files_month['180h_pr']),
        '180h_night': round(fact_mashine_count['180h_night'] * max_files_month['180h_night'])
    }

    # факт разница нагрузки
    fact_diff_stress = {
        '180-79h': round(
            fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - fact_mid_files_month['180-79h']),
        '180h_pr': round(fact_max_files['180h_pr'] - fact_mid_files_month['180h_pr']),
        '180h_night': round(fact_max_files['180h_night'] - fact_mid_files_month['180h_night'])
    }

    # факт нагрузка в %
    fact_percent_stress = {
        '180-79h': round(100 * fact_mid_files_month['180-79h'] / (
                    fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'])),
        '180h_pr': round(100 * fact_mid_files_month['180h_pr'] / fact_max_files['180h_pr']),
        '180h_night': round(100 * fact_mid_files_month['180h_night'] / fact_max_files['180h_night'])
    }

    # факт нехватка машин
    fact_mashine_loss = {
        '180h': 'black',
        '168h': 0,
        '79h': -1,
        '180h_pr_night': 0
    }


    # среднее количесво файлов новых уз в месяц
    plan_mid_fileUZ_month = {
        '180-79h': 'задастся ниже',
        '180h_pr': round((new_users * 1.3) * 0.15),
        '180h_night': round((new_users * 1.3) * 0.27)
    }
    plan_mid_fileUZ_month['180-79h'] = round(
        new_users * 1.3 - plan_mid_fileUZ_month['180h_night'] - plan_mid_fileUZ_month['180h_pr'])

    # среднее кол-во файлов с учетом новых уз в месяц
    plan_mid_newfiles_month = {
        '180-79h': round(fact_mid_files_month['180-79h'] + plan_mid_fileUZ_month['180-79h']),
        '180h_pr': round(fact_mid_files_month['180h_pr'] + plan_mid_fileUZ_month['180h_pr']),
        '180h_night': round(fact_mid_files_month['180h_night'] + plan_mid_fileUZ_month['180h_night'])
    }

    # факт кол-во машин уже указано в fact_mashine_count

    # факт максимальное кол-во файлов указано в fact_max_files

    # пданируемая разница нагрузки
    plan_diff_stress = {
        '180-79h': round(
            fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - plan_mid_newfiles_month[
                '180-79h']),
        '180h_pr': round(fact_max_files['180h_pr'] - plan_mid_newfiles_month['180h_pr']),
        '180h_night': round(fact_max_files['180h_night'] - plan_mid_newfiles_month['180h_night'])
    }

    # пданируемая разница нагрузки в процентах
    plan_percent_stress = {
        '180-79h': round(100 * plan_mid_newfiles_month['180-79h'] / (
                    fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'])),
        '180h_pr': round(100 * plan_mid_newfiles_month['180h_pr'] / fact_max_files['180h_pr']),
        '180h_night': round(100 * plan_mid_newfiles_month['180h_night'] / fact_max_files['180h_night'])
    }

    # планируемая нехватка машин
    plan_mashine_loss = {
        '180h': 'black',
        '168h': -1,
        '79h': -2,
        '180h_pr_night': 0
    }

    if plan_percent_stress['180-79h'] < 86:
        plan_mashine_loss['168h'] = 0
    else:
        plan_mashine_loss['168h'] = round((fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - plan_mid_newfiles_month['180-79h']) / max_files_month['168h'])

    if plan_percent_stress['180-79h'] < 86:
        plan_mashine_loss['79h'] = 0
    else:
        plan_mashine_loss['79h'] = round((fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - plan_mid_newfiles_month['180-79h']) / max_files_month['79h'] * 1.5)



    if plan_diff_stress['180h_night'] < 0:
        plan_mashine_loss['180h_pr_night'] = round(fact_max_files['180h_night'] / max_files_month['180h_night'] * 1.8)
    elif plan_percent_stress['180h_night'] < 80:
        plan_mashine_loss['180h_pr_night'] = 0
    else:
        plan_mashine_loss['180h_pr_night'] = round(max_files_month['180h_night'] / fact_max_files['180h_night'])



    data[1] = format_dicts_side_by_side([fact_mid_files_month, {'180h': new_users, '168h': new_users, '79h': new_users,'180h_pr': new_users,'180h_night': new_users}, plan_mid_fileUZ_month, plan_mid_newfiles_month, fact_mashine_count, fact_max_files, plan_diff_stress, plan_percent_stress, plan_mashine_loss])

    plan_result = (
        #"План: <br><br>"
        #f"{format_dicts_side_by_side([fact_mid_files_month, {'180h': new_users, '168h': new_users, '79h': new_users,'180h_pr': new_users,'180h_night': new_users}, plan_mid_fileUZ_month, plan_mid_newfiles_month, fact_mashine_count, fact_max_files, plan_diff_stress, plan_percent_stress, plan_mashine_loss])}<br>"
        f"Планируемая нехватка машин<br>"
        f"для 168ч: {plan_mashine_loss['168h']}<br>"
        f"для 79ч: {plan_mashine_loss['79h']}<br>"
        f"для 180ч пр/ночь: {plan_mashine_loss['180h_pr_night']}<br>"
    )
    return plan_result



#тут хранятся данные для экспорта, полученные после расчета
data = [{}, {}]

#экспорт результатов в табличку эксель
@app.route('/export_excel')
def export_excel():
    # Получаем ключи и определяем количество колонок
    keys = list(data[0].keys())
    num_columns = max(len(values) for dict_data in data for values in dict_data.values())
    columns = [f'Col{i + 1}' for i in range(num_columns)]

    # Формируем строки для таблицы
    rows = []
    for key in keys:
        row = [key]  # Добавляем ключ как первую ячейку
        for i in range(len(data)):
            values = data[i].get(key, [])
            row.extend(
                values + [""] * (num_columns - len(values)))  # Заполняем пустыми ячейками, если не хватает значений
        rows.append(row)

    # Заголовки для таблицы, добавляем Value1 и Value2
    headers = ["Key"] + [f"Value 1 {col}" for col in columns] + [f"Value 2 {col}" for col in columns]

    # Создаем DataFrame
    df = pd.DataFrame(rows, columns=headers)

    # Экспорт в Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    # Сохраняем в сессии
    session['excel_file'] = output.read()
    output.close()

    return send_file(BytesIO(session['excel_file']), as_attachment=True, download_name="report.xlsx")


#экспорт результатов в файл пдф
@app.route('/export_pdf')
def export_pdf():
    # Получаем ключи и количество колонок
    keys = list(data[0].keys())
    num_columns = max(len(values) for dict_data in data for values in dict_data.values())
    columns = [f'Col{i + 1}' for i in range(num_columns)]

    headers = ["Key"] + [f"Value 1 {col}" for col in columns] + [f"Value 2 {col}" for col in columns]
    table_data = [headers]

    for key in keys:
        row = [key]
        for i in range(len(data)):
            values = data[i].get(key, [])
            row.extend(values + [""] * (num_columns - len(values)))
        table_data.append(row)

    output = BytesIO()
    pdf = SimpleDocTemplate(output, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20,
                            bottomMargin=20)
    table = Table(table_data)

    # Стиль таблицы с настройкой ширины
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 6),  # Уменьшаем размер шрифта
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
    ])
    table.setStyle(style)

    # Автоматическая подгонка ширины таблицы
    table._argW = [pdf.width / len(table_data[0])] * len(table_data[0])

    pdf.build([table])
    output.seek(0)

    session['pdf_file'] = output.read()
    output.close()

    return send_file(BytesIO(session['pdf_file']), as_attachment=True, download_name="report.pdf")

#эта функция срабатывает, когда нажали кнопку рассчитать(в index.html)
@app.route('/calculate', methods=['POST'])
def calculate():
    data[0].clear()
    data[1].clear()
    ###########################
    # Получаем данные из формы:
    ###########################

    try:
        #факт среднее количество файлов в месяц
        total_files = int(request.form['total_files']) #общее
        day_files = int(request.form['day_files']) #день
        night_files = int(request.form['night_files']) #ночь/пр/вых
        day_pr_files = int(request.form['day_pr_files']) #день/пр/вых

        # факт количетсво машин
        machines_180h = int(request.form['machines_180h'])  # 180 часов
        machines_168h = int(request.form['machines_168h'])  # 168 часов
        machines_79h = int(request.form['machines_79h'])  # 79 часов
        machines_180h_night = int(request.form['machines_180h_night'])  # 180 часов ночь
    except:
        return jsonify(result= ("вы ввели не все данные"))

    isNewUsersexist = False
    #Кол-во новых пользователей (УЗ)
    try:
        new_users = int(request.form['new_users']) #общее
        isNewUsersexist = True
    except:
        pass





    ###############
    #Расчеты:
    ###############

    ###############
    #расчет факт
    ###############

    fact_result = fact(total_files, day_files, night_files, day_pr_files, machines_180h, machines_168h, machines_79h, machines_180h_night)


    ###############
    #рассчет план
    ###############
    if isNewUsersexist:
        plan_result = plan(total_files, day_files, night_files, day_pr_files, machines_180h, machines_168h, machines_79h,
                       machines_180h_night, new_users)
    else:
        plan_result = ('')





    print(data)
    ###########################################
    # Создаём итоговый результат в виде строк
    ###########################################
    result = fact_result + plan_result
    #возвращаем полученные расчеты (они отобразяться в Итоговые результаты id=result-output)
    return jsonify(result=result)


if __name__ == '__main__':
    app.run(debug=True)
