from flask import Flask, render_template, request, jsonify
from settings import * #тут зашитые значения
import math

app = Flask(__name__)


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


    rows = []
    for key in column_keys:
        row = f"{key}: " + " | ".join(str(v) for v in table[key])
        rows.append(row)

    return "<br>".join(rows)

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






    fact_result = (
        "Факт: <br><br>"
        f"{format_dicts_side_by_side([fact_mid_files_month, fact_mashine_count, fact_max_files, fact_diff_stress, fact_percent_stress, fact_mashine_loss])}<br><br>"
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




    plan_result = (
        "План: <br><br>"
        f"{format_dicts_side_by_side([fact_mid_files_month, {'180h': new_users, '168h': new_users, '79h': new_users,'180h_pr': new_users,'180h_night': new_users}, plan_mid_fileUZ_month, plan_mid_newfiles_month, fact_mashine_count, fact_max_files, plan_diff_stress, plan_percent_stress, plan_mashine_loss])}<br>"
    )
    return plan_result

#экспорт результатов в табличку эксель
def export_to_excel():
    pass

#экспорт результатов в файл pdf
def export_to_pdf():
    pass


#эта функция срабатывает, когда нажали кнопку рассчитать(в index.html)
@app.route('/calculate', methods=['POST'])
def calculate():
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






    ###########################################
    # Создаём итоговый результат в виде строк
    ###########################################
    result = fact_result + plan_result
    #возвращаем полученные расчеты (они отобразяться в Итоговые результаты id=result-output)
    return jsonify(result=result)


if __name__ == '__main__':
    app.run(debug=True)
