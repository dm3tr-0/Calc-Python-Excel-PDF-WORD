from flask import Flask, render_template, request, jsonify

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')

#эта функция срабатывает, когда нажали кнопку рассчитать(в index.html)
@app.route('/calculate', methods=['POST'])
def calculate():
    ###########################
    # Получаем данные из формы:
    ###########################

    #факт среднее количество файлов в месяц
    total_files = float(request.form['total_files']) #общее
    day_files = float(request.form['day_files']) #день
    night_files = float(request.form['night_files']) #ночь/пр/вых
    day_pr_files = float(request.form['day_pr_files']) #день/пр/вых

    #Кол-во новых пользователей (УЗ)
    new_users = float(request.form['new_users']) #общее

    #факт количетсво машин
    machines_180h = float(request.form['machines_180h']) #180 часов
    machines_168h = float(request.form['machines_168h']) #168 часов
    machines_79h = float(request.form['machines_79h']) #79 часов
    machines_180h_night = float(request.form['machines_180h_night']) #180 часов ночь

    ###################
    #зашитые значения:
    ###################

    #количество часов работы для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    work_hours = {
        '180h' : 11,
        '168h' : 8,
        '79h' : 4,
        '180h_pr' : 11,
        '180h_night' : 11
    }
    #среднее количество дней для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    mid_days_count = {
        '180h': 15,
        '168h': 20,
        '79h': 20,
        '180h_pr': 15,
        '180h_night': 15
    }
    #максимальное колвичество файлов в день для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    max_files_day = {
        '180h': 58,
        '168h': 43,
        '79h': 21,
        '180h_pr': 78,
        '180h_night': 78
    }
    # максимальное колвичество файлов в месяц для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    max_files_month = {
        '180h': 877,
        '168h': 850,
        '79h': 425,
        '180h_pr': 1169,
        '180h_night': 1169
    }

    ###############
    #Расчеты:
    ###############

    ###############
    #расчет факт
    ###############

    #факт среднее количество файлов в месяц для 180-79ч 180ч пр/вых 180чночь
    fact_mid_files_month = {
        '180-79h' : day_files,
        '180h_pr' : day_pr_files,
        '180h_night' : night_files
    }

    #факт кол-во машин для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_mashine_count = {
        '180h': machines_180h,
        '168h': machines_168h,
        '79h': machines_79h,
        '180h_pr': machines_180h_night, #надо уточнить берем ли мы тут значение из F17(см калькулятор.xls)
        '180h_night': machines_180h_night
    }

    #факт максимальное количество файлов для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
    fact_max_files ={
        '180h': fact_mashine_count['180h'] * max_files_month['180h'],
        '168h': fact_mashine_count['168h'] * max_files_month['168h'],
        '79h': fact_mashine_count['79h'] * max_files_month['79h'],
        '180h_pr': fact_mashine_count['180h_pr'] * max_files_month['180h_pr'],
        '180h_night': fact_mashine_count['180h_night'] * max_files_month['180h_night']
    }

    #факт разница нагрузки
    fact_diff_stress = {
        '180-79h': fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - fact_mid_files_month['180-79h'],
        '180h_pr': fact_max_files['180h_pr'] - fact_mid_files_month['180h_pr'],
        '180h_night': fact_max_files['180h_night'] - fact_mid_files_month['180h_night']
    }

    #факт нагрузка в %
    fact_percent_stress = {
        '180-79h': 100 * fact_mid_files_month['180-79h'] / (fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h']),
        '180h_pr': 100 * fact_mid_files_month['180h_pr'] / fact_max_files['180h_pr'],
        '180h_night': 100 * fact_mid_files_month['180h_night'] / fact_max_files['180h_night']
    }

    #факт нехватка машин
    fact_mashine_loss = {
        '180h': 'black',
        '168h': 0,
        '79h': -1,
        '180h_pr_night': 0
    }
    #if fact_percent_stress['180-79h'] < 86:
    #    fact_mashine_loss['168h'] = 0
    #else:
    #    fact_mashine_loss['168h'] = ((fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h']) - fact_mid_files_month['180-79h']) / (-max_files_month['168h'])

    #if fact_percent_stress['180-79h'] < 86:
    #    fact_mashine_loss['168h'] = 0




    ###############
    #рассчет план
    ###############

    #факт среднее кол-во файлов в месяц уже есть(fact_mid_files_month)
    #кол-во новыз уз в new_users

    #среднее количесво файлов новых уз в месяц
    plan_mid_fileUZ_month = {
        '180-79h' : 'задастся ниже',
        '180h_pr' : (new_users * 1.3) * 0.15,
        '180h_night' : (new_users * 1.3) * 0.27
    }
    plan_mid_fileUZ_month['180-79h'] = new_users * 1.3 - plan_mid_fileUZ_month['180h_night'] - plan_mid_fileUZ_month['180h_pr']

    #среднее кол-во файлов с учетом новых уз в месяц
    plan_mid_newfiles_month = {
        '180-79h' : fact_mid_files_month['180-79h'] + plan_mid_fileUZ_month['180-79h'],
        '180h_pr' : fact_mid_files_month['180h_pr'] + plan_mid_fileUZ_month['180h_pr'],
        '180h_night' : fact_mid_files_month['180h_night'] + plan_mid_fileUZ_month['180h_night']
    }

    #факт кол-во машин уже указано в fact_mashine_count

    #факт максимальное кол-во файлов указано в fact_max_files

    #пданируемая разница нагрузки
    plan_diff_stress = {
        '180-79h' : fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] - plan_mid_newfiles_month['180-79h'],
        '180h_pr' : fact_max_files['180h_pr'] - plan_mid_newfiles_month['180h_pr'],
        '180h_night' : fact_max_files['180h_night'] - plan_mid_newfiles_month['180h_night']
    }

    #пданируемая разница нагрузки в процентах
    plan_percent_stress = {
        '180-79h' : 100 * plan_mid_newfiles_month['180-79h'] / (fact_max_files['180h'] + fact_max_files['168h'] + fact_max_files['79h'] ),
        '180h_pr' : 100 * plan_mid_newfiles_month['180h_pr'] / fact_max_files['180h_pr'],
        '180h_night' : 100 * plan_mid_newfiles_month['180h_night'] / fact_max_files['180h_night']
    }

    #планируемая нехватка машин
    plan_mashine_loss = {
        '180h' : 'black',
        '168h' : -1,
        '79h' : -2,
        '180h_pr_night': 0
    }







    ###########################################
    # Создаём итоговый результат в виде строк
    ###########################################
    result = (
        "Факт: <br>"
        "<br>"
        f"{fact_mid_files_month} <br>"
        f"{fact_mashine_count} <br>"
        f"{fact_max_files} <br>"
        f"{fact_diff_stress} <br>"
        f"{fact_percent_stress} <br>"
        f"{fact_mashine_loss} <br>"
        "<br>"
        "План: <br>"
        f"{fact_mid_files_month} <br>"
        f"{new_users}<br>"
        f"{plan_mid_fileUZ_month}<br>"
        f"{plan_mid_newfiles_month}<br>"
        f"{fact_mashine_count} <br>"
        f"{fact_max_files} <br>"
        f"{plan_diff_stress} <br>"
        f"{plan_percent_stress} <br>"
        f"{plan_mashine_loss} <br>"
    )
    #возвращаем полученные расчеты (они отобразяться в Итоговые результаты id=result-output)
    return jsonify(result=result)


if __name__ == '__main__':
    app.run(debug=True)
