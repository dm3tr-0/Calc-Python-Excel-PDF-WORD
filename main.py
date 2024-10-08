from flask import Flask, render_template, request, jsonify

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')

#эта функция срабатывает, когда нажали кнопку рассчитать(index.html id=calculate)
@app.route('/calculate', methods=['POST'])
def calculate():
    #получаем словарь со значениями
    data = request.json
    #вытаскиваем значения по ключу и сохраняем в переменные
    work_hours = float(data['work_hours'])
    days_count = float(data['days_count'])
    max_files_day = float(data['max_files_day'])

    #пример расчетов
    max_files_month = max_files_day * days_count
    fact_files_month = work_hours * 0.9  # Условный расчет
    load_difference = max_files_month - fact_files_month

    #сохраняем полученные расчеты в словарь result
    result = {
        'max_files_month': max_files_month,
        'fact_files_month': fact_files_month,
        'load_difference': load_difference
    }
    #возвращаем полученные расчеты (они отобразяться в Итоговые результаты id=result-output)
    return jsonify(result)


if __name__ == '__main__':
    app.run(debug=True)
