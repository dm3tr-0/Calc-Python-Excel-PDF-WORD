from flask import Flask, render_template, request, jsonify

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')

#эта функция срабатывает, когда нажали кнопку рассчитать(в index.html)
@app.route('/calculate', methods=['POST'])
def calculate():
    # Получаем данные из формы
    work_hours = float(request.form['work_hours'])
    days_count = float(request.form['days_count'])
    max_files_day = float(request.form['max_files_day'])

    #пример расчетов
    max_files_month = max_files_day * days_count
    fact_files_month = work_hours * 0.9  # Условный расчет
    load_difference = max_files_month - fact_files_month

    # Создаём итоговый результат в виде строк
    result = (
        f"Максимальное количество файлов в месяц: {max_files_month:.2f}<br>"
        f"Фактическое количество файлов в месяц: {fact_files_month:.2f}<br>"
        f"Разница в нагрузке: {load_difference:.2f}"
    )
    #возвращаем полученные расчеты (они отобразяться в Итоговые результаты id=result-output)
    return jsonify(result=result)


if __name__ == '__main__':
    app.run(debug=True)
