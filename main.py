from flask import Flask, render_template, request, jsonify

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/calculate', methods=['POST'])
def calculate():
    data = request.json
    work_hours = float(data['work_hours'])
    days_count = float(data['days_count'])
    max_files_day = float(data['max_files_day'])

    # Пример расчетов
    max_files_month = max_files_day * days_count
    fact_files_month = work_hours * 0.9  # Условный расчет
    load_difference = max_files_month - fact_files_month

    result = {
        'max_files_month': max_files_month,
        'fact_files_month': fact_files_month,
        'load_difference': load_difference
    }

    return jsonify(result)


if __name__ == '__main__':
    app.run(debug=True)
