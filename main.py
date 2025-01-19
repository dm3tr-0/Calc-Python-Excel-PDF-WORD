import os
from flask import Flask, render_template

# запускаем приложение
app = Flask(__name__)
app.secret_key = os.urandom(24)  # ключ защищающий сессию пользователя


@app.route('/')  # главная страница сайта
def index():
    return render_template('index.html')


# импортируем методы для 1го режима
from first_mode import *


@app.route('/export_excel')
def export_to_excel():
    return export_excel()


@app.route('/export_pdf')
def export_to_pdf():
    return export_pdf()


@app.route('/export_word')
def export_to_word():
    return export_word()


@app.route('/calculate', methods=['POST'])
def calculate_first_mode():
    return calculate()


if __name__ == '__main__':
    app.run(debug=True)
