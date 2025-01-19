# Calc-Python-Excel-PDF-WORD 📊🧮

## Описание проекта
Веб-приложение в рамках проектного практикума

🌐 **Демо**: [https://dm3tr0.pythonanywhere.com/](https://dm3tr0.pythonanywhere.com/)
🌐 **ТЗ**: [Паспорт проекта](https://github.com/user-attachments/files/17063806/-17936.pdf)
  
## 🚀 Возможности
- Выполнение расчетов
- Генерация отчетов в форматах:
  - Excel (.xlsx)
  - PDF (.pdf)
  - Word (.docx)
- Интуитивно понятный веб-интерфейс

## 🛠 Установка

### Клонирование репозитория
```bash
git clone https://github.com/dm3tr-0/Calc-Python-Excel-PDF-WORD.git
cd Calc-Python-Excel-PDF-WORD 
```

### Создание виртуального окружения
```bash
python -m venv venv
source venv/bin/activate  # Для Windows: venv\Scripts\activate
```

### Установка зависимостей
```bash
pip install -r requirements.txt
```

###🖥 Запуск приложения
```bash
flask run
```

###📂 Структура проекта

Calc-Python-Excel-PDF-WORD/
│
├── main.py              # Основной файл приложения
├── settings.py         # Настройки и константы
├── first_mode.py         # 1-й режим работы калькулятора
├── templates/          # HTML-шаблоны
│   └── index.html
├── static/             # Статические файлы
│   ├── css/
│   └── js/
└── requirements.txt    # Зависимости проекта

###🔑 Основные библиотеки
 -Flask: Веб-фреймворк
 -Pandas: Работа с данными
 -ReportLab: Генерация PDF
 -OpenPyXL: Работа с Excel
 -python-docx: Создание Word-документов
