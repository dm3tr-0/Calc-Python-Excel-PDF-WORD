
document.getElementById('calculate').addEventListener('click', function() {
    //функция запускается, когда нажали кнопочку рассчитать (index.html id = calculate)
    //по id из полей input(в index.html для режима1) берутся их значения

    const workHours = parseFloat(document.getElementById('work-hours').value);
    const daysCount = parseFloat(document.getElementById('days-count').value);
    const maxFilesDay = parseFloat(document.getElementById('max-files-day').value);

    //проверяем, что поля заполнены и их значение не nan
    const fields = ['work-hours', 'days-count', 'max-files-day'];
    const invalidField = fields.find(field => isNaN(parseFloat(document.getElementById(field).value)));
    if (invalidField) {
        alert('Пожалуйста, введите все данные корректно.');
        return;
    }
    //еще наверное надо добавить проверку, что значения не отрицательные и тд

    //теперь значения запихиваем в словарь(ключ : значение)
    const data = {
        work_hours: workHours,
        days_count: daysCount,
        max_files_day: maxFilesDay
    };
    //передаем этот словарик в main.py (def calculate())
    fetch('/calculate', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        //в json-формате
        body: JSON.stringify(data)
    })
    .then(response => response.json())
    .then(result => {
        //это мы получили ответ(из main.py def calculate()), он лежит в result
        //теперь запихиваем полученные значения в Итоговый результат(index.html id=result-ouput)
        document.getElementById('result-output').innerHTML = `
            <strong>Максимальное количество файлов в месяц:</strong> ${result.max_files_month.toFixed(2)}<br>
            <strong>Фактическое количество файлов в месяц:</strong> ${result.fact_files_month.toFixed(2)}<br>
            <strong>Разница в нагрузке:</strong> ${result.load_difference.toFixed(2)}
        `;
    })
    .catch(error => {
        console.error('Ошибка:', error);
        alert('Произошла ошибка при расчете.');
    });
});
