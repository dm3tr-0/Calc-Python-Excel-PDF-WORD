document.getElementById('mode-1').addEventListener('submit', function(event) {
    event.preventDefault();

    //собираем данные формы
    const formData = new FormData(this);

    //отправляем запрос на сервер через fetch
    fetch('/calculate', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        //выводим результат в блоке result-output
        document.getElementById('result-output').innerHTML = data.result;
    })
    .catch(error => {
        console.error('Ошибка:', error);
        alert('Произошла ошибка при расчете.');
    });
});
