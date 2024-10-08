//меняет фреймы в зависимости от выбранного режима
const calculatorType = document.getElementById('calculator-type');
const mode1 = document.getElementById('mode-1');
const mode2 = document.getElementById('mode-2');
const mode3 = document.getElementById('mode-3');

calculatorType.addEventListener('change', function() {
    mode1.style.display = 'none';
    mode2.style.display = 'none';
    mode3.style.display = 'none';

    if (this.value === '1') {
        mode1.style.display = 'block';
    } else if (this.value === '2') {
        mode2.style.display = 'block';
    } else if (this.value === '3') {
        mode3.style.display = 'block';
    }
});
