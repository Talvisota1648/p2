1) Установи зависимости
pip install requests prettytable

2) Использование
# Только вывод в консоль
python script.py --queue MYQUEUE

# Сохранить в CSV
python script.py --queue MYQUEUE --csv

# Сохранить в Excel
python script.py --queue MYQUEUE --xlsx

# Сохранить в оба формата
python script.py --queue MYQUEUE --csv --xlsx
