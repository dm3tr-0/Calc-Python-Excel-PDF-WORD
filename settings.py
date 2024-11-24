###################
#зашитые значения:
###################

#время обработки(мин)
time_obr_day = 8 #день
time_obr_night = 6 #ночь

#кол-во минут в часе
min_in_hour = 50
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
    '180h': work_hours['180h'] * min_in_hour / time_obr_day * 0.85, #58.43
    '168h': work_hours['168h'] * min_in_hour / time_obr_day * 0.85, #42.5
    '79h': work_hours['79h'] * min_in_hour / time_obr_day * 0.85, #21.25
    '180h_pr': work_hours['180h_pr'] * min_in_hour / time_obr_night * 0.85, #77.92
    '180h_night': work_hours['180h_night'] * min_in_hour / time_obr_night * 0.85 #77.92
}
# максимальное колвичество файлов в месяц для 180ч, 168ч, 79ч, 180ч пр/вых, 180ч ночь соответсвтенно
max_files_month = {
    '180h': mid_days_count['180h'] * max_files_day['180h'], #876.56
    '168h': mid_days_count['168h'] * max_files_day['168h'], #850.0
    '79h': mid_days_count['79h'] * max_files_day['79h'], #425.0
    '180h_pr': mid_days_count['180h_pr'] * max_files_day['180h_pr'], #1168.75
    '180h_night': mid_days_count['180h_night'] * max_files_day['180h_night'] #1168.75
}
