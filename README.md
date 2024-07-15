## Автоматизирует отчёт по тратам из мобильного приложения

для работы необходимо:
* загрузить файл с тратами
* переименовать его в формате **месяц_год_имя.xlsx**
* если были переводы между собой в течение месяца,
то необходимо их прописать внутри файла вручную (но только в одном!)

вызов скрипта:
> python -m app --file C:\Users\iii\Desktop\budget\июнь_2024_iii.xlsx

если такой месяц уже есть в файле отчёта, данные добавятся к нему,
если нет - создаётся новый месяц

