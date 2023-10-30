<div align="Center"><h1>Project Calendar</h1></div>

Проект сделан совместно с [Jadedboat780](https://github.com/Jadedboat780/) и [CaptainSliva](https://github.com/CaptainSliva).

Это проект для нашего колледжа [ГБПОУ РКСИ](https://www.rksi.ru/). Данная программа собирает данные о расписании
с сайта колледжа [ссылка](https://www.rksi.ru/mobile_schedule) и из таблицы Google Sheets [ссылка](https://drive.google.com/drive/folders/19yyXXullGGMIT3XISiZ33wkDxHJy0zvb). После чего она заносит эти данные в Google
Calendar [ссылка](https://calendar.google.com/)

## Понятия

---

Что такое `планшетка`? - это Google Sheets документ с актуальным расписанием на день. этот файл постоянно обновляется в течении дня учебным отделом колледжа ГБПОУ РКСИ

## Как работает программа?

---

С понедельника по субботу включительно в определенное время программа запускает скачивание
планшетки за текущий день
<image src="DocumentationImage/img.png" alt="Структура планшетки">
после чего проверяет ее заполненность. Затем
происходит поиск и обновление эвентов за текущий день. Потом данные из Google Sheets
таблицы заносятся в Google Calendar в виде новых эвентов. Запуск этих процессов происходит
в конце каждой парой.

В воскресенье запускается процесс поиска и удаления эвентов за будущую неделю. 
<image src="DocumentationImage/img_1.png" alt="Структура сайта">
Затем происходит
парсинг расписания сайта колледжа для групп и преподавателей, а после это расписание обрабатывается
и заносится в Google Calendar.

<image src="DocumentationImage/img_2.png" alt="Google Calendar">

Оформление эвента - пары
<image src="DocumentationImage/img_4.png" alt="Events">

Пример расписания на мобильных устройствах
<image src="DocumentationImage/img_5.jpg" alt="Events">

## Структура

---

Основа проекта разделена на два ключевых файла: `main.py` и `Schedule_Officws.py`. Это 
сделано для ускорения и создание эвентов 

## Дополнительно

---

Также для этого проекта был разработан Telegram-бот `@RKSI_Calendar_bot`, находится в папке 
под названием "RKSI_bot", который выдает
ссылку на Google Calendar с расписанием преподавателям и студентам колледжа ГБПОУ РКСИ
<image src="DocumentationImage/img_3.png" alt="Google Calendar">

## Файлы

---

- `main.py` - файл м программой календаря, который создает расписание для групп и преподавателей
- `id.py` - файд с id-шниками календарей групп и преподавателей
- `Schedule_Offices.py` - файл, с программой календаря, который создает заполненность кабинетов
- `id_offices.py` - файл с id-шниками кабинетов
- RKSI_bot - папка с файлами Telegram-бота
  - `Aiogram_bot.py` - основной файл бота
  - `id.py` - файл с общедоступными id-шниками календарей для групп и преподавателей
  - `TEST_VIDEO.mp4` - видео с инструкцией как добавить расписание к себе в Google Calendar
- `TEST.json` - файл с авторизационными данными для работы с api Google Calendar (данный файл добавлен в `.gitignore` для конфидициальности)

## Установка

---

1. Клонируйте репозиторий с github
2. Создайте виртуальное окружение 
3. Установите зависимости
4. Запустите программу командой `python3 main.py`

# Ссылки

---

1. Официальный сайт колледжа [ГБПОУ РКСИ](https://www.rksi.ru/)
2. Папка, где хранятся [планшетки](https://drive.google.com/drive/folders/19yyXXullGGMIT3XISiZ33wkDxHJy0zvb)
3. Сайт с расписанием занятий на [неделю](https://www.rksi.ru/mobile_schedule)
4. Сайт [Google Calendar](https://calendar.google.com/)
5. Команда: [Jadedboat780](https://github.com/Jadedboat780/), [CaptainSliva](https://github.com/CaptainSliva)