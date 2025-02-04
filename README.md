# Для чего этот скрипт
Этот скрипт предназначен для автоматического сбора вопросов о товарах с платформы Ozon (через личный кабинет селлера) и сохранения их в Excel-файл.<br>
Пример готового файла:<br>
![image](https://github.com/user-attachments/assets/efe66e66-4c29-424f-947b-cfa4ecadb87c)


____
# Начало работы

* [Установка зависимостей](#установка-зависимостей)  
* [Настройка параметров](#настройка-параметров)  
* [Запуск скрипта](#запуск-скрипта)  
* [Как работает скрипт](#как-работает-скрипт)  
* [Структура выходного файла](#структура-выходного-файла)  
* [Обработка ошибок](#обработка-ошибок)  
* [Пример использования](#пример-использования)  
* [Важные замечания](#важные-замечания)  




## Установка зависимостей
Перед запуском скрипта убедитесь, что у вас установлены все необходимые библиотеки. Для этого выполните команду:
```
pip install pandas requests openpyxl
```
## Настройка параметров
В начале скрипта находятся переменные, которые необходимо настроить перед запуском:<br>
`company_id`: Уникальный идентификатор компании на Ozon. Его можно найти в URL страницы компании на Ozon. Например, для URL https://www.ozon.ru/seller/ooo-mebelnaya-fabrika-volzhanka-1234/products/?miniapp=seller_1234 идентификатор компании — 1234.<br>
Или в настройках ЛК селлера, раздел ["Информация о компании"](https://seller.ozon.ru/app/settings/info)<br>
![image](https://github.com/user-attachments/assets/c2d8d32f-fd5c-43b7-8041-4424fd1fd6b4)<br>
`date_to`: Дата, до которой скрипт будет собирать вопросы. Формат даты: "YYYY-MM-DD". Рекомендуется указывать дату на один день раньше нужной (нужны вопросы до 21.01 включительно - указываем 20.01).<br>
`cookie`: Актуальные куки для авторизации на Ozon. Их можно получить через браузер, авторизовавшись на сайте Ozon.<br>
`cases`: Количество запросов, которые скрипт выполнит для сбора вопросов. Один запрос возвращает 10 вопросов.


## Запуск скрипта
После настройки параметров запустите скрипт:
```
python main.py
```

## Как работает скрипт
Скрипт отправляет запросы к (не публичному) API Ozon для получения вопросов.<br>
Каждый запрос возвращает 10 вопросов.<br>
Данные сохраняются в Excel-файл с именем `ozon_questions_YYYY-MM-DD.xlsx`, где YYYY-MM-DD — текущая дата.<br>
Если файл уже существует, новые данные добавляются в конец файла.<br>
Скрипт останавливается, если достигает указанной даты (`date_to`) или выполняет заданное количество запросов (`cases`).

## Структура выходного файла
Выходной Excel-файл содержит следующие столбцы:
* id: Уникальный идентификатор вопроса.
* Артикул OZON: Артикул товара на Ozon.
* Текст вопроса: Текст вопроса, заданного покупателем.
* Дата и время публикации вопроса: Дата и время, когда вопрос был опубликован.
* Имя автора: Имя пользователя, задавшего вопрос.
* Торговая сеть: Название бренда или торговой сети.
* Ответов: Количество ответов на вопрос.
* Название товара: Название товара, к которому относится вопрос.
* Ссылка на товар: Ссылка на страницу товара.
* Артикул товара: Артикул товара.

## Обработка ошибок
Если скрипт сталкивается с ошибкой (например, проблемы с сетью или API), он делает до 60 повторных попыток с задержкой между ними.
В случае успешного завершения скрипт выводит сообщение "Отчет готов!".

## Пример использования
Установите зависимости.<br>
Настройте параметры (`company_id`, `date_to`, `cookie`, `cases`).<br>
Запустите скрипт.<br>
После завершения работы скрипта откройте файл `ozon_questions_YYYY-MM-DD.xlsx` для анализа данных.

## Важные замечания
Убедитесь, что куки актуальны. Если куки истекли, скрипт не сможет получить данные.<br>
Не указывайте слишком большое значение для `cases`, чтобы избежать блокировки со стороны Ozon.<br>
Скрипт делает паузы между запросами, чтобы снизить нагрузку на сервер Ozon.



