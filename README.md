# Word Document Generator

## Описание

Этот проект представляет собой генератор документов Word на C#. Программа:
1. Создает новый документ на основе шаблона
2. Строит таблицу с данными из CSV.
3. Генерирует номер документа с помощью счетчика и указывает его в названии
4. Добавляет поля для подписи сотрудника на основе данных из конфигурационного файла
5. Сохранение документа в папке out

## Файлы

- **`src/Program.cs`**: Основной код программы.
- **`resources/config.json`**: Конфигурационный файл с настройками. **Обязателен к редактированию перед использованием!**
- **`resources/counter.txt`**: Файл счётчика номера документа. Создаётся автоматически.

## Установка и запуск

### Требования

- .NET SDK 7.0 или выше.

### Инструкции по запуску

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/KruglovEgor/Link-task.git
   cd Link-task
   ```
2. Отредактируйте **`resources/config.json`** под себя!
3. Откройте проект в Visual Studio или выполните сборку из командной строки:
	```bash
	dotnet build
	```
4. Запустите проект:
	```bash
	dotnet run
	```
