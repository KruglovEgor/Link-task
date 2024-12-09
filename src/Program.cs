using System;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordDocumentGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Определяем путь к корневой папке
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string rootDirectory = Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\.."));

                // Указываем путь к папке out
                string outDirectory = Path.Combine(rootDirectory, "out");

                // Проверяем наличие папки и создаём её, если она отсутствует
                if (!Directory.Exists(outDirectory))
                {
                    Directory.CreateDirectory(outDirectory);
                }

                // Пути к файлам
                string configPath = Path.Combine(rootDirectory, @"resources\config.json"); // Файл конфигурации
                string counterPath = Path.Combine(rootDirectory, @"resources\counter.txt"); // Файл счётчика

                // Проверка доступности конфигурационного файла
                if (!File.Exists(configPath))
                {
                    throw new FileNotFoundException($"Конфигурационный файл '{configPath}' не найден.");
                }

                // Чтение конфигурационного файла
                Config config;
                try
                {
                    config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(configPath));
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Ошибка при чтении конфигурационного файла: " + ex.Message);
                }

                // Проверяем наличие CSV-файла
                if (!File.Exists(config.CsvFilePath))
                {
                    throw new FileNotFoundException($"CSV-файл '{config.CsvFilePath}' не найден.");
                }

                // Чтение данных из CSV
                var csvLines = File.ReadAllLines(config.CsvFilePath);
                if (csvLines.Length == 0)
                {
                    throw new InvalidOperationException("CSV-файл пуст.");
                }

                // Получение заголовков и данных
                var headers = csvLines[0].Split(',');
                var rows = csvLines.Skip(1).Select(line => line.Split(',')).ToList();

                // Генерация номера документа
                int documentNumber = GetAndUpdateDocumentNumber(counterPath);

                // Формируем имя файла из DocumentTitle
                string resultPath = Path.Combine(outDirectory, $"Документ №{documentNumber}.docx");

                // Создание документа
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(resultPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    Body body = mainPart.Document.Body;

                    // Добавление заголовка
                    AddTitle(body, config.DocumentTitle);

                    // Добавление таблицы
                    AddTable(body, headers, rows);

                    // Добавление трех пустых строк
                    AddEmptyLines(body, 3);

                    // Добавление поля для подписи
                    AddSignature(body, config.Employee);

                    // Сохранение документа
                    mainPart.Document.Save();
                }

                Console.WriteLine($"Документ успешно создан! Сохранен по пути: {resultPath}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine("Ошибка в конфигурационном файле: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла непредвиденная ошибка: " + ex.Message);
            }
        }

        // Функция для чтения, использования и обновления номера документа
        static int GetAndUpdateDocumentNumber(string counterPath)
        {
            int currentNumber = 1; // Начинаем с 1, если файл отсутствует

            // Если файл существует, читаем текущий номер
            if (File.Exists(counterPath))
            {
                string content = File.ReadAllText(counterPath);
                if (int.TryParse(content, out int parsedNumber))
                {
                    currentNumber = parsedNumber;
                }
            }

            // Записываем следующий номер обратно в файл
            File.WriteAllText(counterPath, (currentNumber + 1).ToString());

            return currentNumber; // Возвращаем текущий номер
        }
        
        // Функция добавления названия документа
        static void AddTitle(Body body, string title)
        {
            Paragraph paragraph = new Paragraph(
                new ParagraphProperties
                {
                    Justification = new Justification { Val = JustificationValues.Center } // Выравнивание по центру
                },
                new Run(
                    new RunProperties
                    {
                        Bold = new Bold(),
                        FontSize = new FontSize { Val = "36" } // Размер шрифта (36 = 18pt)
                    },
                    new Text(title)
                )
            );
            body.Append(paragraph);
        }

        // Функция добавления подписи для сотрудника employee
        static void AddSignature(Body body, Employee employee)
        {
            string signatureText = $"Обучение провел ";
            string boldText = $"{employee.Position} {employee.LastName} {employee.FirstName[0]}.{employee.MiddleName[0]}. ";

            Paragraph paragraph = new Paragraph(
                new ParagraphProperties
                {
                    Justification = new Justification { Val = JustificationValues.Center } // Выравнивание по центру
                },
                new Run(
                    new Text(signatureText)
                    {
                        Space = SpaceProcessingModeValues.Preserve // Сохраняем пробел после текста
                    }
                ),
                new Run(
                    new RunProperties { Bold = new Bold() },
                    new Text(boldText)
                    {
                        Space = SpaceProcessingModeValues.Preserve // Сохраняем пробел после текста, если нужно
                    }
                ),
                new Run(new Text(" ___________ (подпись)"))
            );

            body.Append(paragraph);
        }

        // Функция создания таблицы. На вход поступает body для создания таблицы в нем, headers - заголовки столбцов, rows - данные
        static void AddTable(Body body, string[] headers, System.Collections.Generic.List<string[]> rows)
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct } // Ширина таблицы
            );
            table.AppendChild(tableProperties);

            // Добавление строки с заголовком "Протокол"
            TableRow titleRow = new TableRow();
            TableCell titleCell = new TableCell(
                new TableCellProperties
                {
                    TableCellBorders = new TableCellBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 6 },
                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
                        new RightBorder { Val = BorderValues.Single, Size = 6 }
                    ),
                    GridSpan = new GridSpan { Val = headers.Length + 1 } // Объединяем все колонки
                },
                new Paragraph(
                    new ParagraphProperties
                    {
                        Justification = new Justification { Val = JustificationValues.Center }
                    },
                    new Run(
                        new RunProperties { Bold = new Bold() },
                        new Text("Протокол")
                    )
                )
            );
            titleRow.Append(titleCell);
            table.Append(titleRow);

            // Добавление заголовков таблицы (включая колонку "№")
            TableRow headerRow = new TableRow();
            headerRow.Append(CreateCell("№", isHeader: true)); // Колонка "№"

            foreach (var header in headers)
            {
                headerRow.Append(CreateCell(header, isHeader: true));
            }
            table.Append(headerRow);

            // Добавление строк данных
            int rowIndex = 1;
            foreach (var row in rows)
            {
                TableRow dataRow = new TableRow();
                dataRow.Append(CreateCell(rowIndex.ToString())); // Номер строки
                foreach (var cellData in row)
                {
                    dataRow.Append(CreateCell(cellData));
                }
                table.Append(dataRow);
                rowIndex++;
            }

            body.Append(table);
        }

        // Функция создания ячейки таблицы с текстом text
        static TableCell CreateCell(string text, bool isHeader = false)
        {
            TableCell cell = new TableCell(
                // Задаем видимые границы ячеек
                new TableCellProperties
                {
                    TableCellBorders = new TableCellBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 6 },
                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
                        new RightBorder { Val = BorderValues.Single, Size = 6 }
                    )
                },
                // Вставляем текст в ячейку
                new Paragraph(
                    new ParagraphProperties
                    {                     
                        Justification = new Justification { Val = JustificationValues.Center } // Выравнивание по центру 
                    },
                    new Run(
                        new RunProperties { Bold = isHeader ? new Bold() : null }, // Делаем текст с флагом isHeader жирным
                        new Text(text)
                    )
                )
            );
            return cell;
        }

        // Функция добавления count пустых строчек в body
        static void AddEmptyLines(Body body, int count)
        {
            for (int i = 0; i < count; i++)
            {
                body.Append(new Paragraph(new Run(new Text(" "))));
            }
        }
    }

    // Класс, описывающий конфиг (config.json)
    public class Config
    {
        public Employee Employee { get; set; }
        public string DocumentTitle { get; set; }
        public string CsvFilePath { get; set; }
    }

    // Класс, описывающий сотрудника
    public class Employee
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string Position { get; set; }
    }
}
