using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using XLSX_toXML_toDoc.Model;

namespace XLSX_toXML_toDoc.ViewModel
{
    class MainViewModel
    {
        private string? excelFilePath;
        private static string xmlFilePath;
        private string? reportString;
        public static bool IsFileImported { get; private set; }
        public static bool IsReportDone { get; private set; }
        public static bool IsReportSaved { get; private set; }
        private List<XlsxPerson> xlsxPersons;

        public List<XlsxPerson> XlsxPersons
        {
            get { return xlsxPersons; }
            set { xlsxPersons = value; }
        }

        

        public MainViewModel()
        {
            XlsxPersons = new List<XlsxPerson>();
        }
        // функции генераторы
        public static string? ChooseExcelFile()
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "XLSX Source Files | *.xlsx",
                Title = "Pick an xlsx source file"
            };

            IsFileImported = fileDialog.ShowDialog() == true;
            return IsFileImported ? fileDialog.FileName : null;
        }
        public void ImportDataFromExcel(string filePath)
        {
            XSSFWorkbook workbook;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }

            ISheet worksheet = workbook.GetSheetAt(0);

            for (int row = 1; row <= worksheet.LastRowNum; row++)
            {
                IRow excelRow = worksheet.GetRow(row);
                if (excelRow != null)
                {
                    XlsxPerson person = new XlsxPerson
                    {
                        Surname = excelRow.GetCell(0).StringCellValue,
                        Name = excelRow.GetCell(1).StringCellValue,
                        Gender = excelRow.GetCell(2).StringCellValue,
                        Age = Convert.ToInt32(excelRow.GetCell(3).NumericCellValue),
                        AccountStatus = excelRow.GetCell(4).StringCellValue,
                        Salary = Convert.ToInt32(excelRow.GetCell(5).NumericCellValue)
                    };
                    XlsxPersons.Add(person);
                }
            }
        }

        private XDocument CreateXmlDocument()
        {
            XElement generatedXml = new XElement("GeneratedXml");
            
            foreach (var person in XlsxPersons)
            {
                XElement personElement = new XElement("Person",
                    new XElement("Surname", person.Surname),
                    new XElement("Name", person.Name),
                    new XElement("Gender", person.Gender),
                    new XElement("Age", person.Age),
                    new XElement("AccountStatus", person.AccountStatus),
                    new XElement("Salary", person.Salary)
                );

                generatedXml.Add(personElement);
            }
            XlsxPersons.Clear();
            return new XDocument(generatedXml);
        }

        private static string? SaveXmlDocument(XDocument xmlDocument, string path)
        {
            string directoryPath = Path.GetDirectoryName(path); 
            xmlFilePath = Path.Combine(directoryPath, "GeneratedXML.xml");
            xmlDocument.Save(xmlFilePath);
            return xmlFilePath;
        }


        public static string GenerateReportString(XDocument xmlDocument)
        {
            var persons = xmlDocument.Root.Elements("Person");

            int totalMen = persons.Count(p => string.Equals(p.Element("Gender")?.Value.Trim(), "м", StringComparison.OrdinalIgnoreCase));
            int totalWomen = persons.Count(p => string.Equals(p.Element("Gender")?.Value.Trim(), "ж", StringComparison.OrdinalIgnoreCase));

            int menAge30_40 = persons.Count(p => string.Equals(p.Element("Gender")?.Value.Trim(), "м", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(p.Element("Age")?.Value.Trim(), out int age) && age >= 30 && age <= 40);

            int standardAccounts = persons.Count(p => string.Equals(p.Element("AccountStatus")?.Value.Trim(), "стандарт", StringComparison.OrdinalIgnoreCase));
            int premiumAccounts = persons.Count(p => string.Equals(p.Element("AccountStatus")?.Value.Trim(), "премиум", StringComparison.OrdinalIgnoreCase));

            int womenPremiumAccountsUnder30 = persons.Count(p => string.Equals(p.Element("Gender")?.Value.Trim(), "ж", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(p.Element("Age")?.Value.Trim(), out int age) && age < 30 && string.Equals(p.Element("AccountStatus")?.Value.Trim(), "премиум", StringComparison.OrdinalIgnoreCase));

            int womenWithMaxSalaryAge23_35 = persons.Where(p => string.Equals(p.Element("Gender")?.Value.Trim(), "ж", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(p.Element("Age")?.Value.Trim(), out int age) && age >= 23 && age <= 35)
                .GroupBy(p => Convert.ToInt32(p.Element("Age")?.Value.Trim()))
                .Select(g => g.OrderByDescending(p => Convert.ToInt32(p.Element("Salary")?.Value.Trim())).First())
                .Count();

            var menBySalary = persons.Where(p => string.Equals(p.Element("Gender")?.Value.Trim(), "м", StringComparison.OrdinalIgnoreCase))
                .OrderBy(p => int.TryParse(p.Element("Salary")?.Value.Trim(), out int salary) ? salary : int.MaxValue)
                .Take(3);

            var womenBySalary = persons.Where(p => string.Equals(p.Element("Gender")?.Value.Trim(), "ж", StringComparison.OrdinalIgnoreCase))
                .OrderBy(p => int.TryParse(p.Element("Salary")?.Value.Trim(), out int salary) ? salary : int.MaxValue)
                .Take(3);

            StringBuilder reportString = new StringBuilder();
            reportString.AppendLine("Отчет:")
            .AppendLine($"1. Всего мужчин: {totalMen}, Всего женщин: {totalWomen}")
            .AppendLine($"2. Мужчин в возрасте 30-40 лет: {menAge30_40}")
            .AppendLine($"3. Стандартных аккаунтов: {standardAccounts}, Премиум-аккаунтов: {premiumAccounts}")
            .AppendLine($"4. Женщин с премиум-аккаунтом до 30 лет: {womenPremiumAccountsUnder30}")
            .AppendLine($"5. Женщин с максимальным окладом в возрасте от 23 до 35 лет: {womenWithMaxSalaryAge23_35}")
            .AppendLine("6.1. Мужчины с наименьшей зарплатой:");

            foreach (var man in menBySalary)
            {
                reportString.AppendLine($"{man.Element("Name")?.Value.Trim()} {man.Element("Surname")?.Value.Trim()}, Зарплата: {man.Element("Salary")?.Value.Trim()}");
            }

            reportString.AppendLine("6.2. Женщины с наименьшей зарплатой: ");
            foreach (var woman in womenBySalary)
            {
                reportString.AppendLine($"{woman.Element("Name")?.Value.Trim()} {woman.Element("Surname")?.Value.Trim()}, Зарплата: {woman.Element("Salary")?.Value.Trim()}");
            }
            

            return reportString.ToString();
        }

        private XWPFDocument GenerateDoc()
        {
            XWPFDocument document = new XWPFDocument();
            XWPFParagraph paragraph = document.CreateParagraph();
            string[] lines = reportString.Split('\n');

            foreach (string line in lines)
            {
                XWPFRun run = paragraph.CreateRun();
                run.SetText(line.Trim());
                run.AddBreak(BreakType.TEXTWRAPPING); // перенос строки
            }

            return document;
        }

        // сборные функции
        public void ImportAndSetDirectory()
        {
            excelFilePath = ChooseExcelFile();
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                ImportDataFromExcel(excelFilePath);
            }
        }
        public void GenerateReport()
        {
            XDocument xmlDocument = CreateXmlDocument();
            SaveXmlDocument(xmlDocument, excelFilePath);
            reportString = GenerateReportString(xmlDocument);
            IsReportDone = true;
        }
        public void GenerateAndSaveDoc()
        {
            string fileName = "generatedDOCX.docx";
            string directoryPath = Path.GetDirectoryName(excelFilePath);

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = fileName,
                InitialDirectory = directoryPath
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                XWPFDocument document = GenerateDoc();

                using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    document.Write(stream);
                }
                DeleteFile(xmlFilePath);
                IsReportSaved = true;
                excelFilePath = null;
                reportString = null;
            }
        }
        public void DeleteFile(string filePath)
        {
                File.Delete(filePath);
        }

    }
}

