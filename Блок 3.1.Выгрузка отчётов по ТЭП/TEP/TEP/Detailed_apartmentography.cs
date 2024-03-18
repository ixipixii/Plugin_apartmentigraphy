using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TEP
{
    [Transaction(TransactionMode.Manual)]
    internal class Detailed_apartmentography : IExternalCommand
    {
        class Info
        {
            public string CountRoom { get; set; }
            public string Range { get; set; }
            string Floor { get; set; }
            public string Section { get; set; }
            public string Number { get; set; }
            public Info(string countRoom, string range, string floor, string section, string number) => (CountRoom, Range, Floor, Section, Number) = (countRoom, range, floor, section, number);
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //файл-шаблон с детальной квартирографией
            String pathTemplate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Детальная квартирография.xlsx");

            //Открываем диалог выбора сохранения отчёта
            var saveDialogImg = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "Отчёты.xlsx",
                DefaultExt = ".xlsx"
            };

            string selectedFilePath = string.Empty;

            if (saveDialogImg.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialogImg.FileName;
            }

            var rooms = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .WhereElementIsNotElementType()
                            .ToList();

            List<string> listParameter = new List<string>();
            List<string> listSections = new List<string>();
            List<string> listFloor = new List<string>();

            List<Info> info = new List<Info>();
            foreach (var room in rooms)
            {
                if (info.Count != 0)
                {
                    if (info.Any(i => i.Number != room.LookupParameter("ADSK_Номер квартиры").AsString()))
                    {
                        Info InfoElement = new Info(room.LookupParameter("ADSK_Количество комнат").AsString(),
                                                    room.LookupParameter("ADSK_Позиция отделки").AsString(),
                                                    room.LookupParameter("ADSK_Этаж").AsString(),
                                                    room.LookupParameter("ADSK_Номер секции").AsString(),
                                                    room.LookupParameter("ADSK_Номер квартиры").AsString());
                        info.Add(InfoElement);
                    }
                }
                else
                {
                    Info InfoElement = new Info(room.LookupParameter("ADSK_Количество комнат").AsString(),
                                                room.LookupParameter("ADSK_Позиция отделки").AsString(),
                                                room.LookupParameter("ADSK_Этаж").AsString(),
                                                room.LookupParameter("ADSK_Номер секции").AsString(),
                                                room.LookupParameter("ADSK_Номер квартиры").AsString());
                    info.Add(InfoElement);
                }

            }

            List<string> countSection = info.GroupBy(i => i.Section)
                .Select(group => group.Key)
                .ToList();

            List<string> column = new List<string>
            {
                "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"
            };

            //Открываем файл-шаблон
            using (var packageTemplate = new ExcelPackage(new FileInfo(pathTemplate)))
            {
                //Создаём и копируем информацию
                ExcelWorksheet worksheetTemplate = packageTemplate.Workbook.Worksheets["Выгрузка на корпус"];

                //Диапазон для чтения
                string rangeTemplate = "A2:S100";

                //Открываем файл для записи
                using (var package = new ExcelPackage(new FileInfo(selectedFilePath)))
                {
                    // Выбираем лист
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Выгрузка на корпус"];

                    int row = 21; //Начинаем заполнение с 21 строки
                    foreach(var section in countSection)
                    {
                        worksheet.Cells["А" + row.ToString()].Value = "Секция номер" + section; 

                        //Однокомнатные
                        if (info.Any(apart => apart.CountRoom == "1"))
                        {
                            worksheet.Cells[column[0] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "S");
                            worksheet.Cells[column[1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "M");
                        }

                        //Двухкомнатные
                        countApartCountRoom("2", worksheet, column, row, info);

                        //3
                        countApartCountRoom("3", worksheet, column, row, info);

                        //4
                        countApartCountRoom("4", worksheet, column, row, info);

                        //5
                        countApartCountRoom("5", worksheet, column, row, info);

                        //6
                        countApartCountRoom("6", worksheet, column, row, info);

                        row++;  
                    }

                    // Сохраняем изменения
                    package.Save();
                }
            }

            System.Diagnostics.Process.Start(selectedFilePath);

            return Result.Succeeded;
        }
        private void countApartCountRoom(string countRoom, ExcelWorksheet worksheet, List<string> column, int row, List<Info> info)
        {
            int col = 0;
            switch (countRoom)
            {
                case "2":
                    col = 2;
                    break;
                case "3":
                    col = 5;
                    break;
                case "4":
                    col = 8;
                    break;
                case "5":
                    col = 11;
                    break;
                case "6":
                    col = 14;
                    break;

            }
            if (info.Any(apart => apart.CountRoom == countRoom))
            {
                worksheet.Cells[column[col]     + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "S");
                worksheet.Cells[column[col + 1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "M");
                worksheet.Cells[column[col + 2] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "L");
            }
        }
    }
}
