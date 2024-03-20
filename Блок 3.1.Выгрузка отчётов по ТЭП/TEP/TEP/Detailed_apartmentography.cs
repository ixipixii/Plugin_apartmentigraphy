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
            public string Floor { get; set; }
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
            String pathTemplate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Отчёты.xlsx");

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
                    if (room.LookupParameter("ADSK_Номер квартиры") != null &&
                       room.LookupParameter("ADSK_Номер квартиры").AsString() != null &&
                       room.LookupParameter("ADSK_Номер квартиры").AsString() != "" &&
                       room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                    {
                        if (!info.Any(i => i.Number == room.LookupParameter("ADSK_Номер квартиры").AsString()))
                        {
                            Info InfoElement = new Info(room.LookupParameter("ADSK_Количество комнат").AsInteger().ToString(),
                                                        room.LookupParameter("ADSK_Тип квартиры").AsString(),
                                                        room.LookupParameter("ADSK_Этаж").AsString(),
                                                        room.LookupParameter("ADSK_Номер секции").AsString(),
                                                        room.LookupParameter("ADSK_Номер квартиры").AsString());
                            info.Add(InfoElement);
                        }
                    }
                }
                else
                {
                    if (room.LookupParameter("ADSK_Номер квартиры") != null &&
                        room.LookupParameter("ADSK_Номер квартиры").AsString() != null &&
                        room.LookupParameter("ADSK_Номер квартиры").AsString() != "" &&
                        room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                    {
                        Info InfoElement = new Info(room.LookupParameter("ADSK_Количество комнат").AsInteger().ToString(),
                                                room.LookupParameter("ADSK_Тип квартиры").AsString(),
                                                room.LookupParameter("ADSK_Этаж").AsString(),
                                                room.LookupParameter("ADSK_Номер секции").AsString(),
                                                room.LookupParameter("ADSK_Номер квартиры").AsString());
                        info.Add(InfoElement);
                    }
                }

            }

            List<string> countSection = info.GroupBy(i => i.Section)
                .Select(group => group.Key)
                .ToList();

            List<string> countFloor = info.GroupBy(i => i.Floor)
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
                ExcelWorksheet worksheetTemplateHousing = packageTemplate.Workbook.Worksheets["Выгрузка на корпус"];
                ExcelWorksheet worksheetTemplateSection = packageTemplate.Workbook.Worksheets["Выгрузка на секцию"];

                //Открываем файл для записи
                using (var package = new ExcelPackage(new FileInfo(selectedFilePath)))
                {
                    // Удаляем и добавляем лист "Выгрузка на корпус"
                    package.Workbook.Worksheets.Delete("Выгрузка на корпус");
                    ExcelWorksheet worksheetHousing = package.Workbook.Worksheets.Add("Выгрузка на корпус");

                    // Копируем данные, форматирование и другие свойства из исходного листа в новый
                    for (int i = 1; i <= worksheetTemplateHousing.Dimension.Rows; i++)
                    {
                        for (int col = 1; col <= worksheetTemplateHousing.Dimension.Columns; col++)
                        {
                             worksheetTemplateHousing.Cells[i, col].Copy(worksheetHousing.Cells[i, col]);
                        }
                    }

                    int row = 21; //Начинаем заполнение с 21 строки
                    foreach (var section in countSection)
                    {
                        worksheetHousing.Cells["A" + row.ToString()].Value = "Секция номер " + section;

                        //Однокомнатные
                        if (info.Any(apart => apart.CountRoom == "1"))
                        {
                            worksheetHousing.Cells[column[0] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "S" && apart.Section == section);
                            worksheetHousing.Cells[column[1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "M" && apart.Section == section);
                        }

                        //Двухкомнатные
                        countApartCountRoom("2", section, null, worksheetHousing, column, row, info);

                        //3
                        countApartCountRoom("3", section, null, worksheetHousing, column, row, info);

                        //4
                        countApartCountRoom("4", section, null, worksheetHousing, column, row, info);

                        //5
                        countApartCountRoom("5", section, null, worksheetHousing, column, row, info);

                        //6
                        countApartCountRoom("6", section, null, worksheetHousing, column, row, info);

                        //Переносим строки с итогами
                        //Если больше данных нет, копирование не делаем
                        if (countSection[countSection.Count - 1] != section)
                        {
                            for (int col = 1; col <= worksheetHousing.Dimension.End.Column; col++)
                            {
                                worksheetHousing.Cells[row + 6, col].Copy(worksheetHousing.Cells[row + 7, col]);
                                worksheetHousing.Cells[row + 5, col].Copy(worksheetHousing.Cells[row + 6, col]);
                                worksheetHousing.Cells[row + 4, col].Copy(worksheetHousing.Cells[row + 5, col]);
                                worksheetHousing.Cells[row + 3, col].Copy(worksheetHousing.Cells[row + 4, col]);
                                worksheetHousing.Cells[row + 2, col].Copy(worksheetHousing.Cells[row + 3, col]);
                                worksheetHousing.Cells[row + 1, col].Copy(worksheetHousing.Cells[row + 2, col]);
                                worksheetHousing.Cells[row + 1, col].Clear();
                            }
                        }
                        row++;
                    }

                    // Удаляем и добавляем лист "Выгрузка на секцию"
                    package.Workbook.Worksheets.Delete("Выгрузка на секцию");
                    ExcelWorksheet worksheetSection = package.Workbook.Worksheets.Add("Выгрузка на секцию");

                    int countNameSection = 3;
                    row = 9;
                    foreach (var section in countSection)
                    {
                        worksheetSection.Cells["A" + countNameSection.ToString()].Value = "КОРПУС №_1_ СЕКЦИЯ №_" + section + "_";
                        
                        foreach (var floor in countFloor)
                        {
                            worksheetSection.Cells["A" + row].Value = "Этаж №" + floor;

                            //Однокомнатные
                            if (info.Any(apart => apart.CountRoom == "1"))
                            {
                                worksheetSection.Cells[column[0] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "S" && apart.Section == section && apart.Floor == floor);
                                worksheetSection.Cells[column[1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == "1" && apart.Range == "M" && apart.Section == section && apart.Floor == floor);
                            }

                            //Двухкомнатные
                            countApartCountRoom("2", section, floor, worksheetSection, column, row, info);

                            //3
                            countApartCountRoom("3", section, floor, worksheetSection, column, row, info);

                            //4
                            countApartCountRoom("4", section, floor, worksheetSection, column, row, info);

                            //5
                            countApartCountRoom("5", section, floor, worksheetSection, column, row, info);

                            //6
                            countApartCountRoom("6", section, floor, worksheetSection, column, row, info);

                            row++;
                        }

                        //Переносим строки с итогами
                        //Если больше данных нет, копирование не делаем
                        if (countSection[countSection.Count - 1] != section)
                        {
                            for (int col = 1; col <= worksheetHousing.Dimension.End.Column; col++)
                            {
                                worksheetSection.Cells[countNameSection + 5, col].Copy(worksheetSection.Cells[row + 8, col]);
                                worksheetSection.Cells[countNameSection + 4, col].Copy(worksheetSection.Cells[row + 7, col]);
                                worksheetSection.Cells[countNameSection + 3, col].Copy(worksheetSection.Cells[row + 6, col]);
                                worksheetSection.Cells[countNameSection + 2, col].Copy(worksheetSection.Cells[row + 5, col]);
                                worksheetSection.Cells[countNameSection + 1, col].Copy(worksheetSection.Cells[row + 4, col]);
                                worksheetSection.Cells[countNameSection, col].Copy(worksheetSection.Cells[row + 3, col]);
                            }
                        }

                        countNameSection = row + 3;
                        row = countNameSection + 6;
                    }

                    // Сохраняем изменения
                    package.Save();
                }
            }

            System.Diagnostics.Process.Start(selectedFilePath);

            return Result.Succeeded;
        }
        private void countApartCountRoom(string countRoom, string section, string floor, ExcelWorksheet worksheet, List<string> column, int row, List<Info> info)
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
            if(floor != null)
            {
                if (info.Any(apart => apart.CountRoom == countRoom))
                {
                    worksheet.Cells[column[col] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "S" && apart.Section == section && apart.Floor == floor);
                    worksheet.Cells[column[col + 1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "M" && apart.Section == section && apart.Floor == floor);
                    worksheet.Cells[column[col + 2] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "L" && apart.Section == section && apart.Floor == floor);
                }
            }
            else
            {
                if (info.Any(apart => apart.CountRoom == countRoom))
                {
                    worksheet.Cells[column[col] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "S" && apart.Section == section);
                    worksheet.Cells[column[col + 1] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "M" && apart.Section == section);
                    worksheet.Cells[column[col + 2] + row.ToString()].Value = info.Count(apart => apart.CountRoom == countRoom && apart.Range == "L" && apart.Section == section);
                }
            }
        }
    }
}
