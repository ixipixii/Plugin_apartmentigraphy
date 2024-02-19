using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using MathNet.Numerics;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TEP
{
    internal abstract class Unloading
    {
        public Unloading() { }
        public string GetExeDirectory() //Получить путь до DLL
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);
            return path;
        }
        public string CopyFile(string nameFile) //Копировать файл-шаблон
        {
            //Путь до DLL
            string absPath = GetExeDirectory();
            //Создаём файл - шаблон
            String pathFile = Path.Combine(absPath, nameFile + ".xlsx");
            pathFile = Path.GetFullPath(pathFile);

            FileInfo fileInf = new FileInfo(pathFile);

            //Открываем диалог выбора сохранения отчёта
            var saveDialogImg = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = nameFile + ".xlsx",
                DefaultExt = ".xlsx"
            };

            string selectedFilePath = string.Empty;

            if (saveDialogImg.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialogImg.FileName;
            }

            //Копируем в выбранную папку
            fileInf.CopyTo(selectedFilePath);
            //Возвращаем путь
            return selectedFilePath;
        }
        public void FillCell(int rowIndex, int columnIndex, String cellValue, String excelPath) //Заполнить ячейку
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Отчёт"];

                // Координаты ячейки, которую вы хотите изменить
                int rowNumber = rowIndex;
                int columnNumber = columnIndex;

                // Новое значение для ячейки
                string newValue = cellValue;

                // Устанавливаем значение ячейки
                worksheet.Cells[rowNumber, columnNumber].Value = newValue;

                // Сохраняем изменения
                package.Save();
            }
        }
        public int FillCellParameter(int rowIndex, int columnIndex, String cellValue, String excelPath, string value_1, string value_3) //Заполнить ячейку
        {
            //Пустые данные не заносим
            if(cellValue == "0")
                return rowIndex;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Отчёт"];

                // Координаты ячейки, которую вы хотите изменить
                int rowNumber = rowIndex;
                int columnNumber = columnIndex;

                // Новое значение для ячейки
                string newValue = cellValue;

                // Устанавливаем значение ячейки
                worksheet.Cells[rowNumber, columnNumber].Value = newValue;
                worksheet.Cells[rowNumber, 1].Value = value_1;
                worksheet.Cells[rowNumber, 3].Value = value_3;

                // Сохраняем изменения
                package.Save();

                rowIndex++;
            }
            return rowIndex;
        }

        //Методы Elements возвращают лист определённых элементов
        //Методы Area возвращают площади элементов
        //Метод Values возвращает лист значений определённых параметров определённых элементов

        public virtual List<Element> Elements(BuiltInCategory category, Document doc)
        {
            List<Element> elements = new FilteredElementCollector(doc)
                                                    .OfCategory(category)
                                                    .WhereElementIsNotElementType()
                                                    .ToList();

            return elements;
        }
        public virtual List<Element> Elements(BuiltInCategory category, Document doc, String filterParameter, String valueFilterParameter)
        {
            List<Element> elements = new FilteredElementCollector(doc)
                                                    .OfCategory(category)
                                                    .WhereElementIsNotElementType()
                                                    .Where(e => e != null 
                                                        && e.LookupParameter(filterParameter) != null 
                                                        && e.LookupParameter(filterParameter).AsString() != null
                                                        && e.LookupParameter(filterParameter).AsString() != "")
                                                    .Where(e => e.LookupParameter(filterParameter).AsString().Contains(valueFilterParameter))
                                                    .ToList();

            return elements;
        }
        public virtual List<Element> Elements(BuiltInCategory category, Document doc, String filterParameter_1, String valueFilterParameter_1,
                                                                                      String filterParameter_2, String valueFilterParameter_2)
        {
            List<Element> elements = new FilteredElementCollector(doc)
                                                    .OfCategory(category)
                                                    .WhereElementIsNotElementType()
                                                    .Where(e => e != null
                                                        && e.LookupParameter(filterParameter_1) != null && e.LookupParameter(filterParameter_2) != null 
                                                        && e.LookupParameter(filterParameter_2).AsString() != null && e.LookupParameter(filterParameter_1).AsString() != null
                                                        && e.LookupParameter(filterParameter_2).AsString() != "" && e.LookupParameter(filterParameter_1).AsString() != "")
                                                    .Where(e => e.LookupParameter(filterParameter_1).AsString().Contains(valueFilterParameter_1) &&
                                                                e.LookupParameter(filterParameter_2).AsString() == valueFilterParameter_2)
                                                    .ToList();

            return elements;
        }
        public virtual List<Element> Elements(BuiltInCategory category, Document doc, String filterParameter, int valueFilterParameter, bool valeuCompare)
        {
            //Если знак >= то valeuCompare = 1. Если знак < то valeuCompare = 0.
            if(valeuCompare)
            {
                List<Element> elements = new FilteredElementCollector(doc)
                                        .OfCategory(category)
                                        .WhereElementIsNotElementType()
                                        .Where(e => int.Parse(e.LookupParameter(filterParameter).AsString()) >= (valueFilterParameter))
                                        .ToList();
                return elements;
            }
            else
            {
                List<Element> elements = new FilteredElementCollector(doc)
                                        .OfCategory(category)
                                        .WhereElementIsNotElementType()
                                        .Where(e => int.Parse(e.LookupParameter(filterParameter).AsString()) < (valueFilterParameter))
                                        .ToList();
                return elements;
            }
        }
        public virtual List<Element> ElementsEquals(BuiltInCategory category, Document doc, String filterParameter, String valueFilterParameter)
        {
            List<Element> elements = new FilteredElementCollector(doc)
                                                    .OfCategory(category)
                                                    .WhereElementIsNotElementType()
                                                    .Where(e => e.LookupParameter(filterParameter).AsString() == valueFilterParameter)
                                                    .ToList();

            return elements;
        }
        public virtual double Areas(List<Element> elements, Document doc)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                    area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
            }
            area = area.Round(3);
            return area;
        }
        public virtual double Areas(List<Element> elements, Document doc, String PARAMETER)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                try
                {
                    if (element.LookupParameter("PNR_Имя помещения").AsString() != null && element.LookupParameter("PNR_Имя помещения").AsString().Contains(PARAMETER))
                    {
                        area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }
        public virtual double AreasHeigher(List<Element> elements, Document doc, string height)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                try
                {
                    if (element.LookupParameter("PNR_Имя помещения").AsString() != null && int.Parse(element.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()) > int.Parse(height))
                    {
                        area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }
        public virtual double AreasLower(List<Element> elements, Document doc, string height)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                try
                {
                    if (element.LookupParameter("PNR_Имя помещения").AsString() != null && int.Parse(element.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()) < int.Parse(height))
                    {
                        area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }
        public virtual List<String> Values(String parameter, List<Element> elements)
        {
            List<String> values = new List<String>();           
            foreach (Element element in elements)
            {
                Room room = element as Room;
                values.Add(room.LookupParameter(parameter).AsString());
            }
            return values;
        }

        public int FillCellsArea(ExcelWorksheet worksheet1, int rowFunc, List<String> roomName, List<Element> roomList)
        {
            int rowName = rowFunc + 1;
            double areaName = 0.0;
            double areaFunc = 0.0;
            int j = 0;
            String nameFunc = roomList[0].LookupParameter("PNR_Функция помещения").AsString();
            //Проходимся по всем именам найденных помещений
            for (int i = 0; i < roomName.Count(); i++)
            {
                //Сравниваем помещения на нужное имя
                for (j = 0; j < roomList.Count(); j++)
                {
                    String nameFuncNew = roomList[j].LookupParameter("PNR_Функция помещения").AsString();
                    //Если у помещения нужное имя, суммируем его площадь
                    if (roomList[j].LookupParameter("PNR_Имя помещения").AsString() == roomName[i])
                    {
                        areaName += UnitUtils.ConvertFromInternalUnits(roomList[j].get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                        double areaOneRoom = UnitUtils.ConvertFromInternalUnits(roomList[j].get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                        //Если функция у помещений не меняется, суммируем её
                        if (nameFuncNew == nameFunc)
                        {
                            areaFunc += areaOneRoom;
                        }
                        //Если функция поменялась, заносим в ячейку сведения о функции
                        else
                        {
                            areaFunc= areaFunc.Round(3);
                            areaName= areaName.Round(3);
                            worksheet1.Cells[rowFunc, 1].Value = nameFunc; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                            nameFunc = nameFuncNew;
                            areaFunc = areaOneRoom;
                            rowFunc = rowName;
                            rowName++;
                        }
                    }
                }
                //После того, как прошлись по всем помещениям и сравнили их на имя заносим в строку сведения о площади помещения
                areaFunc = areaFunc.Round(3);
                areaName = areaName.Round(3);
                worksheet1.Cells[rowName, 1].Value = roomName[i]; worksheet1.Cells[rowName, 1].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 2].Value = areaName; worksheet1.Cells[rowName, 2].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 3].Value = "кв.м"; worksheet1.Cells[rowName, 3].Style.Font.Italic = true;
                areaName = 0;
                rowName++;

                //Если имён больше нет, заносим сведения о последней функции
                if (i == roomName.Count - 1)
                {
                    areaFunc = areaFunc.Round(3);
                    areaName = areaName.Round(3);
                    worksheet1.Cells[rowFunc, 1].Value = nameFunc; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                }
            }

            Style(worksheet1, rowName);
            return rowName;
        }
        public int FillCellsArea(ExcelWorksheet worksheet1, int rowFunc, List<String> roomName, List<Element> roomList, string nameFuncRow)
        {
            int rowName = rowFunc + 1;
            double areaName = 0.0;
            double areaFunc = 0.0;
            int j = 0;
            String nameFunc = roomList[0].LookupParameter("PNR_Функция помещения").AsString();
            //Проходимся по всем именам найденных помещений
            for (int i = 0; i < roomName.Count(); i++)
            {
                //Сравниваем помещения на нужное имя
                for (j = 0; j < roomList.Count(); j++)
                {
                    String nameFuncNew = roomList[j].LookupParameter("PNR_Функция помещения").AsString();
                    //Если у помещения нужное имя, суммируем его площадь
                    if (roomList[j].LookupParameter("PNR_Имя помещения").AsString() == roomName[i])
                    {
                        areaName += UnitUtils.ConvertFromInternalUnits(roomList[j].get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                        double areaOneRoom = UnitUtils.ConvertFromInternalUnits(roomList[j].get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                        //Если функция у помещений не меняется, суммируем её
                        if (nameFuncNew == nameFunc)
                        {
                            areaFunc += areaOneRoom;
                        }
                        //Если функция поменялась, заносим в ячейку сведения о функции
                        else
                        {
                            areaFunc = areaFunc.Round(3);
                            areaName = areaName.Round(3);
                            worksheet1.Cells[rowFunc, 1].Value = nameFuncRow; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                            areaFunc = areaOneRoom;
                            rowFunc = rowName;
                            rowName++;
                        }
                    }
                }
                //После того, как прошлись по всем помещениям и сравнили их на имя заносим в строку сведения о площади помещения
                areaFunc = areaFunc.Round(3);
                areaName = areaName.Round(3);
                worksheet1.Cells[rowName, 1].Value = roomName[i]; worksheet1.Cells[rowName, 1].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 2].Value = areaName; worksheet1.Cells[rowName, 2].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 3].Value = "кв.м"; worksheet1.Cells[rowName, 3].Style.Font.Italic = true;
                areaName = 0;
                rowName++;

                //Если имён больше нет, заносим сведения о последней функции
                if (i == roomName.Count - 1)
                {
                    areaFunc = areaFunc.Round(3);
                    areaName = areaName.Round(3);
                    worksheet1.Cells[rowFunc, 1].Value = nameFuncRow; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                }
            }

            Style(worksheet1, rowName);
            return rowName;
        }
        public int FillCellsAreaUnder(ExcelWorksheet worksheet1, int rowFunc, List<String> roomName, List<Element> roomList, String nameCells)
        {
            int rowName = rowFunc + 1;
            double areaName = 0.0;
            double areaFunc = 0.0;
            String nameFunc = roomList[0].LookupParameter("PNR_Функция помещения").AsString();
            for (int i = 0; i < roomName.Count(); i++)
            {
                for (int j = 0; j < roomList.Count(); j++)
                {
                    String nameFuncNew = roomList[j].LookupParameter("PNR_Функция помещения").AsString();
                    if (roomList[j].LookupParameter("PNR_Имя помещения").AsString() == roomName[i])
                    {
                        areaName += UnitUtils.ConvertFromInternalUnits(roomList[j].get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                        if (nameFuncNew == nameFunc)
                        {
                            areaFunc += areaName;
                        }
                        else
                        {
                            areaFunc = areaFunc.Round(3);
                            areaName = areaName.Round(3);
                            worksheet1.Cells[rowFunc, 1].Value = nameCells; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                            worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                            nameFunc = nameFuncNew;
                            areaFunc = areaName;
                            rowFunc = rowName;
                            rowName++;
                        }
                    }
                }
                areaFunc = areaFunc.Round(3);
                areaName = areaName.Round(3);
                worksheet1.Cells[rowName, 1].Value = roomName[i]; worksheet1.Cells[rowName, 1].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 2].Value = areaName; worksheet1.Cells[rowName, 2].Style.Font.Italic = true;
                worksheet1.Cells[rowName, 3].Value = "кв.м"; worksheet1.Cells[rowName, 3].Style.Font.Italic = true;
                areaName = 0.0;
                rowName++;
                if (i == roomList.Count - 1)
                {
                    areaFunc = areaFunc.Round(3);
                    areaName = areaName.Round(3);
                    worksheet1.Cells[rowFunc, 1].Value = nameCells; worksheet1.Cells[rowFunc, 1].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 2].Value = areaFunc; worksheet1.Cells[rowFunc, 2].Style.Font.Bold = true;
                    worksheet1.Cells[rowFunc, 3].Value = "кв.м"; worksheet1.Cells[rowFunc, 3].Style.Font.Bold = true;
                }
            }

            Style(worksheet1, rowName);
            return rowName;
        }
        public List<Element> RoomList(Document doc, List<string> func)
        {
            var allFuncRoom = new FilteredElementCollector(doc);
            //Лист всех комнат
            List<Element> roomList = new List<Element>();
            foreach (var f in func)
            {
                roomList.AddRange(allFuncRoom
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("PNR_Функция помещения").AsString() == f).ToList());
            }
            return roomList;
        }
        public List<Element> RoomListFloor(Document doc, List<string> func, string nameFloor)
        {
            var allFuncRoom = new FilteredElementCollector(doc);
            //Лист всех комнат
            List<Element> roomList = new List<Element>();
            foreach (var f in func)
            {
                roomList.AddRange(allFuncRoom
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("PNR_Функция помещения").AsString() == f)
                    .Where(g => g.LookupParameter("ADSK_Этаж").AsString() == nameFloor)
                    .ToList());
            }
            return roomList;
        }
        public List<Element> RoomListUnder(Document doc, List<string> func, string parameter, int value)
        {
            var allFuncRoom = new FilteredElementCollector(doc);
            //Лист всех комнат
            List<Element> roomList = new List<Element>();
            foreach (var f in func)
            {
                roomList.AddRange(allFuncRoom
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("PNR_Функция помещения").AsString() == f)
                    .Where(g => int.Parse(g.LookupParameter(parameter).AsString()) < value).ToList());
            }
            return roomList;
        }
        public List<Element> RoomListEquals(Document doc, List<string> func, string parameter, int value)
        {
            var allFuncRoom = new FilteredElementCollector(doc);
            //Лист всех комнат
            List<Element> roomList = new List<Element>();
            foreach (var f in func)
            {
                roomList.AddRange(allFuncRoom
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("PNR_Функция помещения").AsString() == f)
                    .Where(g => int.Parse(g.LookupParameter(parameter).AsString()) == value).ToList());
            }
            return roomList;
        }
        public List<String> RoomNames(List<Element> roomList)
        {
            var roomNames = new List<String>();
            foreach (var room in roomList)
            {
                roomNames.Add(room.LookupParameter("PNR_Имя помещения").AsString());
            }
            return roomNames.Distinct().ToList();
        }
        public List<string> Func(String NameFuncStart, String NameFuncEnd)
        {
            List<String> function = new List<String>();
            using (var package = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                string range = "A2:B1000";
                var rangeCells = worksheet.Cells[range];
                object[,] Allvalues = rangeCells.Value as object[,];
                if (Allvalues != null)
                {
                    int rows = Allvalues.GetLength(0);
                    int columns = Allvalues.GetLength(1);
                    int start = 0;

                    for (int i = 0; i <= rows; i++)
                    {
                        int j = 1;
                        if (Allvalues[i, j].ToString() == NameFuncStart)
                        {
                            start = i;
                            break;
                        }
                    }
                    for (int i = start; i < rows; i++)
                    {
                        if (Allvalues[i, 1].ToString() == NameFuncEnd)
                            break;
                        function.Add(Allvalues[i, 1].ToString());
                    }
                }
            }
            return function.Distinct().ToList();
        }
        public void Style(ExcelWorksheet worksheet1, int rowName)
        {
            var rangeStyleA = worksheet1.Cells["A1" + ":" + "A" + (rowName - 1).ToString()];
            rangeStyleA.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleA.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleA.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleA.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

            var rangeStyleB = worksheet1.Cells["B1" + ":" + "B" + (rowName - 1).ToString()];
            rangeStyleB.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rangeStyleB.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rangeStyleB.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rangeStyleB.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            var rangeStyleC = worksheet1.Cells["C1" + ":" + "C" + (rowName - 1).ToString()];
            rangeStyleC.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleC.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleC.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            rangeStyleC.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
        }
    }
}
