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
using System.Windows.Media;

namespace TEP
{
    internal abstract class Unloading
    {
        public Unloading() { }
        static public string GetExeDirectory() //Получить путь до DLL
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);
            return path;
        }
        static public string CopyFile(string nameFile) //Копировать файл-шаблон
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
            try
            {
                fileInf.CopyTo(selectedFilePath);
            }
            catch (System.IO.FileNotFoundException)
            {
                var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), $@"Autodesk\Revit\Addins\{nameFile}.xlsx"));
                FileInfo fileInfTemplate = new FileInfo(path.FullName);
                try
                {
                    fileInfTemplate.CopyTo(selectedFilePath);
                }
                catch (System.IO.IOException)
                {
                    FileInfo deleteFile = new FileInfo(selectedFilePath);
                    deleteFile.Delete();
                    fileInfTemplate.CopyTo(selectedFilePath);
                }

            }
            catch (System.IO.IOException)
            {
                FileInfo deleteFile = new FileInfo(selectedFilePath);
                deleteFile.Delete();
                fileInf.CopyTo(selectedFilePath);
            }
            //Возвращаем путь
            return selectedFilePath;
        }
        public void FillCell(int rowIndex, int columnIndex, String cellValue, String excelPath) //Заполнить ячейку
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ТЭП по модели АР"];

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
            if (cellValue == "0")
                return rowIndex;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ТЭП по модели АР"];

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

        //Список элементов нужной категории
        public virtual List<Data> Elements(BuiltInCategory category, Document doc)
        {
            List<Element> elements = new FilteredElementCollector(doc)
                                                    .OfCategory(category)
                                                    .WhereElementIsNotElementType()
                                                    .ToList();
            List<Data> rooms = new List<Data>();
            foreach (var room in elements)
            {
                Data data = new Data();
                data.element = room;
                data.function = room.LookupParameter("PNR_Функция помещения").AsString();
                data.name = room.LookupParameter("PNR_Имя помещения").AsString();
                data.section = room.LookupParameter("ADSK_Номер секции").AsString();
                data.number_apart = room.LookupParameter("ADSK_Номер квартиры").AsString();
                data.number_room = room.LookupParameter("PNR_Номер помещения").AsString();
                data.floor = room.LookupParameter("ADSK_Этаж").AsString();
                data.area = UnitUtils.ConvertFromInternalUnits(room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                data.height = double.Parse(room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString());
                rooms.Add(data);
            }

            return rooms;
        }
        public virtual List<Element> Elements(BuiltInCategory category, Document doc, String filterParameter, int valueFilterParameter, bool valeuCompare)
        {
            //Если знак >= то valeuCompare = 1. Если знак < то valeuCompare = 0.
            if (valeuCompare)
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

        //Формируем списки по нужным параметрам
        public virtual List<Data> ParameterValueEquals(List<Data> list, string parameter, string value)
        {
            List<Data> newList = new List<Data>();
            if (list != null)
            {
                foreach (var element in list)
                {
                    if (element.element.LookupParameter(parameter) != null)
                    {
                        if (element.element.LookupParameter(parameter).AsString() != null || element.element.LookupParameter(parameter).AsString() != "")
                        {
                            if (parameter == "ADSK_Этаж")
                            {
                                if (element.floor != null && element.floor.TrimStart('0') == value)
                                {
                                    newList.Add(element);
                                    continue;
                                }
                            }
                            if (element.element.LookupParameter(parameter).AsString() == value)
                            {
                                newList.Add(element);
                            }
                        }
                    }
                }
            }
            return newList;
        }
        public virtual List<Data> ParameterValueContains(List<Data> list, string parameter, string value)
        {
            List<Data> newList = new List<Data>();
            if (list != null)
            {
                foreach (var element in list)
                {
                    if (element.element.LookupParameter(parameter) != null)
                    {
                        if (element.element.LookupParameter(parameter).AsString() != null || element.element.LookupParameter(parameter).AsString() != "")
                        {
                            try
                            {
                                if (element.element.LookupParameter(parameter).AsString().Contains(value))
                                {
                                    newList.Add(element);
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }
            return newList;
        }

        public virtual List<Data> ParameterValueContainsOr(List<Data> list, string parameter, string value1, string value2)
        {
            List<Data> newList = new List<Data>();
            if (list != null)
            {
                foreach (var element in list)
                {
                    if (element.element.LookupParameter(parameter) != null)
                    {
                        if (element.element.LookupParameter(parameter).AsString() != null || element.element.LookupParameter(parameter).AsString() != "")
                        {
                            try
                            {
                                if (element.element.LookupParameter(parameter).AsString().Contains(value1) || element.element.LookupParameter(parameter).AsString().Contains(value2))
                                {
                                    newList.Add(element);
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                    }
                }
            }
            return newList;
        }

        //Расчитываем площадь по списку элементов
        public virtual double Areas(List<Data> elements, Document doc)
        {
            double area = 0;
            foreach (var element in elements)
            {
                area += element.area;
            }
            area = area.Round(3);
            return area;
        }
        public virtual double Areas(List<Data> elements, Document doc, String PARAMETER)
        {
            double area = 0;
            foreach (var element in elements)
            {
                try
                {
                    if (element.name != null && element.name.Contains(PARAMETER))
                    {
                        area += element.area;
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }
        public virtual double AreasHeigher(List<Data> elements, Document doc, string height)
        {
            double area = 0;
            foreach (var element in elements)
            {
                try
                {
                    if (element.name != null && element.height > double.Parse(height))
                    {
                        area += element.area;
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }
        public virtual double AreasLower(List<Data> elements, Document doc, string height)
        {
            double area = 0;
            foreach (var element in elements)
            {
                try
                {
                    if (element.name != null && element.height < double.Parse(height))
                    {
                        area += element.area;
                    }
                }
                catch { }
            }
            area = area.Round(3);
            return area;
        }

        //Достаём из списка элементов значения параметров
        public virtual List<String> Values(String parameter, List<Data> elements)
        {
            List<String> values = new List<String>();
            foreach (var element in elements)
            {
                Room room = element.element as Room;
                values.Add(room.LookupParameter(parameter).AsString());
            }
            return values;
        }

        //Расчитываем площадь списка помещений
        public int FillCellsArea(ExcelWorksheet worksheet1, int rowFunc, List<String> roomName, List<Data> roomList)
        {
            int rowName = rowFunc + 1;
            double areaName = 0.0;
            double areaFunc = 0.0;
            int j = 0;
            String nameFunc = roomList[0].function;
            //Проходимся по всем именам найденных помещений
            for (int i = 0; i < roomName.Count(); i++)
            {
                //Сравниваем помещения на нужное имя
                for (j = 0; j < roomList.Count(); j++)
                {
                    String nameFuncNew = roomList[j].function;
                    //Если у помещения нужное имя, суммируем его площадь
                    if (roomList[j].name == roomName[i])
                    {
                        areaName += roomList[j].area;
                        double areaOneRoom = roomList[j].area;
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
        //Расчитываем площадь списка помещений с названием функции
        public int FillCellsArea(ExcelWorksheet worksheet1, int rowFunc, List<String> roomName, List<Data> roomList, string nameFuncRow)
        {
            int rowName = rowFunc + 1;
            double areaName = 0.0;
            double areaFunc = 0.0;
            int j = 0;
            String nameFunc = roomList[0].function;
            //Проходимся по всем именам найденных помещений
            for (int i = 0; i < roomName.Count(); i++)
            {
                //Сравниваем помещения на нужное имя
                for (j = 0; j < roomList.Count(); j++)
                {
                    String nameFuncNew = roomList[j].function;
                    //Если у помещения нужное имя, суммируем его площадь
                    if (roomList[j].name == roomName[i])
                    {
                        areaName += roomList[j].area;
                        double areaOneRoom = roomList[j].area;
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
        public List<Data> RoomListUnder(Document doc, List<string> func, List<Data> element)
        {
            //Лист всех комнат
            List<Data> roomList = new List<Data>();
            int i = 0;
            foreach (var f in func)
            {
                foreach (var e in element)
                {
                    if (e.function == f)
                    {
                        if (int.Parse(e.floor) < 1)
                        {
                            roomList.Add(e);
                        }
                    }
                }
            }
            return roomList;
        }
        public List<Data> RoomListEquals(Document doc, List<string> func, List<Data> element)
        {
            //Лист всех комнат
            List<Data> roomList = new List<Data>();
            foreach (var f in func)
            {
                foreach (var e in element)
                {
                    if (e.function == f)
                    {
                        if (int.Parse(e.floor) == 1)
                        {
                            roomList.Add(e);
                        }
                    }
                }
            }
            return roomList;
        }
        public List<String> RoomNames(List<Data> roomList)
        {
            var roomNames = new List<String>();
            foreach (var room in roomList)
            {
                roomNames.Add(room.name);
            }
            return roomNames.Distinct().ToList();
        }
        public List<string> Func(String NameFuncStart, String NameFuncEnd)
        {
            List<String> function = new List<String>();
            var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Имена помещений.xlsx"));
            using (var package = new ExcelPackage(path))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Name"];
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
        public void StyleCell(ExcelWorksheet worksheet1, int rowName, System.Drawing.Color color)
        {
            var rangeStyleA = worksheet1.Cells["A" + (rowName).ToString()];
            var rangeStyleB = worksheet1.Cells["B" + (rowName).ToString()];
            var rangeStyleC = worksheet1.Cells["C" + (rowName).ToString()];

            rangeStyleA.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rangeStyleB.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rangeStyleC.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

            rangeStyleA.Style.Fill.BackgroundColor.SetColor(color);
            rangeStyleB.Style.Fill.BackgroundColor.SetColor(color);
            rangeStyleC.Style.Fill.BackgroundColor.SetColor(color);

            rangeStyleA.Style.Font.Bold = true;
            rangeStyleB.Style.Font.Bold = true;
            rangeStyleC.Style.Font.Bold = true;
        }
    }
}
