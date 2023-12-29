﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

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
            System.Diagnostics.Process.Start(excelPath);
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
                                                    .Where(e => e.LookupParameter(filterParameter_1).AsString().Contains(valueFilterParameter_1) &&
                                                                e.LookupParameter(filterParameter_2).AsString().Contains(valueFilterParameter_2))
                                                    .ToList();

            return elements;
        }
        public virtual List<Element> Elements(BuiltInCategory category, Document doc, String filterParameter, int valueFilterParameter, bool valeuCompare)
        {
            //Если знак >= то valeuCompare = 0. Если знак <= то valeuCompare = 1.
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
                                        .Where(e => int.Parse(e.LookupParameter(filterParameter).AsString()) <= (valueFilterParameter))
                                        .ToList();
                return elements;
            }
        }
        public virtual double Areas(List<Element> elements, Document doc)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                    area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
            }
            return area;
        }
        public virtual double Areas(List<Element> elements, Document doc, String PARAMETER)
        {
            double area = 0;
            foreach (Element element in elements)
            {
                try
                {
                    if (element.LookupParameter("PNR_Имя помещения").AsString().Contains(PARAMETER))
                    {
                        area += UnitUtils.ConvertFromInternalUnits(element.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble(), UnitTypeId.SquareMeters);
                    }
                }
                catch { }
            }
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

    }
}
