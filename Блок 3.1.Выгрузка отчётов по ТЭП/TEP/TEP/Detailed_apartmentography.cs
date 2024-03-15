using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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
            string CounrRoom { get; set; }
            string Range { get; set; }
            string Floor { get; set; }
            string Section { get; set; }
            public Info(string countRoom, string range, string floor, string section) => (CounrRoom, Range, Floor, Section) = (countRoom, range, floor, section);
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //Открываем файл-шаблон с детальной квартирографией
            //String pathTemplate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Отчёты.xlsx");

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

            foreach ( var room in rooms )
            {
                listParameter.Add(room.LookupParameter("ADSK_Номер квартиры").AsString());
                listSections.Add(room.LookupParameter("ADSK_Номер секции").AsString());
                listFloor.Add(room.LookupParameter("ADSK_Этаж").AsString());
            }

            IEnumerable<string> uniqueParameterValueParameter = listParameter.Distinct();
            IEnumerable<string> uniqueParameterValueSections  = listSections.Distinct();
            IEnumerable<string> uniqueParameterValueFloor     = listFloor.Distinct();

            Dictionary<string, string> sectionCounts = new Dictionary<string, string>();

            foreach(var section in uniqueParameterValueSections)
            {
                string s = (section.Length == 1) ? "0" + section : section;
                int count = uniqueParameterValueParameter.Count(room => room.Substring(room.Length - 5, 2) == section);
                sectionCounts.Add(s, count.ToString());
            }

            using (var package = new ExcelPackage(new FileInfo(selectedFilePath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Выгрузка на корпус"];

                //Диапазон для чтения
                string range = "A2:S100";

                //Координаты для записи
                int row = 2;
                int col_1 = 1;
                int col_3 = 3;
                int col_4 = 4;

                // Координаты для чтения
                var rangeCells = worksheet.Cells[range];
                object[,] Allvalues = rangeCells.Value as object[,];

                int start = 0;
                if (Allvalues != null)
                {
                    int rows = Allvalues.GetLength(0);

                    for (int i = 0; i <= rows; i++)
                    {
                        if (Allvalues[i, 0].ToString() == "Количество квартир")
                        {
                            worksheet_YK.Cells[row, col_1].Value = Allvalues[i, 0].ToString();
                            worksheet_YK.Cells[row, col_3].Value = Allvalues[i, 1].ToString();
                            worksheet_YK.Cells[row, col_4].Value = Allvalues[i, 2].ToString();
                            row++;
                        }
                        if (Allvalues[i, 0].ToString() == "Суммарная площадь МОП, в том числе:")
                        {
                            start = i;
                            break;
                        }
                    }

                    for (int i = start; i <= rows; i++)
                    {
                        if (Allvalues[i, 0].ToString().Contains("ПОКАЗАТЕЛИ ЭФФЕКТИВНОСТИ"))
                            break;
                        worksheet_YK.Cells[row, col_1].Value = Allvalues[i, 0].ToString();
                        worksheet_YK.Cells[row, col_3].Value = Allvalues[i, 1].ToString();
                        worksheet_YK.Cells[row, col_4].Value = Allvalues[i, 2].ToString();
                        row++;
                    }
                }
                // Сохраняем изменения
                package.Save();
            }

            System.Diagnostics.Process.Start(selectedFilePath);

            return Result.Succeeded;
        }
    }
}
