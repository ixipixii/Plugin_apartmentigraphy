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
using static OfficeOpenXml.ExcelErrorValue;

namespace TEP
{
    [Transaction(TransactionMode.Manual)]
    internal class REPORT_YK : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            String path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Отчёты.xlsx");          
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ТЭП по модели АР"];
                ExcelWorksheet worksheet_YK = package.Workbook.Worksheets["Отчёт УК"];

                //Диапазон для чтения
                string range = "A2:C300";

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
                    
                    for(int i = start; i <= rows; i++)
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

            System.Diagnostics.Process.Start(path);

            return Result.Succeeded;
        }
    }
}
