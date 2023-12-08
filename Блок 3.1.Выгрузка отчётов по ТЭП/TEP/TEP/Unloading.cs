using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
    internal class Unloading
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
            TaskDialog.Show("s", $"{absPath}");
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
        public void FillCell(int rowIndex, int columnIndex, String cellValue) //Заполнить ячейку
        {
            //Открываем нужный файл
            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "s.xlsx");
            //Записываем значение в нужную папку
            /*using(FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite)) 
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.GetSheet("Лист1");
                //ICell cell;
                try
                {
                    IRow row = sheet.CreateRow(rowIndex);
                    ICell cell = row.CreateCell(columnIndex);
                    //cell = sheet.GetRow(0).GetCell(0);
                    cell.SetCellValue(cellValue);
                    workbook.Write(stream);
                    workbook.Close();
                }
                catch (Exception ex) { TaskDialog.Show(ex.Source, ex.Message); }

                //workbook.Close();

                *//*                using (FileStream fs = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
                                {
                                    workbook.Write(fs);
                                }*//*


            }*/

            //System.Diagnostics.Process.Start(excelPath);

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                // Выбираем лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Координаты ячейки, которую вы хотите изменить
                int rowNumber = 2;
                int columnNumber = 2;

                // Новое значение для ячейки
                string newValue = "Новое значение";

                // Устанавливаем значение ячейки
                worksheet.Cells[rowNumber, columnNumber].Value = newValue;

                // Сохраняем изменения
                package.Save();
            }
        }

    }
}
