using Autodesk.Revit.DB;
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
        public virtual double Area()
        {
            double area = 0;
            return area;
        }

    }
}
