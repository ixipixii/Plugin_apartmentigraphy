using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using Autodesk.Revit.Exceptions;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ADSK_Room_Function
{
    /// <summary>
    /// Логика взаимодействия для RoomSelection.xaml
    /// </summary>
    public partial class RoomSelection : Window
    {
        static public string RoomFunction = String.Empty;
        public RoomSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            List<String> function = new List<String>();
            var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Имена помещений.xlsx"));
            using (var package = new ExcelPackage(path))
            {
                TaskDialog.Show("test", $"{path.LastAccessTimeUtc}");
                var count = package.Workbook.Worksheets.Count;
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Name"];
                string range = "A2:B200";
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
                        if (Allvalues[i, j].ToString() == "Квартиры бед доп. отделки")
                        {
                            start = i;
                            break;
                        }
                    }
                    for (int i = start; i < 171; i++)
                    {
                        try
                        {
                            if (Allvalues[i, 1].ToString() == null)
                                break;
                            function.Add(Allvalues[i, 1].ToString());
                        }
                        catch { break; }
                    }
                }
            }

            List<string> func = function.Distinct().ToList();

            foreach(var f in func)
            {
                if (f != null)
                {
                    LVR.Items.Add(f);
                }
            }

/*                LVR.Items.Add("Квартиры бед доп. отделки");
            LVR.Items.Add("Апартаменты без доп. Отделки");
            LVR.Items.Add("Коммерческие помещения без доп. отделки");
            LVR.Items.Add("МОП входной группы 1 этажа");
            LVR.Items.Add("МОП входной группы -1 этажа");
            LVR.Items.Add("МОП типовых этажей");
            LVR.Items.Add("МОП входной группы 1 этажа");
            LVR.Items.Add("МОП типовых этажей");
            LVR.Items.Add("Лестницы эвакуации  (с -1го до последнего этажа)");
            LVR.Items.Add("Паркинг");
            LVR.Items.Add("Квартиры с отделкой");
            LVR.Items.Add("Апартаменты с отделкой");
            LVR.Items.Add("Коммерческие помещения с отделкой");
            LVR.Items.Add("Помещения загрузки");
            LVR.Items.Add("Помещения мусороудаления");
            LVR.Items.Add("Инженерно-технические помещения");
            LVR.Items.Add("Помещение управляющей компании");
            LVR.Items.Add("Объединенный диспетчерский пункт");
            LVR.Items.Add("Помещения линейного и обслуживающего персонала");
            LVR.Items.Add("Помещения охраны");
            LVR.Items.Add("Помещения Клининговых служб");
            LVR.Items.Add("Помещения кладовых");*/
        }

        private void LVR_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LVR.SelectedItem;
            RoomFunction = selected.ToString();
        }
    }
}
