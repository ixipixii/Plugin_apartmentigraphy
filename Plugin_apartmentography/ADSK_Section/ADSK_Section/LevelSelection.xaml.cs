using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ADSK_Section
{
    /// <summary>
    /// Логика взаимодействия для LevelSelection.xaml
    /// </summary>
    /// 
    public partial class LevelSelection : Window
    {
        private ExternalCommandData _commandData;

        static public String nameSection;

        //Листы с координатами осей
        static public List<Double> valueX = new List<double>();
        static public List<Double> valueY = new List<double>();
        public LevelSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;

            Document doc = _commandData.Application.ActiveUIDocument.Document;

            //Собираем все оси
            List<Autodesk.Revit.DB.Grid> grids = new List<Autodesk.Revit.DB.Grid>(
                new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                .Cast<Autodesk.Revit.DB.Grid>());

            //Проходимся по всем осям проекта
            foreach (var grid in grids)
            {
                Autodesk.Revit.DB.Line line = grid.Curve as Autodesk.Revit.DB.Line;
                if (line != null)
                {
                    //Определяем вертикальные и горизонтальные оси и добавляем на combobox-ы
                    if(line.Direction.X == 1.0)
                    {
                        CB_Gorizontal_1.Items.Add(grid);
                        CB_Gorizontal_2.Items.Add(grid);
                    }
                    else
                    {
                        CB_Vertical_1.Items.Add(grid);
                        CB_Vertical_2.Items.Add(grid);
                    }
                }
                //На combobox-ах будте показано свойство Name объектов-осей
                CB_Gorizontal_1.DisplayMemberPath = "Name";
                CB_Gorizontal_2.DisplayMemberPath = "Name";
                CB_Vertical_1.DisplayMemberPath = "Name";
                CB_Vertical_2.DisplayMemberPath = "Name";
            }

            //Название секции
            TB.Text = nameSection;
            nameSection = "123";

            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;
        }

        //Обработчики события combobox-ов. Сразу добавляем координаты в массивы координат
        private void CB_Gorizontal_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Autodesk.Revit.DB.Grid grid_Gorizontal_1 = (Autodesk.Revit.DB.Grid)CB_Gorizontal_1.SelectedItem;
            valueY.Add(grid_Gorizontal_1.Curve.GetEndPoint(1).Y);
        }

        private void CB_Gorizontal_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Autodesk.Revit.DB.Grid grid_Gorizontal_2 = (Autodesk.Revit.DB.Grid)CB_Gorizontal_2.SelectedItem;
            valueY.Add(grid_Gorizontal_2.Curve.GetEndPoint(1).Y);
        }

        private void CB_Vertical_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Autodesk.Revit.DB.Grid grid_Vertical_1 = (Autodesk.Revit.DB.Grid)CB_Vertical_1.SelectedItem;
            valueX.Add(grid_Vertical_1.Curve.GetEndPoint(1).X);
        }

        private void CB_Vertical_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Autodesk.Revit.DB.Grid grid_Vertical_2 = (Autodesk.Revit.DB.Grid)CB_Vertical_2.SelectedItem;
            valueX.Add(grid_Vertical_2.Curve.GetEndPoint(1).X);
        }
    }
}
