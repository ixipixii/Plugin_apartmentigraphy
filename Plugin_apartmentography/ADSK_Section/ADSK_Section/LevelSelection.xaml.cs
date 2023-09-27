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

        static public String parameterValue;

        //Листы с координатами сетки
        static public List<Double> valueX = new List<double>();
        static public List<Double> valueY = new List<double>();
        public LevelSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;

            Document doc = _commandData.Application.ActiveUIDocument.Document;

            //Собираем все сетки
            List<Autodesk.Revit.DB.Grid> grids = new List<Autodesk.Revit.DB.Grid>(
                new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                .Cast<Autodesk.Revit.DB.Grid>());

            //Вводим координаты
            foreach (var grid in grids)
            {
                Autodesk.Revit.DB.Line line = grid.Curve as Autodesk.Revit.DB.Line;
                if (line != null)
                {
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
                /*if(grid.Name == "3" || grid.Name == "2")
                {
                    LB.Items.Add(grid);
                    valueX.Add(grid.Curve.GetEndPoint(1).X);                   
                }
                if (grid.Name == "А" || grid.Name == "Б")
                {
                    LB.Items.Add(grid);
                    valueY.Add(grid.Curve.GetEndPoint(1).Y);
                }*/
                CB_Gorizontal_1.DisplayMemberPath = "Name";
                CB_Gorizontal_2.DisplayMemberPath = "Name";
                CB_Vertical_1.DisplayMemberPath = "Name";
                CB_Vertical_2.DisplayMemberPath = "Name";
            }

            TB.Text = parameterValue;

            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;
        }
        public void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
/*            Autodesk.Revit.DB.Grid grid_Vertical_1 = (Autodesk.Revit.DB.Grid)CB_Vertical_1.SelectedItem;
            valueX.Add(grid_Vertical_1.Curve.GetEndPoint(1).X);

            Autodesk.Revit.DB.Grid grid_Vertical_2 = (Autodesk.Revit.DB.Grid)CB_Vertical_2.SelectedItem;
            valueX.Add(grid_Vertical_2.Curve.GetEndPoint(1).X);*/

            Autodesk.Revit.DB.Grid grid_Gorizontal_1 = (Autodesk.Revit.DB.Grid)CB_Gorizontal_1.SelectedItem;
            valueX.Add(grid_Gorizontal_1.Curve.GetEndPoint(1).X);

/*            Autodesk.Revit.DB.Grid grid_Gorizontal_2 = (Autodesk.Revit.DB.Grid)CB_Gorizontal_2.SelectedItem;
            valueX.Add(grid_Gorizontal_2.Curve.GetEndPoint(1).X);*/

            /*//Добавляем значения с вертикальных комбобоксов
            valueX.Add(Convert.ToDouble(CB_Vertical_1.SelectedItem.ToString()));
            valueX.Add(Convert.ToDouble(CB_Vertical_2.SelectedItem.ToString()));

            //Добавляем значения с горизонтальных комбобоксов
            valueY.Add(Convert.ToDouble(CB_Gorizontal_1.SelectedItem.ToString()));
            valueY.Add(Convert.ToDouble(CB_Gorizontal_2.SelectedItem.ToString()));*/
        }

/*        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }*/
    }
}
