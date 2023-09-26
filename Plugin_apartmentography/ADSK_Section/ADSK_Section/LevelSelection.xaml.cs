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
    public partial class LevelSelection : Window
    {
        private ExternalCommandData _commandData;

        static public Autodesk.Revit.DB.Grid selectedLevel;

        static public String parameterValue;

        public static double y = 0.0;

        public LevelSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;

            Document doc = _commandData.Application.ActiveUIDocument.Document;

            List<Autodesk.Revit.DB.Grid> grids = new List<Autodesk.Revit.DB.Grid>(
                new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                .Cast<Autodesk.Revit.DB.Grid>());


            foreach (var grid in grids)
            {
                if(grid.Name == "Б")
                {
                    LB.Items.Add(grid);
                    y = grid.Curve.GetEndPoint(1).Y;
                }
                LB.DisplayMemberPath = "Name";
            }

            TB.Text = parameterValue;

            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;
        }
        public void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LB.SelectedItem;
            selectedLevel = (Autodesk.Revit.DB.Grid)selected;
        }

/*        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }*/
    }
}
