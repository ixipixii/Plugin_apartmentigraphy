using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

namespace Number
{
    /// <summary>
    /// Логика взаимодействия для Renumber.xaml
    /// </summary>
    public partial class Renumber : Window
    {
        public Renumber(UIApplication uiapp, UIDocument uidoc, Document doc, string SelectedSectionValue, int v, List<Element> AllRoomsRenumber)
        {
            InitializeComponent();
            var numberSelection = new NumberSelection(uiapp, uidoc, doc, SelectedSectionValue, v);
            numberSelection.AllRoomsRenumber = AllRoomsRenumber;
            numberSelection.CloseRequest += (s, e) => this.Close();
            DataContext = numberSelection;
        }
    }
}
