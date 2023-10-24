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
    /// Логика взаимодействия для MainPanel.xaml
    /// </summary>
    public partial class MainPanel : Window
    {
        public MainPanel(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            InitializeComponent();
            var numberSelection = new NumberSelection(uiapp, uidoc, doc, "0");
            numberSelection.CloseRequest += (s, e) => this.Close();
            DataContext = numberSelection;           
        }
    }
}
