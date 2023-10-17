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
    /// Логика взаимодействия для SelectSection.xaml
    /// </summary>
    public partial class SelectSection : Window
    {
        public SelectSection(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            InitializeComponent();
            var numberSelection = new NumberSelection(uiapp, uidoc, doc);
            numberSelection.CloseRequest += (s, e) => this.Close();
            DataContext = numberSelection;
        }
    }
}
