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
    /// Логика взаимодействия для Apart.xaml
    /// </summary>
    public partial class Apart : Window
    {
        public Apart(UIApplication uiapp, UIDocument uidoc, Document doc, string SelectedSectionValue, int v)
        {
            InitializeComponent();
            var numberSelection = new NumberSelection(uiapp, uidoc, doc, SelectedSectionValue, v);
            if(v == 1)
                numberSelection.ApartList(); //Заносим все нужные группы в ListView по 1 варианту
            else
                numberSelection.ApartList_2(); //Заносим все нужные группы в ListView по 2 варианту
            numberSelection.CloseRequest += (s, e) => this.Close();
            DataContext = numberSelection;
        }
    }
}
