using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.Formula.Functions;
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
using static System.Collections.Specialized.BitVector32;

namespace Plugin_Kvartiry2
{
    /// <summary>
    /// Логика взаимодействия для CreatingApartment.xaml
    /// </summary>
    public partial class CreatingApartment : Window
    {
        public CreatingApartment(ExternalCommandData commandData)
        {
            InitializeComponent();
            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;

            LVR.Items.Add($"ADSK_Номер квартиры: {Selection.section}.{Selection.level}.{Selection.index[Selection.index.Count + 1]}");
        }
    }
}
