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

namespace Plugin_Kvartiry
{
    /// <summary>
    /// Логика взаимодействия для RoomSelection.xaml
    /// </summary>
    public partial class RoomSelection : Window
    {
        static public string NameRoom = String.Empty;
        public RoomSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;

            LVR.Items.Add("Жилая комната");
            LVR.Items.Add("Гостиная");
            LVR.Items.Add("Кухня");
            LVR.Items.Add("Кухня-Ниша");
            LVR.Items.Add("Гардеробная");
            LVR.Items.Add("Коридор");
            LVR.Items.Add("Санузел");
            LVR.Items.Add("Постирочная");
            LVR.Items.Add("Терраса");
            LVR.Items.Add("Кабинет");
            LVR.Items.Add("Балкон, Лоджия");
            LVR.Items.Add("Столовая");
        }

        private void LVR_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LVR.SelectedItem;
            NameRoom = selected.ToString();
        }
    }
}
