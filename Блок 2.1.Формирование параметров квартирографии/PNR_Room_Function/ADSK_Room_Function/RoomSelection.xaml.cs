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

            LVR.Items.Add("Квартиры бед доп. отделки");
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
            LVR.Items.Add("Помещения кладовых");
        }

        private void LVR_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LVR.SelectedItem;
            RoomFunction = selected.ToString();
        }
    }
}
