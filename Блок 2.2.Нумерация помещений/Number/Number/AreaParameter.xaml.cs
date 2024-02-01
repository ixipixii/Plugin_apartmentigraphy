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
    /// Логика взаимодействия для AreaParameter.xaml
    /// </summary>
    public partial class AreaParameter : Window
    {
        public string selectedType = string.Empty;
        public string selectedCount = string.Empty;
        public string selectedRange = string.Empty;
        public AreaParameter()
        {
            InitializeComponent();
            Type.ItemsSource = new object[] { "S", "M", "L" };
            Count.ItemsSource = new object[] { "1", "2", "3", "4", "5" };
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            selectedType = Type.SelectedItem.ToString();
            selectedCount = Count.SelectedItem.ToString();
            selectedRange = Range.Text;
            this.Close();
        }
    }
}
