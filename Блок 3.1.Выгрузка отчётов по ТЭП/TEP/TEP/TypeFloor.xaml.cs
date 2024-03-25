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

namespace TEP
{
    /// <summary>
    /// Логика взаимодействия для TypeFloor.xaml
    /// </summary>
    public partial class TypeFloor : Window
    {
        public string Start { get; set; }
        public string End { get; set; }
        public string Sect { get; set; }
        public bool Type { get; set; }
        public TypeFloor()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Start = StartFloor.Text;
            End = EndFloor.Text;
            Sect = Section.Text;
            Type = true;

            DialogResult = true;
            Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Type = false;
            Close();
        }
    }
}
