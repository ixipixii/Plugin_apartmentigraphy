using MathNet.Numerics.LinearAlgebra.Factorization;
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
    /// Логика взаимодействия для Housing.xaml
    /// </summary>
    public partial class Housing : Window
    {
        public string Hous { get; set; }
        public Housing()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Hous = NumberHousing.Text;

            this.Close();
        }
    }
}
