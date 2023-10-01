﻿using Autodesk.Revit.DB;
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

namespace ADSK_Floor
{
    /// <summary>
    /// Логика взаимодействия для LevelSelection.xaml
    /// </summary>
    public partial class LevelSelection : Window
    {
        private ExternalCommandData _commandData;

        static public ViewPlan selectedLevel;

        static public String parameterValue;
        public LevelSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;

            Document doc = _commandData.Application.ActiveUIDocument.Document;

            List<ViewPlan> views = new List<ViewPlan>(
                new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where<ViewPlan>(v => v.CanBePrinted
                && ViewType.FloorPlan == v.ViewType));

            foreach (ViewPlan viewPlan in views)
            {
                LB.Items.Add(viewPlan);
                LB.DisplayMemberPath = "Name";
            }

            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;
        }
        public void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LB.SelectedItem;
            selectedLevel = (ViewPlan)selected;
        }

        private void TB_TextChanged(object sender, TextChangedEventArgs e)
        {
            TB.Text = parameterValue;
        }
    }
}
