using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Number
{
    internal class Section
    {
        UIApplication _uiapp;
        UIDocument _uidoc;
        Document _doc;
        public DelegateCommand SelectSection { get; } //Делегат кнопки "Выбрать" для секций
        public List<string> ListSections { get; set; } = new List<string>(); //Вспомогательный лист для ComboBox
        public String SelectedSectionValue { get; set; } //выбранная секция в ComboBox
        public Section(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            SelectSection = new DelegateCommand(SetSelectSection);
            CreateListSection(); //Заносим все секции в ComboBox
        }

        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        private void SetSelectSection()
        {
            RaiseCloseRequest();
            //TaskDialog.Show("d", $"{SelectedSectionValue}");
            var apartWindow = new Apart(_uiapp, _uidoc, _doc, SelectedSectionValue);
            apartWindow.ShowDialog();
        }

        private void CreateListSection()
        {
            //Фильтруем по всем комнатам
            var rooms = new FilteredElementCollector(_doc)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .WhereElementIsNotElementType()
                            .ToList();

            //Считываем у всех комнат параметр секция
            List<string> nameSections = new List<string>();
            foreach (var room in rooms)
            {
                nameSections.Add(room.LookupParameter("ADSK_Номер секции").AsString());
            }
            //Заносим список в лист-секций для ComboBox
            ListSections = nameSections.Distinct().ToList();
        }

    }
}
