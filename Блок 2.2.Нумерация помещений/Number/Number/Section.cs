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
        public DelegateCommand SelectNotSection { get; } //Делегат кнопки "Продолжить без секции" для секций
        public List<string> ListSections { get; set; } = new List<string>(); //Вспомогательный лист для ComboBox
        public String SelectedSectionValue { get; set; } //выбранная секция в ComboBox
        public int _v { get; set; } //выбранный метод нумерации квартир
        public Section(UIApplication uiapp, UIDocument uidoc, Document doc, int v)
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            SelectSection = new DelegateCommand(SetSelectSection);
            SelectNotSection = new DelegateCommand(NotSelectSection);
            _v = v;
            CreateListSection(); //Заносим все секции в ComboBox
        }

        private void NotSelectSection()
        {
            RaiseCloseRequest();
            var apartWindow = new Apart(_uiapp, _uidoc, _doc, "", _v);
            apartWindow.ShowDialog();
        }

        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        private void SetSelectSection()
        {
            RaiseCloseRequest();
            if(SelectedSectionValue == null)
            {
                TaskDialog.Show("Ошибка выбора", "Не выбрана секция");
                return;
            }
            var apartWindow = new Apart(_uiapp, _uidoc, _doc, SelectedSectionValue, _v);
            apartWindow.ShowDialog();
        }

        private void CreateListSection() //Список всех секций
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
