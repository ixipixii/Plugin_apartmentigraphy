using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Plugin_Kvartiry2
{
    public class Selection
    {
        private ExternalCommandData _commandData;
        public static List<Element> elements = new List<Element>();
        public static Group group;
        public static String section = String.Empty;
        public static String level = String.Empty;
        public static int position = 0;
        public DelegateCommand SelectionLevel { get; }
        public DelegateCommand Appoint { get; }
        public DelegateCommand Сontinue { get; }

        public Selection(ExternalCommandData commandData)
        {
            _commandData = commandData;
            SelectionLevel = new DelegateCommand(OnSelectionLevel);
            Appoint = new DelegateCommand(OnAppoint);
            Сontinue = new DelegateCommand(OnContinue);
        }

        private void OnContinue()
        {
            RaiseCloseRequest();
            Select();
            var window = new CreatingApartment(_commandData);
            window.ShowDialog();
        }

        private void OnAppoint()
        {
            List<ElementId> elementsId = new List<ElementId>();
            foreach(var element in elements)
            {
                elementsId.Add(element.Id);
            }

            Transaction transaction = new Transaction(_commandData.Application.ActiveUIDocument.Document, "Create group");
            transaction.Start();
            group = _commandData.Application.ActiveUIDocument.Document.Create.NewGroup(elementsId);

            if (group.LookupParameter("ADSK_Номер квартиры") == null)
            {
                CreateParameter();
            }
            group.LookupParameter("ADSK_Номер квартиры").Set($"{section}.{level}.{position}");
            elements.Clear();
            elementsId.Clear();
            transaction.Commit();
        }

        private void CreateParameter()
        {
            DefinitionFile definitionFile = _commandData.Application.Application.OpenSharedParameterFile();
            if (definitionFile == null)
                TaskDialog.Show("Ошибка", "Не найден ФОП");

            Definition definition = definitionFile.Groups.SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals("ADSK_Номер квартиры"));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            CategorySet categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(_commandData.Application.ActiveUIDocument.Document, BuiltInCategory.OST_IOSModelGroups));

            Binding binding = _commandData.Application.Application.Create.NewInstanceBinding(categorySet);

            BindingMap map = _commandData.Application.ActiveUIDocument.Document.ParameterBindings;
            map.Insert(definition, binding, BuiltInParameterGroup.PG_IDENTITY_DATA);
        }

        private void OnSelectionLevel()
        {
            RaiseCloseRequest();
            Select();
            var window = new CreatingApartment(_commandData);
            window.ShowDialog();
        }

        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        public void Select()
        {
            var uiapp = _commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = _commandData.Application.ActiveUIDocument.Document;

            uidoc.ActiveView = LevelSelection.selectedLevel;

            var roomFilter = new GroupPickFilter();
            var selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите помещения");

            //var selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, "Выберите комнаты");

            foreach(var selectedElement in selectedRef)
            {
                elements.Add(doc.GetElement(selectedElement));
            }

            section = elements[0].LookupParameter("PNR_Корпус").AsString();
            level = elements[0].LookupParameter("PNR_Этаж").AsString();

            List<Element> groups = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();

            if(groups.Count > 0)
                position = groups.Count + 1;
            else position = 1;
        }

        public class GroupPickFilter : ISelectionFilter
        {
            public bool AllowElement(Element e)
            {
                return e is Room;
            }
            public bool AllowReference(Reference r, XYZ p)
            {
                return false;
            }

        }
    }
}
