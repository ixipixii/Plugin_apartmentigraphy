using ADSK_Room_Function;
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
using ReitAPIPluginLibrary;

namespace ADSK_Room_Function
{
    public class Selection
    {
        public DelegateCommand SelectionLevel { get; }

        public static List<Element> selectedElement = new List<Element>();
        public DelegateCommand Appoint { get; }

        public static List<ElementId> groupElements = new List<ElementId>();
        public DelegateCommand Сontinue { get; }

        private ExternalCommandData _commandData;

        public Selection(ExternalCommandData commandData)
        {
            _commandData = commandData;
            SelectionLevel = new DelegateCommand(OnSelectionLevel);
            Appoint = new DelegateCommand(OnAppoint);
            Сontinue = new DelegateCommand(OnContinue);
            Select();
        }

        private void OnContinue()
        {
            RaiseCloseRequest();
            //Select();
            var window = new RoomSelection(_commandData);
            window.ShowDialog();
        }

        private void OnAppoint()
        {
            //RaiseCloseRequest();
            using (Transaction tr = new Transaction(_commandData.Application.ActiveUIDocument.Document, "Start"))
            {
                if (selectedElement != null)
                {
/*                    if (selectedElement.GroupId.ToString() != "-1")
                    {
                        Group group = (Group)_commandData.Application.ActiveUIDocument.Document.GetElement(selectedElement.GroupId);
                        Transaction t = new Transaction(_commandData.Application.ActiveUIDocument.Document, "UnGroup");
                        t.Start();
                        groupElements = group.UngroupMembers().ToList();
                        t.Commit();
                    }*/
                }
                
                tr.Start();
                foreach (var element in selectedElement)
                {
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_Rooms)
                    {
                        if (element.LookupParameter("PNR_Функция помещения") == null)
                        {
                            var categorySet = new CategorySet();
                            categorySet.Insert(element.Category);
                            CreateShared createShared = new CreateShared();
                            createShared.CreateSharedParameter(_commandData.Application.Application,
                                                       _commandData.Application.ActiveUIDocument.Document,
                                                       "PNR_Функция помещения",
                                                       categorySet,
                                                       BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                       true);
                        }
                        element.LookupParameter("PNR_Функция помещения").Set(RoomSelection.RoomFunction);
                    }
                    else
                    {
                        TaskDialog.Show("Ошибка", "Выбранный элемент не является помещением");
                        tr.RollBack();
                        OnContinue();
                        RaiseCloseRequest();
                        return;
                    }
                }

                selectedElement.Clear();

                tr.Commit();

/*                tr.Start();
                if (groupElements.Count > 0)
                {
                    _commandData.Application.ActiveUIDocument.Document.Create.NewGroup(groupElements);
                    groupElements.Clear();
                }
                tr.Commit();*/
            }
        }

        private void OnSelectionLevel()
        {
            RaiseCloseRequest();
            //Select();
            var window = new RoomSelection(_commandData);
            window.ShowDialog();
        }


        public void Select()
        {
            var uiapp = _commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = _commandData.Application.ActiveUIDocument.Document;

            var roomFilter = new GroupPickFilter();
            var selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите комнаты");


            foreach(var references in selectedRef)
            {
                selectedElement.Add(doc.GetElement(references));
            }

            //selectedElement = doc.GetElement(selectedRef);
        }



        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
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
