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
using System.Windows.Markup;

namespace Number
{
    internal class NumberSelection
    {
        UIApplication _uiapp;
        UIDocument _uidoc;
        Document _doc;

        public DelegateCommand SelectCommand { get; }
        public NumberSelection(UIApplication uiapp, UIDocument uidoc, Document doc) 
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            SelectCommand = new DelegateCommand(Selection);
        }

        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        public void Selection()
        {
            RaiseCloseRequest();
            var roomFilter = new GroupPickFilter();
            /*            var selectedRef = _uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите помещения");
                        List<Element> RoomListElement = new List<Element>();

                        foreach (var reference in selectedRef)
                        {
                            Element element = _doc.GetElement(reference);
                            RoomListElement.Add(element);
                            TaskDialog.Show("test", $"{element.Name}");
                        }*/
            Transaction tr = new Transaction(_doc, "s");
            tr.Start();
            IList<Autodesk.Revit.DB.Reference> refrence = new List<Autodesk.Revit.DB.Reference>();
            try
            {
                while (true)
                    refrence.Add(_uidoc.Selection.PickObject(ObjectType.Element, roomFilter, "Select lement, esc when finished"));
            }
            catch { }

            foreach (var _ref in refrence)
            {
                Element element = _doc.GetElement(_ref);
                TaskDialog.Show("test", $"{element.Name}");
            }
            tr.Commit();

            //NumberRoom(RoomListElement);
        }

        public void NumberRoom(List<Element> RoomListElement)
        {
            foreach(var element in RoomListElement)
            {
                String roomFunction = element.LookupParameter("PNR_Функция помещения").AsString();
                String roomSection = element.LookupParameter("ADSK_Номер секции").AsString();
                String roomBuilding = element.LookupParameter("ADSK_Номер здания").AsString();

            }
        }
    }
    public class GroupPickFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            return e is Room;
        }
        public bool AllowReference(Autodesk.Revit.DB.Reference r, XYZ p)
        {
            return false;
        }

    }
}
