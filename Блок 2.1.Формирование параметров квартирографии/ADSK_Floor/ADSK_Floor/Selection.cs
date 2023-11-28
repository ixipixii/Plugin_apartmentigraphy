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

namespace ADSK_Floor
{
    public class Selection
    {
        public DelegateCommand SelectionLevel { get; }

        private ExternalCommandData _commandData;
        public Selection(ExternalCommandData commandData)
        {
            _commandData = commandData;
            SelectionLevel = new DelegateCommand(OnSelectionLevel);
        }

        private void OnSelectionLevel()
        {
            RaiseCloseRequest();
            Select();
        }

        public void Select()
        {
            var uiapp = _commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = _commandData.Application.ActiveUIDocument.Document;

            //Делаем активным вид, который выбрал пользователь
            uidoc.ActiveView = LevelSelection.selectedLevel;

            //Считываем помещения
            List<Room> collector = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .ToList();

            Transaction tr = new Transaction(doc, "Rename level");
            tr.Start();

            foreach (Room room in collector)
            {
                if (room != null)
                {
                    if(room.LookupParameter("ADSK_Этаж") == null)
                    {
                        var categorySet = new CategorySet();
                        categorySet.Insert(room.Category);
                        CreateShared createShared = new CreateShared();
                        createShared.CreateSharedParameter(uiapp.Application,
                                                   doc,
                                                   "ADSK_Этаж",
                                                   categorySet,
                                                   BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                   true);
                    }
                    room.LookupParameter("ADSK_Этаж").Set($"{LevelSelection.parameterValue}");
                }
            }

            tr.Commit();

        }

        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

    }
}

