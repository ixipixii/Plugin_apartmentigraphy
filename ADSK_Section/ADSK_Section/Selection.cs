using ADSK_Section;
using Autodesk.Revit.Creation;
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

namespace ADSK_Section
{
    public class Data
    {
        public double X { get; set; }
        public double Y { get; set; }
        public Room Room { get; set; }
        public Data() { }
    }
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
            Autodesk.Revit.DB.Document doc = _commandData.Application.ActiveUIDocument.Document;

            //String parameterValue = LevelSelection.nameSection;

            //Считываем помещения
            List<Room> collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .ToList();

            List<Data> elements = new List<Data>();

            //Заполняем координаты помещений
            foreach( var element in collector)
            {
                double x = 0.0;
                double y = 0.0;
                LocationPoint LP = element.Location as LocationPoint;
                if ( LP != null)
                {
                    x = LP.Point.X;
                    y = LP.Point.Y;  
                    Data room = new Data();
                    room.X = x;
                    room.Y = y;
                    room.Room = element;
                    elements.Add(room);
                }
            }

            //Если координаты точки помещения входят в границы секции, заносим текст в параметр ADSK_Секция
            Transaction tr = new Transaction(doc, "Create parameter");
            tr.Start();

            foreach (var element in elements)
            {
                if (element.Y > LevelSelection.valueY.Min() && element.Y < LevelSelection.valueY.Max())
                {
                    if(element.X > LevelSelection.valueX.Min() && element.X < LevelSelection.valueX.Max())
                    {
                        if(element.Room.LookupParameter("ADSK_Секция") == null)
                        {
                            var categorySet = new CategorySet();
                            categorySet.Insert(element.Room.Category);
                            CreateShared createShared = new CreateShared();
                            createShared.CreateSharedParameter(uiapp.Application,
                                                       doc,
                                                       "ADSK_Секция",
                                                       categorySet,
                                                       BuiltInParameterGroup.PG_IDENTITY_DATA,
                                                       true);
                        }
                        TaskDialog.Show("test", $"{element.Room.Name}");
                        element.Room.LookupParameter("ADSK_Секция").Set(LevelSelection.nameSection);
                    }
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

