using ADSK_Section;
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

namespace ADSK_Section
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

            String parameterValue = LevelSelection.parameterValue;

            List<Wall> walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            TaskDialog.Show("result", $"{LevelSelection.y}");

            foreach( var wall in walls )
            {
                if(wall.Lo)
            }
        }

        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

    }
}

