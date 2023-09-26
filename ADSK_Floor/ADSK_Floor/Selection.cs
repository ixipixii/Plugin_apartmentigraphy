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

            uidoc.ActiveView = LevelSelection.selectedLevel;
            String parameterValue = LevelSelection.parameterValue;

        }

        public event EventHandler CloseRequest;

        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

    }
}

