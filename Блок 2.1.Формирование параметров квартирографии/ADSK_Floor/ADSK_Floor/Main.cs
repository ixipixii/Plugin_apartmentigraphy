using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ADSK_Floor
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public static List<ViewPlan> views = new List<ViewPlan>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            var window = new LevelSelection(commandData);
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}