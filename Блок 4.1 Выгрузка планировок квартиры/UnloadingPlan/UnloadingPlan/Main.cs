using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace UnloadingPlan
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            Transaction tr = new Transaction(doc, "g");
            tr.Start();
            var options = new ImageExportOptions();
            options.ExportRange = ExportRange.VisibleRegionOfCurrentView;
            options.ViewName = "image";
            var result = doc.SaveToProjectAsImage(options);
            tr.Commit();
            return Result.Succeeded;
        }


    }
}