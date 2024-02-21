using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TEP
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            TypeFloor typeFloor = new TypeFloor();
            typeFloor.ShowDialog();

            TEP_AR tEP_AR = new TEP_AR(uiapp, uidoc, doc, typeFloor.Start, typeFloor.End, typeFloor.Sect);

            return Result.Succeeded;
        }


    }
}