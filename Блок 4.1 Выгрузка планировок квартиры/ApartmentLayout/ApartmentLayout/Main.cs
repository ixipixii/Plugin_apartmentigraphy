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

namespace ApartmentLayout
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = commandData.Application.ActiveUIDocument.Document;

            //Лист точек
            List<XYZ> points = new List<XYZ>();

            //Собираем все точки комнаты
            while (true)
            {
                try
                {
                    XYZ point = uidoc.Selection.PickPoint();
                    points.Add(point);
                }
                catch (Exception ex)
                {
                    break;
                }
            }

            // Создание объекта цепочки кривых
            CurveLoop loop = new CurveLoop();

            for (int i = 1; i < points.Count; i++)
            {
                loop.Append(Line.CreateBound(points[i - 1], points[i]));
                if (i == points.Count - 1)
                    loop.Append(Line.CreateBound(points[i], points[0]));
            }

            Transaction tr = new Transaction(doc, "view");
            tr.Start();
            // Назначение границ подрезки
            uidoc.ActiveView.GetCropRegionShapeManager().SetCropShape(loop);
            tr.Commit();

            return Result.Succeeded;
        }
    }
}