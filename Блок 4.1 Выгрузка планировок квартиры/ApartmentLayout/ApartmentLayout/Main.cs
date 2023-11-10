using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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

            //Фильтр по всем помещениям на данном виде
            var AllRooms = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                            .WhereElementIsNotElementType()
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .ToList();

            Room room = null;

            foreach(var rooms in AllRooms)
            {
                if(rooms.Name == "Гостиная 256")
                {
                    room = (Room)rooms; 
                    break;
                }
            }

            Transaction tr = new Transaction(doc, "view");
            tr.Start();



            //CropAroundRoom(room, uidoc.ActiveView);
            tr.Commit();

            return Result.Succeeded;
        }

        public void CropAroundRoom(Room room, View view)
        {
            if (view != null)
            {
                IList<IList<Autodesk.Revit.DB.BoundarySegment>> segments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());

                if (null != segments)  //комната может быть не связана
                {
                    foreach (IList<Autodesk.Revit.DB.BoundarySegment> segmentList in segments)
                    {
                        CurveLoop loop = new CurveLoop();
                        foreach (Autodesk.Revit.DB.BoundarySegment boundarySegment in segmentList)
                        {
                            loop.Append(boundarySegment.GetCurve());
                        }

                        ViewCropRegionShapeManager vcrShapeMgr = view.GetCropRegionShapeManager();
                        vcrShapeMgr.SetCropShape(loop);
                        break;  // если для комнаты имеется более одного набора граничных сегментов, обрежьте первый из них
                    }
                }
            }
        }


    }

    internal class WallSegment
    {
        public Wall Wall { get; set; }
        public Curve Curve { get; set; }
    }
}