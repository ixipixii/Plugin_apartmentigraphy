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

            List<Room> rooms = new List<Room>();

            foreach (var room in AllRooms)
            {
                if (room.LookupParameter("ADSK_Номер квартиры").AsString() == "1/КВ-02004")
                {
                    var room_kv = (Room)room;
                    rooms.Add(room_kv);
                }
            }

            double MAX_X = -10000000000000000000.0;
            double MAX_Y = -10000000000000000000.0;
            double MIN_X = 10000000000;
            double MIN_Y = 100000000000;

            foreach (var room in rooms)
            {
                if (MAX_X < room.get_Geometry(new Options()).GetBoundingBox().Max.X)
                    MAX_X = room.get_Geometry(new Options()).GetBoundingBox().Max.X;
                if (MAX_Y < room.get_Geometry(new Options()).GetBoundingBox().Max.Y)
                    MAX_Y = room.get_Geometry(new Options()).GetBoundingBox().Max.Y;
                if (MIN_X > room.get_Geometry(new Options()).GetBoundingBox().Min.X)
                    MIN_X = room.get_Geometry(new Options()).GetBoundingBox().Min.X;
                if (MIN_Y > room.get_Geometry(new Options()).GetBoundingBox().Min.Y)
                    MIN_Y = room.get_Geometry(new Options()).GetBoundingBox().Min.Y;
            }

            List<XYZ> points = new List<XYZ>();

            XYZ point_1 = null;
            XYZ point_2 = null;
            XYZ point_3 = null;
            XYZ point_4 = null;

            //Фильтр по всем стенам на данном виде
            var AllWalls = new FilteredElementCollector(doc, uidoc.ActiveView.Id)
                            .WhereElementIsNotElementType()
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .ToList();

            foreach (var wall in AllWalls)
            {
                //Добавили ширину сверху
                if (Math.Round(wall.get_Geometry(new Options()).GetBoundingBox().Min.Y, 4) == Math.Round(MAX_Y, 4))
                {
                    point_1 = new XYZ(wall.get_Geometry(new Options()).GetBoundingBox().Max.X,
                                           wall.get_Geometry(new Options()).GetBoundingBox().Max.Y,
                                           wall.get_Geometry(new Options()).GetBoundingBox().Max.Z);
                }
                //Добавили ширину слева
                if (Math.Round(wall.get_Geometry(new Options()).GetBoundingBox().Max.X, 4) == Math.Round(MIN_X, 4))
                {
                    point_2 = new XYZ(wall.get_Geometry(new Options()).GetBoundingBox().Min.X,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Max.Y,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Max.Z);
                }
                //Добавили ширину снизу
                if (Math.Round(wall.get_Geometry(new Options()).GetBoundingBox().Max.Y, 4) == Math.Round(MIN_Y, 4))
                {
                    point_3 = new XYZ(wall.get_Geometry(new Options()).GetBoundingBox().Min.X,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Min.Y,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Max.Z);
                }
                //Добавили ширину справа
                if (Math.Round(wall.get_Geometry(new Options()).GetBoundingBox().Min.X, 4) == Math.Round(MAX_X, 4))
                {
                    point_4 = new XYZ(wall.get_Geometry(new Options()).GetBoundingBox().Max.X,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Max.Y,
                                      wall.get_Geometry(new Options()).GetBoundingBox().Max.Z);
                }
            }

            /*            if (point_1 != null && point_2 != null && point_3 != null && point_4 != null)
                        {
                            //point 1
                            points.Add(new XYZ(point_4.X, point_1.Y, point_1.Z));
                            //point 2
                            points.Add(new XYZ(point_2.X, point_1.Y, point_1.Z));
                            //point 3
                            points.Add(new XYZ(point_2.X, point_3.Y, point_1.Z));
                            //point 4
                            points.Add(new XYZ(point_4.X, point_3.Y, point_1.Z));
                        }*/

            points.Add(new XYZ(0, 0, 0));
            points.Add(new XYZ(0, 20, 0));
            points.Add(new XYZ(20, 20, 0));
            points.Add(new XYZ(20, 10, 0));
            points.Add(new XYZ(40, 10, 0));
            points.Add(new XYZ(40, 0, 0));

            // Создание объекта цепочки кривых
            CurveLoop loop = new CurveLoop();

/*            loop.Append(Line.CreateBound(points[0], points[1]));
            loop.Append(Line.CreateBound(points[1], points[2]));
            loop.Append(Line.CreateBound(points[2], points[3]));
            loop.Append(Line.CreateBound(points[3], points[4]));
            loop.Append(Line.CreateBound(points[4], points[5]));
            loop.Append(Line.CreateBound(points[5], points[0]));*/
к
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