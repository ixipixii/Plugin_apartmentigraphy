using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TEP
{
    internal class TEP_AR : Unloading
    {
        public TEP_AR(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            //CopyFile("ТЭП_АР");
            String path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ТЭП_АР.xlsx");
            //B6(doc, path);
            //B7(doc, path); 
            //B8(doc, path);
            B9(doc, path);
        }
        private void B6(Document doc, String path)
        {
            List<Element> elements = Elements(BuiltInCategory.OST_Rooms, doc);
            List<String> values = Values("ADSK_Этаж", elements);
            string value = values.Distinct().Count().ToString();
            FillCell(6, 2, value, path);
        }
        private void B7(Document doc, String path)
        {
            List<Element> elements = Elements(BuiltInCategory.OST_Rooms, doc, "Назначение", "Квартира");
            List<String> values = Values("ADSK_Этаж", elements);
            string value = values.Distinct().Count().ToString();
            FillCell(7, 2, value, path);
        }
        private void B8(Document doc, String path)
        {
            List<Element> elements_all = Elements(BuiltInCategory.OST_Rooms, doc);
            List<Element> elements_live = Elements(BuiltInCategory.OST_Rooms, doc, "Назначение", "Квартира");
            List<String> values_all = Values("ADSK_Этаж", elements_all);
            List<String> values_live = Values("ADSK_Этаж", elements_live);
            string value = (values_all.Distinct().Count() - values_live.Distinct().Count()).ToString();
            FillCell(8, 2, value, path);
        }
        private void B9(Document doc, String path) //Сравниваем ГНС с параметром "ADSK_Назначение вида"
        {
            string value = Area(Elements(BuiltInCategory.OST_Areas, doc ), doc, "ГНС").ToString();
            FillCell(9, 2, value, path);
        }
        private void B10(Document doc, String path)
        {
            String value = String.Empty;
            FillCell(10, 2, value, path);
        }
        private void B11(Document doc, String path)
        {
            String value = String.Empty;
            FillCell(11, 2, value, path);
        }
        private void B12(Document doc, String path)
        {
            String value = String.Empty;
            FillCell(12, 2, value, path);
        }
    }
}
