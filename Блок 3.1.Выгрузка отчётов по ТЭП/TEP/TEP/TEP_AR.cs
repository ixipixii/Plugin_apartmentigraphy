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
            //B9(doc, path);
            //B10(doc, path);
            //B11(doc, path);
            //B12(doc, path);
            //B13(doc, path);

            B14(doc, path);
            B15(doc, path);
            B16(doc, path);
            B17(doc, path);
            B18(doc, path);
            B19(doc, path);
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
        private void B9(Document doc, String path) //Сравниваем ГНС с параметром "PNR_Имя помещения"
        {
            string value = Areas(Elements(BuiltInCategory.OST_Areas, doc ), doc, "ГНС").ToString();
            FillCell(9, 2, value, path);
        }
        private void B10(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "ГНС")
                            - Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "нежилая")).ToString();
            FillCell(10, 2, value, path);
        }
        private void B11(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "ГНС")
                            - Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "нежилая")
                            - Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "входная")).ToString();
            FillCell(11, 2, value, path);
        }
        private void B12(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "входная").ToString();
            FillCell(12, 2, value, path);
        }
        private void B13(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "нежилая").ToString();
            FillCell(13, 2, value, path);
        }
        private void B14(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc), doc).ToString();
            FillCell(14, 2, value, path);
        }
        private void B15(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", ), doc).ToString();
            FillCell(15, 2, value, path);
        }
        private void B16(Document doc, String path)
        {
            string value = String.Empty;
            FillCell(16, 2, value, path);
        }
        private void B17(Document doc, String path)
        {
            string value = String.Empty;
            FillCell(17, 2, value, path);
        }
        private void B18(Document doc, String path)
        {
            string value = String.Empty;
            FillCell(18, 2, value, path);
        }
        private void B19(Document doc, String path)
        {
            string value = String.Empty;
            FillCell(19, 2, value, path);
        }
    }
}
