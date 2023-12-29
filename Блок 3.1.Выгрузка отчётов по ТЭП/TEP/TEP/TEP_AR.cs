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
            List<Element> elements = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ");
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
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", 1, false), doc).ToString();
            FillCell(15, 2, value, path);
        }
        private void B16(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ"), doc).ToString();
            FillCell(16, 2, value, path);
        }
        private void B17(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КМ"), doc).ToString();
            FillCell(17, 2, value, path);
        }
        private void B18(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", 1, true), doc).ToString();
            FillCell(18, 2, value, path);
        }
        private void B19(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ"), doc)
                          + Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КМ"), doc) 
                          + Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КХ"), doc)
                          + Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "АП"), doc)).ToString();
            FillCell(19, 2, value, path);
        }
        private void B20(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ"), doc).ToString();
            FillCell(20, 2, value, path);
        }
        private void B21(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "С ОТДЕЛКОЙ"), doc).ToString();
            FillCell(21, 2, value, path);
        }
        private void B22(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "WHITEBOX"), doc).ToString();
            FillCell(22, 2, value, path);
        }
        private void B23(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "БЕЗ ОТДЕЛКИ"), doc).ToString();
            FillCell(23, 2, value, path);
        }
        private void B24(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(24, 2, value, path);
        }
        private void B25(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(25, 2, value, path);
        }
        private void B26(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(26, 2, value, path);
        }
        private void B27(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(27, 2, value, path);
        }
        private void B28(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Балкон"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B29(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Балкон"), doc) * 0.3).ToString();
            FillCell(27, 2, value, path);
        }
        private void B30(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Лоджия"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B31(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Лоджия"), doc) * 0.5).ToString();
            FillCell(27, 2, value, path);
        }
        private void B32(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Терраса"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B33(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Терраса"), doc) * 0.3).ToString();
            FillCell(27, 2, value, path);
        }
        private void B34(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Кладовая", "PNG_Функция помещения", "Помещения кладовых"), doc).ToString();
            FillCell(34, 2, value, path);
        }
        private void B35(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(35, 2, value, path);
        }
        private void B36(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(36, 2, value, path);
        }
        private void B37(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "КМ"), doc).ToString();
            FillCell(37, 2, value, path);
        }
        private void B38(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(38, 2, value, path);
        }
        private void B39(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(39, 2, value, path);
        }
        private void B40(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(40, 2, value, path);
        }
        private void B41(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(41, 2, value, path);
        }
        private void B42(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(42, 2, value, path);
        }
        private void B43(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(43, 2, value, path);
        }
        private void B44(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNG_Имя помещения", "Детский сад"), doc).ToString();
            FillCell(44, 2, value, path);
        }
        private void B45(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(45, 2, value, path);
        }
        private void B46(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "МОП"), doc).ToString();
            FillCell(46, 2, value, path);
        }
        private void B47(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "МОП", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(47, 2, value, path);
        }
        private void B48(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Вестибюль", "ADSK_Этаж", "1"), doc).ToString(); 
            FillCell(48, 2, value, path);
        }
        private void B49(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лифтовой холл", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(49, 2, value, path);
        }
        private void B50(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Почтовая комната", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(50, 2, value, path);
        }
        private void B51(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(51, 2, value, path);
        }
        private void B52(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(52, 2, value, path);
        }
        private void B53(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(53, 2, value, path);
        }
        private void B54(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(54, 2, value, path);
        }
        private void B55(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(55, 2, value, path);
        }
        private void B56(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(56, 2, value, path);
        }
        private void B57(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(57, 2, value, path);
        }
        private void B58(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(58, 2, value, path);
        }
        private void B59(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(59, 2, value, path);
        }
        private void B60(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(60, 2, value, path);
        }
        private void B61(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(61, 2, value, path);
        }
        private void B62(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(62, 2, value, path);
        }
        private void B63(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(63, 2, value, path);
        }
        private void B64(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(64, 2, value, path);
        }
        private void B65(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(65, 2, value, path);
        }
        private void B66(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(66, 2, value, path);
        }
        private void B67(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(67, 2, value, path);
        }
        private void B68(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(68, 2, value, path);
        }
        private void B69(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(69, 2, value, path);
        }
        private void B70(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(70, 2, value, path);
        }
        private void B71(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(71, 2, value, path);
        }
        private void B72(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(72, 2, value, path);
        }
        private void B73(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(73, 2, value, path);
        }
        private void B74(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(74, 2, value, path);
        }
        private void B75(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(75, 2, value, path);
        }
        private void B76(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(76, 2, value, path);
        }
        private void B77(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(77, 2, value, path);
        }
        private void B78(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(78, 2, value, path);
        }
        private void B79(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(79, 2, value, path);
        }
        private void B80(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(80, 2, value, path);
        }
        private void B81(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(81, 2, value, path);
        }
        private void B82(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(82, 2, value, path);
        }
        private void B83(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(83, 2, value, path);
        }
        private void B84(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(84, 2, value, path);
        }
        private void B85(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(85, 2, value, path);
        }
        private void B86(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(86, 2, value, path);
        }
        private void B87(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(87, 2, value, path);
        }
        private void B88(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(88, 2, value, path);
        }
        private void B89(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(89, 2, value, path);
        }
        private void B90(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(90, 2, value, path);
        }
        private void B91(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(91, 2, value, path);
        }
        private void B92(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(92, 2, value, path);
        }
        private void B93(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(93, 2, value, path);
        }
        private void B94(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(94, 2, value, path);
        }
        private void B95(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(95, 2, value, path);
        }
    }
}
