using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using static OfficeOpenXml.ExcelErrorValue;

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

            /*B14(doc, path);
            B15(doc, path);
            B16(doc, path);
            B17(doc, path);
            B18(doc, path);
            B19(doc, path);*/
            B108_infinity(doc, path);

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
            string value = Areas(Elements(BuiltInCategory.OST_Areas, doc), doc, "ГНС").ToString();
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
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", 1, true), doc).ToString();
            FillCell(15, 2, value, path);
        }
        private void B16(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ");
            List<String> floors = new List<string>(); 
            foreach(var room in list)
            {
                floors.Add(room.LookupParameter("ADSK_Этаж").ToString());
            }
            List<String> floor = floors.Distinct().ToList();
            double area = 0.0;
            foreach(var f in floors)
            {
                area += Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", f), doc);
            }
            FillCell(16, 2, area.ToString(), path);
        }
        private void B17(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", "1"), doc).ToString();
            FillCell(17, 2, value, path);
        }
        private void B18(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", 1, false), doc).ToString();
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
        private string B20(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ");
            string value = Areas(list, doc).ToString();
            FillCell(20, 2, value, path);
            //Определяем количество квартир
            List<string> valuesADSK_KV = Values("ADSK_Номер квартиры", list);
            var listADSK_KV = valuesADSK_KV.Distinct().ToList();
            int max = 0;
            foreach (var kv in listADSK_KV)
            {
                var maxKV = kv.Substring(kv.Length - 3).TrimStart('0');
                if (maxKV != null)
                {
                    if(int.Parse(maxKV) > max)
                    {
                        max = int.Parse(maxKV);
                    }
                }
            }
            return max.ToString();
        }
        private string B21(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "С ОТДЕЛКОЙ");
            string value = Areas(list, doc).ToString();
            FillCell(21, 2, value, path);
            //Определяем количество квартир
            List<string> valuesADSK_KV = Values("ADSK_Номер квартиры", list);
            var listADSK_KV = valuesADSK_KV.Distinct().ToList();
            int max = 0;
            foreach (var kv in listADSK_KV)
            {
                var maxKV = kv.Substring(kv.Length - 3).TrimStart('0');
                if (maxKV != null)
                {
                    if (int.Parse(maxKV) > max)
                    {
                        max = int.Parse(maxKV);
                    }
                }
            }
            return max.ToString();
        }
        private string B22(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "WHITEBOX");
            string value = Areas(list, doc).ToString();
            FillCell(22, 2, value, path);
            //Определяем количество квартир
            List<string> valuesADSK_KV = Values("ADSK_Номер квартиры", list);
            var listADSK_KV = valuesADSK_KV.Distinct().ToList();
            int max = 0;
            foreach (var kv in listADSK_KV)
            {
                var maxKV = kv.Substring(kv.Length - 3).TrimStart('0');
                if (maxKV != null)
                {
                    if (int.Parse(maxKV) > max)
                    {
                        max = int.Parse(maxKV);
                    }
                }
            }
            return max.ToString();
        }
        private string B23(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "БЕЗ ОТДЕЛКИ");
            string value = Areas(list, doc).ToString();
            FillCell(23, 2, value, path);
            //Определяем количество квартир
            List<string> valuesADSK_KV = Values("ADSK_Номер квартиры", list);
            var listADSK_KV = valuesADSK_KV.Distinct().ToList();
            int max = 0;
            foreach (var kv in listADSK_KV)
            {
                var maxKV = kv.Substring(kv.Length - 3).TrimStart('0');
                if (maxKV != null)
                {
                    if (int.Parse(maxKV) > max)
                    {
                        max = int.Parse(maxKV);
                    }
                }
            }
            return max.ToString();
        }
        private string B24(Document doc, String path, string count)
        {
            string value = count;
            FillCell(24, 2, value, path);
            return value;
        }
        private void B25(Document doc, String path, string count)
        {
            string value = count;
            FillCell(25, 2, value, path);
        }
        private void B26(Document doc, String path, string count)
        {
            string value = count;
            FillCell(26, 2, value, path);
        }
        private void B27(Document doc, String path, string count)
        {
            string value = count;
            FillCell(27, 2, value, path);
        }
        private void B28(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Балкон"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B29(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Балкон"), doc) * 0.3).ToString();
            FillCell(27, 2, value, path);
        }
        private void B30(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лоджия"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B31(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лоджия"), doc) * 0.5).ToString();
            FillCell(27, 2, value, path);
        }
        private void B32(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Терраса"), doc).ToString();
            FillCell(27, 2, value, path);
        }
        private void B33(Document doc, String path)
        {
            string value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Терраса"), doc) * 0.3).ToString();
            FillCell(27, 2, value, path);
        }
        private string B34(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Кладовая", "PNR_Функция помещения", "Помещения кладовых");
            string value = Areas(list, doc).ToString();
            FillCell(34, 2, value, path);
            return list.Count.ToString();
        }
        private string B35(Document doc, String path, string count)
        {
            string value = count;
            FillCell(35, 2, value, path);
            return value;
        }
        private void B36(Document doc, String path, string countKl, string countKV)
        {
            string value = (double.Parse(countKl)/double.Parse(countKV)).ToString();
            FillCell(36, 2, value, path);
        }

        //КМ
        /*private void B37(Document doc, String path)
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
        }*/
        /*private void B46(Document doc, String path)
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
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Санузел", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(51, 2, value, path);
        }
        private void B52(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Инвентарная (помещение хранения детского оборудования)", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(52, 2, value, path);
        }
        private void B53(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Тамбур", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(53, 2, value, path);
        }
        private void B54(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей"), doc).ToString();
            FillCell(54, 2, value, path);
        }
        private void B55(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "Лифтовой холл (ПБЗ)"), doc).ToString();
            FillCell(55, 2, value, path);
        }
        private void B56(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "Коридор"), doc).ToString();
            FillCell(56, 2, value, path);
        }
        private void B57(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "ПУИ"), doc).ToString();
            FillCell(57, 2, value, path);
        }
        private void B58(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "Лестничная клетка/ПБЗ"), doc).ToString();
            FillCell(58, 2, value, path);
        }
        private void B59(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "Сервисный коридор"), doc).ToString();
            FillCell(59, 2, value, path);
        }
        private void B60(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "МОП типовых этажей", "PNR_Имя помещения", "Помещение ревизии инженерных коммуникаций"), doc).ToString();
            FillCell(60, 2, value, path);
        }
        private void B61(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Паркинг", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(61, 2, value, path);
        }
        private void B62(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лифтовой холл (тамбур-шлюз)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(62, 2, value, path);
        }
        private void B63(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Инвентарная(помещение хранения велосипедов)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(63, 2, value, path);
        }
        private void B64(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Коридор", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(64, 2, value, path);
        }
        private void B65(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Проход блока кладовых", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(65, 2, value, path);
        }
        private void B66(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лестничная клетка"), doc).ToString();
            FillCell(66, 2, value, path);
        }
        private void B67(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лестничная клетка", "PNR_Функция помещения", "МОП типовых этажей"), doc).ToString();
            FillCell(67, 2, value, path);
        }
        private void B68(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лестничная клетка", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(68, 2, value, path);
        }
        private void B69(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лестничная клетка", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(69, 2, value, path);
        }
        private void B70(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения загрузки"), doc).ToString();
            FillCell(70, 2, value, path);
        }
        private void B71(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения загрузки", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(71, 2, value, path);
        }
        private void B72(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Зона / коридор загрузки ", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(72, 2, value, path);
        }
        private void B73(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение временного хранения", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(73, 2, value, path);
        }
        private void B74(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лифтовой холл сервисного лифта", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(74, 2, value, path);
        }
        private void B75(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения загрузки", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(75, 2, value, path);
        }
        private void B76(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Зона / коридор загрузки ", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(76, 2, value, path);
        }
        private void B77(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение временного хранения", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(77, 2, value, path);
        }
        private void B78(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лифтовой холл сервисного лифта", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(78, 2, value, path);
        }
        private void B79(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения мусороудаления"), doc).ToString();
            FillCell(79, 2, value, path);
        }
        private void B80(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения мусороудаления", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(80, 2, value, path);
        }
        private void B81(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Мусорокамера хранения ТБО", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(81, 2, value, path);
        }
        private void B82(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение временного хранения", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(82, 2, value, path);
        }
        private void B83(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Коридор", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(83, 2, value, path);
        }
        private void B84(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения мусороудаления", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(84, 2, value, path);
        }
        private void B85(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Мусорокамера хранения ТБО", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(85, 2, value, path);
        }
        private void B86(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение временного хранения", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(86, 2, value, path);
        }
        private void B87(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Коридор", "ADSK_Этаж", "1"), doc).ToString();
            FillCell(87, 2, value, path);
        }
        private void B88(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Рампа"), doc).ToString();
            FillCell(88, 2, value, path);
        }
        private void B89(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Паркинг", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(89, 2, value, path);
        }*/
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
        private void B96(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(96, 2, value, path);
        }
        private void B97(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(97, 2, value, path);
        }
        private void B98(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(98, 2, value, path);
        }
        private void B99(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(99, 2, value, path);
        }
        private void B100(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(100, 2, value, path);
        }
        private void B101(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(101, 2, value, path);
        }
        private void B102(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(102, 2, value, path);
        }
        private void B103(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(103, 2, value, path);
        }
        private void B104(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(104, 2, value, path);
        }
        private void B105(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(105, 2, value, path);
        }
        private void B106(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(106, 2, value, path);
        }
        private void B107(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(107, 2, value, path);
        }
        /*private void B108(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Инженерно-технические помещения", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(108, 2, value, path);
        }
        private void B109(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение кабельного ввода", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(109, 2, value, path);
        }
        private void B110(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение венткамеры ОВ (принадлежность обслуживаемой зоны: автостоянка,ТП, ДОО, холодильного центра, ИТП...)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(110, 2, value, path);
        }
        private void B111(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение венткамеры ПДВ (принадлежность обслуживаемой зоны: автостоянки, ТП, ДОО, холодильного центра, ИТП...)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(111, 2, value, path);
        }
        private void B112(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Кроссовая (узел, связи, СС, А/с, АР1, помещение домофонной сети и пр.)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(112, 2, value, path);
        }
        private void B113(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение ввода СС", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(113, 2, value, path);
        }
        private void B114(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ВРУ (принадлежность обслуживаемой зоны: ТП, ДОО, холодильного центра, ИТП...)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(114, 2, value, path);
        }
        private void B115(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ВРУ жилья", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(115, 2, value, path);
        }
        private void B116(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ВРУ автостоянки", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(116, 2, value, path);
        }
        private void B117(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ВРУ ПОН", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(117, 2, value, path);
        }
        private void B118(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ТП", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(118, 2, value, path);
        }
        private void B119(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Трансформаторная камера", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(119, 2, value, path);
        }
        private void B120(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "РУНН", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(120, 2, value, path);
        }
        private void B121(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "РУВН", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(121, 2, value, path);
        }
        private void B122(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Холодильный центр", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(122, 2, value, path);
        }
        private void B123(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Водомерный узел", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(123, 2, value, path);
        }
        private void B124(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Серверная", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(124, 2, value, path);
        }
        private void B125(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение выпуска (ВК, ливневка и т.д)", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(125, 2, value, path);
        }
        private void B126(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ИТП", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(126, 2, value, path);
        }
        private void B127(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Насосная", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(127, 2, value, path);
        }
        private void B128(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Насосная АУПТ", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(128, 2, value, path);
        }
        private void B129(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение очистки воды", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(129, 2, value, path);
        }
        private void B130(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "КНС", "ADSK_Этаж", "-1"), doc).ToString();
            FillCell(130, 2, value, path);
        }
        private void B131(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(131, 2, value, path);
        }
        private void B132(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(132, 2, value, path);
        }
        private void B133(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(133, 2, value, path);
        }
        private void B134(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Тамбур", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(134, 2, value, path);
        }
        private void B135(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Вестибюль", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(135, 2, value, path);
        }
        private void B136(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Переговорная комната для приема населения", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(136, 2, value, path);
        }
        private void B137(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Переговорная комната рабочая", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(137, 2, value, path);
        }
        private void B138(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Рабочее место менеджера по работе с клиентами", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(138, 2, value, path);
        }
        private void B139(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Санузел", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(139, 2, value, path);
        }
        private void B140(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Коридор", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(140, 2, value, path);
        }
        private void B141(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(141, 2, value, path);
        }
        private void B142(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Кабинет управляющего", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(142, 2, value, path);
        }
        private void B143(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Зал менеджеров ", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(143, 2, value, path);
        }
        private void B144(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Комната для отдыха и приема пищи ", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(144, 2, value, path);
        }
        private void B145(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(14, 2, value, path);
        }
        private void B146(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Серверная", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(146, 2, value, path);
        }
        private void B147(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Архив", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(147, 2, value, path);
        }
        private void B148(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Коридор", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(148, 2, value, path);
        }
        private void B149(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Объединенный диспетчерский пункт"), doc).ToString();
            FillCell(149, 2, value, path);
        }
        private void B150(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Объединенный диспетчерский пункт", "PNR_Функция помещения", "Объединенный диспетчерский пункт"), doc).ToString();
            FillCell(150, 2, value, path);
        }
        private void B151(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Тамбур", "PNR_Функция помещения", "Помещение управляющей компании"), doc).ToString();
            FillCell(151, 2, value, path);
        }
        private void B152(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Серверная", "PNR_Функция помещения", "Объединенный диспетчерский пункт"), doc).ToString();
            FillCell(152, 2, value, path);
        }
        private void B153(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "С/у", "PNR_Функция помещения", "Объединенный диспетчерский пункт"), doc).ToString();
            FillCell(153, 2, value, path);
        }
        private void B154(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Объединенный диспетчерский пункт"), doc).ToString();
            FillCell(154, 2, value, path);
        }
        private void B155(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения линейного и обслуживающего персонала"), doc).ToString();
            FillCell(155, 2, value, path);
        }
        private void B156(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString();
            FillCell(156, 2, value, path);
        }
        private void B157(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString();
            FillCell(157, 2, value, path);
        }
        private void B158(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString(); 
            FillCell(158, 2, value, path);
        }
        private void B159(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString(); FillCell(159, 2, value, path);
        }
        private void B160(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString(); FillCell(160, 2, value, path);
        }
        private void B161(Document doc, String path)
        {
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПУИ", "PNR_Функция помещения", "Помещение для размещения круглосуточного дежурного техника"), doc).ToString(); FillCell(161, 2, value, path);
        }
        private void B162(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(162, 2, value, path);
        }
        private void B163(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(163, 2, value, path);
        }
        private void B164(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(164, 2, value, path);
        }
        private void B165(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(165, 2, value, path);
        }
        private void B166(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(166, 2, value, path);
        }
        private void B167(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(167, 2, value, path);
        }
        private void B168(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(168, 2, value, path);
        }
        private void B169(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(169, 2, value, path);
        }
        private void B170(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(170, 2, value, path);
        }
        private void B171(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(171, 2, value, path);
        }
        private void B172(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(172, 2, value, path);
        }
        private void B173(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(173, 2, value, path);
        }
        private void B174(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(174, 2, value, path);
        }
        private void B175(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(175, 2, value, path);
        }
        private void B176(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(176, 2, value, path);
        }
        private void B177(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(177, 2, value, path);
        }
        private void B178(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(178, 2, value, path);
        }
        private void B179(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(179, 2, value, path);
        }
        private void B180(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(180, 2, value, path);
        }
        private void B181(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(181, 2, value, path);
        }
        private void B182(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(182, 2, value, path);
        }
        private void B183(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(183, 2, value, path);
        }
        private void B184(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(184, 2, value, path);
        }
        private void B185(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(185, 2, value, path);
        }
        private void B186(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(186, 2, value, path);
        }
        private void B187(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(187, 2, value, path);
        }
        private void B188(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(188, 2, value, path);
        }
        private void B189(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(189, 2, value, path);
        }
        private void B190(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(190, 2, value, path);
        }
        private void B191(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(191, 2, value, path);
        }
        private void B192(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(192, 2, value, path);
        }
        private void B193(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(193, 2, value, path);
        }
        private void B194(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(194, 2, value, path);
        }
        private void B195(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(195, 2, value, path);
        }
        private void B196(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(196, 2, value, path);
        }
        private void B197(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(197, 2, value, path);
        }
        private void B198(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(198, 2, value, path);
        }
        private void B199(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(199, 2, value, path);
        }
        private void B200(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(200, 2, value, path);
        }
        private void B201(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(201, 2, value, path);
        }
        private void B202(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(202, 2, value, path);
        }
        private void B203(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(203, 2, value, path);
        }
        private void B204(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(204, 2, value, path);
        }
        private void B205(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(205, 2, value, path);
        }
        private void B206(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(206, 2, value, path);
        }
        private void B207(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(207, 2, value, path);
        }
        private void B208(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(208, 2, value, path);
        }
        private void B209(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(209, 2, value, path);
        }
        private void B210(Document doc, String path)
        {
            string value = string.Empty;
            FillCell(210, 2, value, path);
        }*/
        private int МОП(Document doc, String path)
        {
            //Суммарная площадь МОП, в том числе:
            int number = 46;
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "МОП"), doc).ToString();
            FillCellParameter(number, 2, value, path, "Суммарная площадь МОП, в том числе:", "кв.м.");

            //Помещения входной группы (1й этаж)
            number++;
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("МОП входной группы 1 этажа", "МОП типовых этажей"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);
                // Сохраняем изменения
                pack.Save();
            }

            //МОП типовых этажей
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("МОП типовых этажей", "МОП входной группы -1 этажа"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);
                // Сохраняем изменения
                pack.Save();
            }

            //Помещения входной группы паркинг (-1й этаж)
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("МОП входной группы -1 этажа", "Помещения кладовых"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);

                //Добавляем помещение хранения велосипедов
                value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Инвентарная(помещение хранения велосипедов)"), doc).ToString();
                if(value != string.Empty)
                {
                    FillCellParameter(number, 2, value, path, "Инвентарная(помещение хранения велосипедов)", "кв.м.");
                }
                number++;

                // Сохраняем изменения
                pack.Save();
            }
            
            //Проход блока кладовых
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Проход блока кладовых)"), doc).ToString();
            if (value != string.Empty)
            {
                FillCellParameter(number, 2, value, path, "Проход блока кладовых", "кв.м.");
            }
            number++;

            //Площадь лестниц и эвакуации (с -1го до последнего этажа)
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("Лестницы эвакуации  (с -1го до последнего этажа)", "Помещения загрузки"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }

            //Лестничная клетка (1го этажа)
            double value_1 = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "НЛК (наземная лестничная клетка)", "ADSK_Этаж", "1"), doc);
            value = value_1.ToString();
            if (value != string.Empty)
            {
                FillCell(number, 1, "Лестничная клетка (1го этажа)", path);
                FillCell(number, 2, value, path);
                FillCell(number, 3, "кв.м", path);
            }
            number++;

            //Лестничная клетка типового этажа
            value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "НЛК (наземная лестничная клетка)"), doc) - value_1).ToString();
            if (value != string.Empty)
            {
                FillCell(number, 1, "Лестничная клетка типового этажа", path);
                FillCell(number, 2, value, path);
                FillCell(number, 3, "кв.м", path);
            }
            number++;

            //Лестничная клетка (-1го этажа)
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПЛК (подземная лестничная клетка)"), doc).ToString();
            if (value != string.Empty)
            {
                FillCell(number, 1, "Проход блока кладовых", path);
                FillCell(number, 2, value, path);
                FillCell(number, 3, "кв.м", path);
            }
            number++;

            //Помещения загрузки, в том числе:
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomList(doc, Func("Помещения загрузки", "Помещения мусороудаления"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }

            //Подземный этаж
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListUnder(doc, Func("Помещения загрузки", "Помещения мусороудаления"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsAreaUnder(worksheet1, number, roomName, roomList, "Подземный этаж");

                // Сохраняем изменения
                pack.Save();
            }

            //Первый этаж
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListEquals(doc, Func("Помещения загрузки", "Помещения мусороудаления"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsAreaUnder(worksheet1, number, roomName, roomList, "Первый этаж");

                // Сохраняем изменения
                pack.Save();
            }

            //Помещения мусороудаления, в том числе:
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomList(doc, Func("Помещения мусороудаления", "Инженерно-технические помещения"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }

            //Подземный этаж
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListUnder(doc, Func("Помещения мусороудаления", "Инженерно-технические помещения"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsAreaUnder(worksheet1, number, roomName, roomList, "Подземный этаж");

                // Сохраняем изменения
                pack.Save();
            }

            //Первый этаж
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListEquals(doc, Func("Помещения мусороудаления", "Инженерно-технические помещения"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                number = FillCellsAreaUnder(worksheet1, number, roomName, roomList, "Первый этаж");

                // Сохраняем изменения
                pack.Save();
            }

            //Въездная рампа
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Рампа"), doc).ToString();
            if (value != string.Empty)
            {
                FillCell(number, 1, "Въездная рампа", path);
                FillCell(number, 2, value, path);
                FillCell(number, 3, "кв.м", path);
            }
            number++;

            //Площадь подземного паркинга
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение автостоянки"), doc).ToString();
            if (value != string.Empty)
            {
                FillCell(number, 1, "Помещение автостоянки", path);
                FillCell(number, 2, value, path);
                FillCell(number, 3, "кв.м", path);
            }
            number++;

            return number;
        }
        private void Engineer_infinity(Document doc, String path, int number)
        {
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("Инженерно-технические помещения", ""));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets[2];
                FillCellsArea(worksheet1, number, roomName, roomList);
                // Сохраняем изменения
                pack.Save();
            }
        }
        private void TypeFloor(Document doc, String path, int number, string StartFloor, string EndFloor)
        {
            string value = string.Empty;
            //Типовой этаж
            FillCellParameter(number, 2, value, path, "Типовой этаж " + StartFloor + " - " + EndFloor, "шт.");
            number++;
            //Количество квартир
            FillCell(number, 2, value, path);
            number++;
            //На этаж
            FillCell(number, 2, value, path);
            number++;
            //Площадь квартир
            FillCell(number, 2, value, path);
            number++;
            //На этаж
            FillCell(number, 2, value, path);
            number++;
            //Площадь МОП
            FillCell(number, 2, value, path);
            number++;
        }
    }
}
