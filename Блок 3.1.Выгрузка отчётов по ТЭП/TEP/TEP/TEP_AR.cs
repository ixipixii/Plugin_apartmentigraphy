using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using MathNet.Numerics;
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
            //Дописать список всех мопов в мини отчёте
            //и Количество помещений ритейла в коммерции

            //CopyFile("ТЭП_АР");
            String path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ТЭП_АР.xlsx");

            //Создаём отчёт для квартир
            var CellB6 = B6(doc, path);
            var CellB7 = B7(doc, path);
            B8(doc, path, CellB6, CellB7);
            B9(doc, path);
            B10(doc, path);
            B11(doc, path);
            B12(doc, path);
            B13(doc, path);
            B14(doc, path);
            B15(doc, path);
            B16(doc, path);
            B17(doc, path);
            B18(doc, path);
            B19(doc, path);
            string CountApart = B20(doc, path);
            string CountApartWithFinishing = B21(doc, path);
            string CountApartWhiteBox = B22(doc, path);
            string CountApartWithoutFinishing = B23(doc, path);
            B24(doc, path, CountApart);
            B25(doc, path, CountApartWithFinishing);
            B26(doc, path, CountApartWhiteBox);
            B27(doc, path, CountApartWithoutFinishing);
            string value = string.Empty;
            B28(doc, path, out value);
            B29(doc, path, value);
            B30(doc, path, out value);
            B31(doc, path, value);
            B32(doc, path, out value);
            B33(doc, path, value);
            string CountPantry = B34(doc, path);
            B35(doc, path, CountPantry);
            B36(doc, path, CountPantry, CountApart);

            //Создаём отчёт по коммерции
            int number = Commerce(doc, path, 37);

            //Создаём отчёт по МОП
            number = МОП(doc, path, number);

            //Создаём отчёт по машиноместам
            number = Auto(doc, path, number, int.Parse(CountApart));

            //Создаём отчёт по всем оставшимся категорям
            number = Engineer_infinity(doc, path, number);

            //Создаём отчёт по типовым этажам
            TypeFloor(doc, path, number.);

            System.Diagnostics.Process.Start(path);
        }
        #region всё для квартир
        private string CountApart(List<Element> list)
        {
            //Определяем количество квартир
            List<string> valuesADSK_KV = Values("ADSK_Номер квартиры", list);
            var listADSK_KV = valuesADSK_KV.Distinct().ToList();

            //Лист, где будут хранится все скции и квартиры
            List<List<string>> section_and_number = new List<List<string>>();
            string section = string.Empty;
            foreach (var kv in listADSK_KV)
            {
                var number = kv.Substring(kv.Length - 3).TrimStart('0');
                section = kv.Substring(kv.Length - 5, 2).TrimStart('0');
                section_and_number.Add(new List<string> { section, number });
            }

            //Словарь для хранения максимальных номеров каждой секции
            Dictionary<string, int> maxNumbers = new Dictionary<string, int>();

            foreach (var sect in section_and_number)
            {
                string sectionName = sect[0];
                int num = int.Parse(sect[1]);

                //Обновляем максимальный номер для текущей секции
                if (maxNumbers.ContainsKey(sectionName))
                {
                    maxNumbers[sectionName] = Math.Max(maxNumbers[sectionName], num);
                }
                else
                {
                    maxNumbers[sectionName] = num;
                }
            }

            //Создаём новый список с одним элементом с максимальным порядковым номером для каждой секции
            List<List<string>> result = new List<List<string>>();
            foreach (var kvp in maxNumbers)
            {
                result.Add(new List<string> { kvp.Key, kvp.Value.ToString() });
            }

            return result.Select(s => s.Skip(1).Sum(i => int.Parse(i))).Sum().ToString();
        }
        private string B6(Document doc, String path)
        {
            List<Element> elements = Elements(BuiltInCategory.OST_Rooms, doc);
            List<String> values = Values("ADSK_Этаж", elements);
            string value = values.Distinct().Count().ToString();
            FillCell(6, 2, value, path);
            return value;
        }
        private string B7(Document doc, String path)
        {
            List<Element> elements = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ");
            List<String> values = Values("ADSK_Этаж", elements);
            string value = values.Distinct().Count().ToString();
            FillCell(7, 2, value, path);
            return value;
        }
        private void B8(Document doc, String path, string CellB6, string CellB7)
        {
            /*List<Element> elements_all = Elements(BuiltInCategory.OST_Rooms, doc);
            List<Element> elements_live = Elements(BuiltInCategory.OST_Rooms, doc, "Назначение", "Квартира");
            List<String> values_all = Values("ADSK_Этаж", elements_all);
            List<String> values_live = Values("ADSK_Этаж", elements_live);
            string value = (values_all.Distinct().Count() - values_live.Distinct().Count()).ToString();*/
            var value = int.Parse(CellB6) - int.Parse(CellB7);
            FillCell(8, 2, value.ToString(), path);
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
            foreach (var room in list)
            {
                floors.Add(room.LookupParameter("ADSK_Этаж").AsString());
            }
            List<String> floor = floors.Distinct().ToList();
            double area = 0.0;
            foreach (var f in floor)
            {
                area += Areas(ElementsEquals(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", f), doc);
            }
            FillCell(16, 2, area.ToString(), path);
        }
        private void B17(Document doc, String path)
        {
            string value = Areas(ElementsEquals(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", "1"), doc).ToString();
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
                          + Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "ХК"), doc)
                          + Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "АП"), doc)).ToString();
            FillCell(19, 2, value, path);
        }
        private string B20(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ");
            string value = Areas(list, doc).ToString();
            FillCell(20, 2, value, path);
            return CountApart(list);
        }
        private string B21(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "С ОТДЕЛКОЙ");
            string value = Areas(list, doc).ToString();
            FillCell(21, 2, value, path);
            return CountApart(list);
        }
        private string B22(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "WHITEBOX");
            string value = Areas(list, doc).ToString();
            FillCell(22, 2, value, path);
            return CountApart(list);
        }
        private string B23(Document doc, String path)
        {
            var list = Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Позиция отделки", "БЕЗ ОТДЕЛКИ");
            string value = Areas(list, doc).ToString();
            FillCell(23, 2, value, path);
            return CountApart(list);
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
        private void B28(Document doc, String path, out string value)
        {
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Балкон"), doc).ToString();
            FillCell(28, 2, value, path);
        }
        private void B29(Document doc, String path, string value)
        {
            FillCell(29, 2, (double.Parse(value) * 0.3).ToString(), path);
        }
        private void B30(Document doc, String path, out string value)
        {
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Лоджия"), doc).ToString();
            FillCell(30, 2, value, path);
        }
        private void B31(Document doc, String path, string value)
        {
            FillCell(31, 2, (double.Parse(value) * 0.5).ToString(), path);
        }
        private void B32(Document doc, String path, out string value)
        {
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Терраса"), doc).ToString();
            FillCell(32, 2, value, path);
        }
        private void B33(Document doc, String path, string value)
        {
            FillCell(33, 2, (double.Parse(value) * 0.3).ToString(), path);
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
            string value = (double.Parse(countKl) / double.Parse(countKV)).ToString();
            FillCell(36, 2, value, path);
        }

        #endregion 
        private int Auto(Document doc, String path, int number, int countApart)
        {
            double areasFloorP = 0;
            var list = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_Parking)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            string value = string.Empty;
            List<int> numbers = new List<int>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            foreach (var v in list)
            {
                if (v != null)
                {
                    if (v.Name == "Обычные малые")
                        numbers[0]++;
                    if (v.Name == "Обычные средние")
                        numbers[1]++;
                    if (v.Name == "Обычные большие")
                        numbers[2]++;
                    if (v.Name == "Зависимые малые")
                        numbers[3]++;
                    if (v.Name == "Зависимые средние")
                        numbers[4]++;
                    if (v.Name == "Зависимые большие")
                        numbers[5]++;
                    if (v.Name == "С зарядкой малые")
                        numbers[6]++;
                    if (v.Name == "С зарядкой средние")
                        numbers[7]++;
                    if (v.Name == "С зарядкой большие")
                        numbers[8]++;
                    if (v.Name == "Гостевые малые")
                        numbers[9]++;
                    if (v.Name == "Гостевые средние")
                        numbers[10]++;
                    if (v.Name == "Гостевые большие")
                        numbers[11]++;
                    if (v.Name == "Мотоместа")
                        numbers[12]++;
                }
            }
            List<string> nameCells = new List<string>
            {
                "обычных малых",
                "обычных средних",
                "обычных больших",
                "зависимых малых",
                "зависимых средних",
                "зависимых больших",
                "с зарядкой малых",
                "с зарядкой средних",
                "с зарядкой больших",
                "гостевых малых",
                "гостевых средних",
                "гостевых больших",
                "мотоместа"
            };

            if(list.Count != 0)
                number = FillCellParameter(number, 2, list.Count.ToString(), path, "Количество машино/мото-мест в подземном паркинге, в том числе:", "шт.");

            for (var i = 0; i < nameCells.Count; i++)
            {
                if (numbers[i] != 0)
                    number = FillCellParameter(number, 2, numbers[i].ToString(), path, nameCells[i], "шт.");
            }

            if(list.Count != 0)
            {
                number = FillCellParameter(number, 2, (list.Count / countApart).ToString(), path, "коэффициент обеспеченности", "шт/кв");
                number = FillCellParameter(number, 2, (areasFloorP / list.Count).ToString(), path, "Отношение кол-ва мест к площади паркинга", "м2/место");
                number = FillCellParameter(number, 2, (areasFloorP / list.Count).ToString(), path, "Отношение кол-ва мест к ОП этажа", "м2/место");
            }

            return number;
        }
        private int Commerce(Document doc, String path, int number)
        {
            //Суммарная площадь коммерческих помещений, в том числе:
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Номер помещения", "КМ"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Суммарная площадь коммерческих помещений, в том числе:", "кв.м.");

            //Одно-уровневые коммерческие пом-я 
            string value_with_technology = (AreasLower(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Кафе"), doc, "5500")
                    + AreasLower(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Ресторан"), doc, "5500")
                    + AreasLower(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Супермаркет"), doc, "5500")
                    + AreasLower(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Фитнес"), doc, "5500")).ToString();

            string value_without_technology = (AreasLower(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения общественного назначения"), doc, "5500")).ToString();

            number = FillCellParameter(number, 2, (int.Parse(value_with_technology) + int.Parse(value_without_technology)).ToString(), path, "Одно-уровневые коммерческие пом-я:", "кв.м.");
            number = FillCellParameter(number, 2, value_with_technology, path, "Помещения с технологией", "кв.м.");
            number = FillCellParameter(number, 2, value_without_technology, path, "Помещения без технологии", "кв.м.");

            //Двух-уровневые коммерческие пом-я 
            value_with_technology = (AreasHeigher(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Кафе"), doc, "5500")
                    + AreasHeigher(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Ресторан"), doc, "5500")
                    + AreasHeigher(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Супермаркет"), doc, "5500")
                    + AreasHeigher(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Фитнес"), doc, "5500")).ToString();

            value_without_technology = (AreasHeigher(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения общественного назначения"), doc, "5500")).ToString();

            number = FillCellParameter(number, 2, (int.Parse(value_with_technology) + int.Parse(value_without_technology)).ToString(), path, "Двух-уровневые коммерческие пом-я:", "кв.м.");
            number = FillCellParameter(number, 2, value_with_technology, path, "Помещения с технологией", "кв.м.");
            number = FillCellParameter(number, 2, value_without_technology, path, "Помещения без технологии", "кв.м.");

            //Детский сад
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "ДОО"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Детский сад", "кв.м.");

            //Количество помещений ритейла
            //value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "ДОО"), doc).ToString();
            //number = FillCellParameter(number, 2, value, path, "Количество помещений ритейла", "кв.м.");

            return number;
        }
        private int МОП(Document doc, String path, int number)
        {
            //Суммарная площадь МОП, в том числе:
            string value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Номер помещения", "МОП"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Суммарная площадь МОП, в том числе:", "кв.м.");

            //ВСЕ МОПЫ
            using (var pack = new ExcelPackage(new FileInfo(path)))
            {
                //Помещения входной группы (1й этаж)
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("МОП входной группы 1 этажа", "МОП типовых этажей"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList);

                //МОП типовых этажей
                //Получаем лист с нужными комнатами по определённым категориям 
                roomList = RoomList(doc, Func("МОП типовых этажей", "МОП входной группы -1 этажа"));

                //Получаем список названий помещений с определёнными категориями
                roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения

                //Помещения входной группы паркинг (-1й этаж)
                //Получаем лист с нужными комнатами по определённым категориям 
                roomList = RoomList(doc, Func("МОП входной группы -1 этажа", "Помещения кладовых"));

                //Получаем список названий помещений с определёнными категориями
                roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }

            //Проход блока кладовых
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Проход блока кладовых"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Проход блока кладовых", "кв.м.");

            //Площадь лестниц и эвакуации (с -1го до последнего этажа)
            using (var pack = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("Лестницы эвакуации  (с -1го до последнего этажа)", "Помещения загрузки"));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }

            //Лестничная клетка (1го этажа)
            double value_1 = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "НЛК (наземная лестничная клетка)", "ADSK_Этаж", "1"), doc);
            value = value_1.ToString();
            number = FillCellParameter(number, 2, value, path, "Лестничная клетка (1го этажа)", "кв.м.");

            //Лестничная клетка типового этажа
            value = (Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "НЛК (наземная лестничная клетка)"), doc) - value_1).ToString();
            number = FillCellParameter(number, 2, value, path, "Лестничная клетка типового этажа", "кв.м.");

            //Лестничная клетка (-1го этажа)
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "ПЛК (подземная лестничная клетка)"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Лестничная клетка (-1го этажа)", "кв.м.");

            //Помещения загрузки, в том числе:
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения загрузки"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Помещения загрузки, в том числе:", "кв.м.");

            using (var pack = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];

                //ПОДЗЕМНЫЙ ЭТАЖ
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListUnder(doc, Func("Помещения загрузки", "Помещения мусороудаления"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList, "Подземный этаж");

                //ПЕРВЫЙ ЭТАЖ
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                roomList = RoomListEquals(doc, Func("Помещения загрузки", "Помещения мусороудаления"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList, "Первый этаж");

                // Сохраняем изменения
                pack.Save();

            }

            //Помещения мусороудаления, в том числе:
            //Помещения загрузки, в том числе:
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Функция помещения", "Помещения мусороудаления"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Помещения мусороудаления, в том числе:", "кв.м.");

            using (var pack = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];

                //ПОДЗЕМНЫЙ ЭТАЖ
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                List<Element> roomList = RoomListUnder(doc, Func("Помещения мусороудаления", "Инженерно-технические помещения"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList, "Подземный этаж");

                //Первый этаж
                //Получаем лист с нужными комнатами по определённым категориям в подземных этажах
                roomList = RoomListEquals(doc, Func("Помещения мусороудаления", "Инженерно-технические помещения"), "ADSK_Этаж", 1);

                //Получаем список названий помещений с определёнными категориями
                roomName = RoomNames(roomList);

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList, "Первый этаж");

                // Сохраняем изменения
                pack.Save();
            }

            //Въездная рампа
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Рампа"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Въездная рампа", "кв.м.");

            //Площадь подземного паркинга
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Имя помещения", "Помещение автостоянки"), doc).ToString();
            number = FillCellParameter(number, 2, value, path, "Помещение автостоянки", "кв.м.");

            return number;
        }
        private int Engineer_infinity(Document doc, String path, int number)
        {
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomList(doc, Func("Инженерно-технические помещения", ""));

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];

                //Если такие помещения есть
                if (roomList.Count != 0)
                    number = FillCellsArea(worksheet1, number, roomName, roomList);

                // Сохраняем изменения
                pack.Save();
            }
            return number;
        }
        private void TypeFloor(Document doc, String path, int number, string StartFloor, string EndFloor)
        {
            string value = string.Empty;

            //Типовой этаж
            value = (int.Parse(StartFloor) - int.Parse(EndFloor)).ToString();
            number = FillCellParameter(number, 2, value, path, "Типовой этаж " + StartFloor + " - " + EndFloor, "шт.");

            List<Element> roomListFloor = new List<Element>();
            //Количество квартир
            int start = int.Parse(StartFloor);
            int end = int.Parse(EndFloor);
            
            while (start == end)
            {
                roomListFloor.AddRange(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ", "ADSK_Этаж", start.ToString()));
                start++;
            }
            
            var nameKV = Values("ADSK_Номер квартиры", roomListFloor);
            var name = nameKV.Distinct().ToList();
            value = name.Count.ToString();
            number = FillCellParameter(number, 2, value, path, "Количество квартир", "шт.");

            //На этаж
            value = (int.Parse(value) / (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, value, path, "На этаж", "шт.");

            //Площадь квартир
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "ADSK_Номер квартиры", "КВ", "ADSK_Этаж", StartFloor), doc).ToString();
            string AreaApart = (int.Parse(value) * (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, AreaApart, path, "Площадь квартир", "кв.м.");

            //На этаж
            value = (int.Parse(AreaApart) / (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, value, path, "На этаж", "кв.м.");

            //Площадь МОП
            value = Areas(Elements(BuiltInCategory.OST_Rooms, doc, "PNR_Номер помещения", "МОП", "ADSK_Этаж", StartFloor), doc).ToString();
            value = (int.Parse(value) * (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, value, path, "Площадь МОП", "кв.м.");

            //Площадь МОП на этаж, в т.ч.:
            value = (int.Parse(value) / (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            FillCellParameter(number, 2, value, path, "Площадь МОП на этаж, в т.ч.:", "кв.м.");

            //Список всех мопов на 1 этаж
            using (var pack = new ExcelPackage(new System.IO.FileInfo(Path.Combine(path))))
            {
                //Получаем лист с нужными комнатами по определённым категориям 
                List<Element> roomList = RoomListFloor(doc, Func("Инженерно-технические помещения", ""), StartFloor);

                //Получаем список названий помещений с определёнными категориями
                List<String> roomName = RoomNames(roomList);

                //Тот лист, на котором будет выводиться отчёт
                ExcelWorksheet worksheet1 = pack.Workbook.Worksheets["Отчёт"];
                number = FillCellsArea(worksheet1, number, roomName, roomList);
                // Сохраняем изменения
                pack.Save();
            }

            //Общая площадь
            value = Areas(ElementsEquals(BuiltInCategory.OST_Rooms, doc, "ADSK_Этаж", StartFloor), doc).ToString();
            value = (int.Parse(value) * (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, value, path, "Общая площадь", "кв.м.");

            //Общая площадь этажа
            string OP = (int.Parse(value) / (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, OP, path, "Общая площадь этажа", "кв.м.");

            //Суммарная поэтажная площадь в габаритах наружных стен
            value = Areas(Elements(BuiltInCategory.OST_Areas, doc, "ADSK_Этаж", StartFloor), doc, "ГНС").ToString();
            string SPP = (int.Parse(value) * (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, SPP, path, "Суммарная поэтажная площадь в габаритах наружных стен", "кв.м.");

            //СПП в ГНС этажа
            string SPP_Floor = (int.Parse(SPP) / (int.Parse(EndFloor) - int.Parse(StartFloor))).ToString();
            number = FillCellParameter(number, 2, SPP_Floor, path, "СПП в ГНС этажа", "кв.м.");

            //Отношение Площади квартир к СПП ГНС
            value = (int.Parse(AreaApart) / int.Parse(SPP)).ToString();
            number = FillCellParameter(number, 2, value, path, "Отношение Площади квартир к СПП ГНС", "кв.м.");

            //Отношение ОП к СПП ГНС
            value = (int.Parse(OP) / int.Parse(SPP)).ToString();
            number = FillCellParameter(number, 2, value, path, "Отношение ОП к СПП ГНС", "кв.м.");

            //Отношение Площади квартир к ОП
            value = (int.Parse(AreaApart) / int.Parse(OP)).ToString();
            number = FillCellParameter(number, 2, value, path, "Отношение Площади квартир к ОП", "кв.м.");
        }
    }
}
