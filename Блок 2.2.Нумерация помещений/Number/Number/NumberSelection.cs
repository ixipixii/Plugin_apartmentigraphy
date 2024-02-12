using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using Document = Autodesk.Revit.DB.Document;

namespace Number
{
    internal class NumberSelection : INotifyPropertyChanged
    {
        UIApplication _uiapp;
        UIDocument _uidoc;
        Document _doc;

        public List<Group> AllGroupsApart = new List<Group>(); //Лист всех групп-помещений
        public DelegateCommand SelectCommandApart { get; } //Делегат кнопки "Пронумеровать" группы
        public DelegateCommand SelectCommandRoom { get; } //Делегат кнопки "Пронумеровать" помещения
        public DelegateCommand RenumberCommandApart { get; } //Делегат кнопки "Перенумеровать квартиру" помещения
        public DelegateCommand RenumberCommandRoom { get; } //Делегат кнопки "Перенумеровать помещение" помещения
        public DelegateCommand SelectApart { get; } //Делегат кнопки "Нумерация квартир по параметру группы"
        public DelegateCommand SelectApart_2 { get; } //Делегат кнопки "Нумерация квартир по параметру помещения"
        public DelegateCommand SelectRoom { get; } //Делегат кнопки "Нумерация помещений"
        public List<ClassApartAndRoom> SelectedApartList { get; set; } = new List<ClassApartAndRoom>(); //Вспомогательный лист для ListView для групп 
        public List<ClassApartAndRoom> SelectedRoomList { get; set; } = new List<ClassApartAndRoom>(); //Вспомогательный лист для ListView для помещений
        static public ClassApartAndRoom SelectedRoom { get; set; } //выбранное помещение на ListView
        public DelegateCommand EnterCommandRenumber { get; } //Делегат кнопки "Ввести"
        public List<Element> AllRoomsRenumberNew = new List<Element>(); //Новые квартиры для переименовки
        public String _SelectedSectionValue { get; set; } //Выбранная секция в ComboBox
        private string NewValueNumberApart; //
        public int index = 0;
        public string Text
        {
            get => NewValueNumberApart;
            set
            {
                NewValueNumberApart = value;
                //обрабатываешь здесь значение из элемента
                OnPropertyChanged();
            }
        }
        public NumberSelection(UIApplication uiapp, UIDocument uidoc, Document doc, string SelectedSectionValue, int v)
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            if (v == 1)
                SelectCommandApart = new DelegateCommand(SelectionApart);
            if (v == 2)
                SelectCommandApart = new DelegateCommand(SelectionApart);
            RenumberCommandApart = new DelegateCommand(RenumberApart);
            RenumberCommandRoom = new DelegateCommand(RenumberRoom);
            SelectCommandRoom = new DelegateCommand(SelectionRoom);
            SelectApart = new DelegateCommand(SelectAparts);
            SelectApart_2 = new DelegateCommand(SelectAparts_2);
            SelectRoom = new DelegateCommand(SelectRooms);
            _SelectedSectionValue = SelectedSectionValue;
            EnterCommandRenumber = new DelegateCommand(EnterRenumber);
        }

        private void SelectRooms()
        {
            RaiseCloseRequest();
            var room = new Room(_uiapp, _uidoc, _doc);
            room.ShowDialog();
        }
        private void SelectAparts()
        {
            RaiseCloseRequest();
            var sectionWindow = new SelectSection(_uiapp, _uidoc, _doc, 1);
            sectionWindow.ShowDialog();
        }
        public event EventHandler CloseRequest;
        public event PropertyChangedEventHandler PropertyChanged;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
        public void RoomList() //Метод, считывающий все помещения для ListView
        {
            //Фильтр по всем помещениям на данном виде
            var AllRooms = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .ToList();

            //Заносим все помещения на активном виде в ListView
            foreach (var room in AllRooms)
            {
                if (room.GroupId.ToString() == "-1")
                {
                    try
                    {
                        ClassApartAndRoom roomInList = new ClassApartAndRoom("0", room.LookupParameter("PNR_Номер помещения").AsString(), room.Name, room.Id);
                        SelectedRoomList.Add(roomInList);
                    }
                    catch
                    {
                        ClassApartAndRoom roomInList = new ClassApartAndRoom("0", "ПАРАМЕТР НЕ НАЙДЕН", room.Name, room.Id);
                        SelectedRoomList.Add(roomInList);
                    }

                }
            }
        }
        public bool OvverideContains(IList<Element> RoomListElement, Element group) //*Переопределение Contains для Element
        {
            foreach (var element in RoomListElement)
            {
                if (element.Name == group.Name)
                    return false;
            }
            return true;
        }
        public void SelectionRoom() //Основной метод действия с помещениями
        {
            RaiseCloseRequest();

            var roomFilter = new RoomPickFilter();

            //Фильтр по всем помещениям на данном виде
            var RoomList = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .ToList();

            IList<Element> RoomListElement = new List<Element>();

            //Добавляем помещения не в группе
            foreach (var roomElement in RoomList)
            {
                if (roomElement.GroupId.ToString() == "-1")
                    RoomListElement.Add(roomElement);
            }

            //Поверяем параметры
            if(CheckParameter(RoomListElement) == 1)
                return;



            //Нумеруем помещения
            NumberRoom(RoomListElement);

            //После завершения нумерации выводим список помещений
            var room = new Room(_uiapp, _uidoc, _doc);
            room.ShowDialog();
        }
        public void NumberRoom(IList<Element> RoomListElement) //Метод, нумерующий помещения
        {
            // Параметры помещений
            String PNR_Function = string.Empty;
            String PNR_Name = string.Empty;
            String PNR_Floor = string.Empty;
            String PNR_Building = string.Empty;

            string PNR_Funс = string.Empty;

            List<List<String>> funcCount = new List<List<string>>();
            List<Element> flagList = new List<Element>();
            List<String> funcList = new List<string>();
            foreach (var room in RoomListElement)
            {
                PNR_Function = room.LookupParameter("PNR_Функция помещения").AsString();
                PNR_Name = room.LookupParameter("PNR_Имя помещения").AsString();
                SetFunc(PNR_Function, PNR_Name, out PNR_Funс, out string flag);
                if(flag == "0" && PNR_Funс != "УК")
                {
                    funcList.Add(PNR_Funс);
                    flagList.Add(room);
                }
            }

            List<string> funcListDistinct = funcList.Distinct().ToList();

            for (int i = 0; i < funcListDistinct.Count; i++)
            {
                funcCount.Add(new List<String> { funcListDistinct[i], "0" });
            }

            foreach (var room in flagList)
            {
                PNR_Function = room.LookupParameter("PNR_Функция помещения").AsString();
                PNR_Name = room.LookupParameter("PNR_Имя помещения").AsString();
                PNR_Floor = room.LookupParameter("ADSK_Этаж").AsString();
                if (PNR_Floor.Substring(0, 1) == "-")
                {
                    PNR_Floor = "П" + PNR_Floor.Substring(1, PNR_Floor.Length - 1);
                }
                try
                {
                    PNR_Building = room.LookupParameter("ADSK_Номер здания").AsString();
                    if (PNR_Building == null)
                        PNR_Building = "";
                    else PNR_Building += "/";
                }
                catch { PNR_Building = ""; }

                SetFunc(PNR_Function, PNR_Name, out PNR_Funс, out string flag);

                Transaction tr = new Transaction(_doc, "Number room");
                tr.Start();

                foreach (var fc in funcCount)
                {
                    if (fc != null)
                    {
                        for (int i = 0; i < fc.Count; i++)
                        {
                            if (fc[i] == PNR_Funс)
                            {
                                string count = fc[1].ToString();
                                fc[1] = (int.Parse(count) + 1).ToString();
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}{PNR_Floor}.{PNR_Funс}.{fc[1]}");
                            }
                        }
                    }
                }
                tr.Commit();
            }
        }

        //version 2
        public void SelectionApart() //Основной метод действия с квартирами
        {
            RaiseCloseRequest();

            String PNR_Function = string.Empty; //Функция помещения
            String PNR_Name = string.Empty; //Имя помещения
            String PNR_Section = string.Empty; //Секция
            String PNR_Building = string.Empty; //Номер здания
            String PNR_Funс = string.Empty; //Сокращение функции
            String PNR_Floor = string.Empty; //Этаж

            var roomFilter = new RoomPickFilter();

            //Создаём листы: ссылки на элементы, лист помещений
            IList<Autodesk.Revit.DB.Reference> refrence = new List<Autodesk.Revit.DB.Reference>();
            IList<Element> ApartListElement = new List<Element>();

            var refElement = _uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите элементы");

            foreach (var element in refElement)
            {
                var apart = _doc.GetElement(element);
                ApartListElement.Add(apart);
            }

            //Проверяем все параметры модели
            if(CheckParameter(ApartListElement) == 1)
                return;

            //Заносим в переменные все нужные параметры
            foreach (var apart in ApartListElement)
            {
                PNR_Function = apart.LookupParameter("PNR_Функция помещения").AsString();
                PNR_Name = apart.LookupParameter("PNR_Имя помещения").AsString();
                PNR_Floor = apart.LookupParameter("ADSK_Этаж").AsString();
                if (PNR_Floor.Substring(0, 1) == "-")
                {
                    PNR_Floor = "П" + PNR_Floor.Substring(1, PNR_Floor.Length - 1);
                }
                try
                {
                    PNR_Section = apart.LookupParameter("ADSK_Номер секции").AsString();
                    if (PNR_Section == null || _SelectedSectionValue == "")
                        PNR_Section = "";
                }
                catch { PNR_Section = ""; }
                try
                {
                    PNR_Building = apart.LookupParameter("ADSK_Номер здания").AsString() + "/";
                    if (PNR_Building == null || PNR_Building == "/")
                        PNR_Building = "";
                }
                catch { PNR_Building = ""; }
            }

            //Фильтр по всем помещениям на всей модели
            List<Element> AllRooms = new List<Element>();
            AllRooms = new FilteredElementCollector(_doc)
                         .WhereElementIsNotElementType()
                         .OfCategory(BuiltInCategory.OST_Rooms)
                         .Where(r => r.LookupParameter("ADSK_Номер секции").AsString() == PNR_Section || r.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor)
                         .ToList();

            //Для поиска максимального номера создаём лист со всеми помещениями на модели
            List<Element> AllRoomsWGroup = new List<Element>();
            foreach (var room in AllRooms)
            {
                if (room.GroupId.IntegerValue == -1)
                    AllRoomsWGroup.Add(room);
            }

            //Сокращаем функцию
            SetFunc(PNR_Function, PNR_Name, out PNR_Funс, out string flag);

            //Для определения максимального номера считываем количество N и секционность
            CountNBoolSection(PNR_Funс, out int countN, out bool section);

            //Находим максимальный номер у помещений
            MaxNumberApart(AllRoomsWGroup, PNR_Floor, PNR_Funс, out int number, countN, section);

            //Заносим параметры апартов и квартир
            SetParameterApart(ApartListElement, PNR_Funс);

            //Нумеруем помещения
            NumberApart(ApartListElement, PNR_Funс, PNR_Building, PNR_Section, PNR_Floor, number);

            var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
            apartWindow.ShowDialog();
        }
        public int CheckParameter(IList<Element> ApartListElement) //Проверка параметров
        {
            Transaction tr = new Transaction(_doc, "Проверка параметров модели");
            tr.Start();
            foreach (var apart in ApartListElement)
            {
                if (apart != null)
                {
                    try
                    {
                        if (apart.LookupParameter("PNR_Функция помещения").StorageType.ToString() != "String" ||
                            apart.LookupParameter("PNR_Имя помещения").StorageType.ToString() != "String" ||
                            apart.LookupParameter("PNR_Номер помещения").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Номер квартиры").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Номер секции").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Номер здания").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Диапазон").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Тип квартиры").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Позиция отделки").StorageType.ToString() != "String" ||
                            apart.LookupParameter("ADSK_Этаж").StorageType.ToString() != "String")
                        {
                            TaskDialog.Show("Тип параметра", "Проверьте типы данных параметров СТРОКА");
                            tr.RollBack();
                            return 1;
                        }

                        if (apart.LookupParameter("ADSK_Площадь квартиры жилая").StorageType.ToString() != "Double" ||
                            apart.LookupParameter("ADSK_Площадь квартиры общая").StorageType.ToString() != "Double" ||
                            apart.LookupParameter("ADSK_Количество комнат").StorageType.ToString() != "Integer")
                        {
                            TaskDialog.Show("Тип параметра", "Проверьте типы данных параметров ЧИСЛО или ПЛОЩАДЬ");
                            tr.RollBack();
                            return 1;
                        }
                    }
                    catch { TaskDialog.Show("Наличие параметров", "Проверьте наличие параметров"); tr.RollBack(); return 1; }

                    if (apart.LookupParameter("ADSK_Этаж").AsString() == null || apart.LookupParameter("ADSK_Этаж").AsString() == "" ||
                        apart.LookupParameter("PNR_Функция помещения").AsString() == null || apart.LookupParameter("PNR_Функция помещения").AsString() == "" ||
                        apart.LookupParameter("PNR_Имя помещения").AsString() == null || apart.LookupParameter("PNR_Имя помещения").AsString() == "")
                    {
                        TaskDialog.Show("Заполнение параметров", "Заполните следующие параметры: ADSK_Этаж, PNR_Функция помещения, PNR_Имя помещения");
                        tr.RollBack();
                        return 1;
                    }

                    if (apart.LookupParameter("PNR_Номер помещения").AsString() == null ||
                        apart.LookupParameter("ADSK_Номер квартиры").AsString() == null ||
                        apart.LookupParameter("ADSK_Диапазон").AsString() == null ||
                        apart.LookupParameter("ADSK_Тип квартиры").AsString() == null ||
                        apart.LookupParameter("ADSK_Позиция отделки").AsString() == null)
                    {
                        apart.LookupParameter("PNR_Номер помещения").Set("");
                        apart.LookupParameter("ADSK_Номер квартиры").Set("");
                        apart.LookupParameter("ADSK_Диапазон").Set("");
                        apart.LookupParameter("ADSK_Тип квартиры").Set("");
                        apart.LookupParameter("ADSK_Позиция отделки").Set("");
                    }
                }
            }
            tr.Commit();
            return 0;
        }
        public void CountNBoolSection(string PNR_Funс, out int countN, out bool section) //Количество N и секционность
        {
            countN = 0;
            section = false;
            var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Имена помещений.xlsx"));
            using (var package = new ExcelPackage(path))
            {
                var count = package.Workbook.Worksheets.Count;
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Name"];
                string range = "A2:I10000";
                var rangeCells = worksheet.Cells[range];
                object[,] Allvalues = rangeCells.Value as object[,];
                if (Allvalues != null)
                {
                    int rows = Allvalues.GetLength(0);
                    int columns = Allvalues.GetLength(1);

                    for (int i = 0; i <= rows; i++)
                    {
                        if (Allvalues[i, 0].ToString() == "end")
                            break;
                        if (Allvalues[i, 2].ToString() == PNR_Funс)
                        {
                            if (Allvalues[i, 6].ToString() != "-")
                                countN = Allvalues[i, 6].ToString().Count();
                            else
                                countN = -1;
                            if (Allvalues[i, 5].ToString() != "-")
                                section = true;
                            break;
                        }
                    }
                }
            }
        }
        public void SetParameterApart(IList<Element> ApartListElement, string PNR_Funс) //Параметры для апартов и квартир
        {
            //Заполняем параметры квартиры или апартов
            Transaction trn = new Transaction(_doc, "Set parameter");
            trn.Start();
            if (PNR_Funс.Contains("КВ") || PNR_Funс.Contains("АП"))
            {
                double LifeArea = 0.0;
                double GeneralArea = 0.0;
                foreach (var apart in ApartListElement)
                {
                    if (apart != null)
                    {
                        if (apart.LookupParameter("PNR_Имя помещения").AsString().Contains("Жилая") ||
                            apart.LookupParameter("PNR_Имя помещения").AsString().Contains("Гостиная") ||
                            apart.LookupParameter("PNR_Имя помещения").AsString().Contains("Спальня") ||
                            apart.LookupParameter("PNR_Имя помещения").AsString().Contains("Мастер-спальня"))
                        {
                            LifeArea += double.Parse(apart.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble().ToString());
                        }
                        GeneralArea += double.Parse(apart.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble().ToString());
                    }
                }
                foreach (var apart in ApartListElement)
                {
                    apart.LookupParameter("ADSK_Площадь квартиры жилая").Set(LifeArea);
                    apart.LookupParameter("ADSK_Площадь квартиры общая").Set(GeneralArea);
                }

                var areaParamter = new AreaParameter();
                areaParamter.ShowDialog();

                foreach (var apart in ApartListElement)
                {
                    apart.LookupParameter("ADSK_Тип квартиры").Set(areaParamter.selectedType);
                    apart.LookupParameter("ADSK_Диапазон").Set(areaParamter.selectedRange);
                    apart.LookupParameter("ADSK_Количество комнат").Set(areaParamter.selectedCount);
                }
            }
            trn.Commit();
        }
        public void SetFunc(string PNR_Function, string PNR_Name, out string PNR_Func, out string flag) //Сокращение функции и флаг шаблона
        {
            flag = string.Empty;
            PNR_Func = null;
            List<String> function = new List<String>();
            var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Имена помещений.xlsx"));
            using (var package = new ExcelPackage(path))
            {
                var count = package.Workbook.Worksheets.Count;
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Name"];
                string range = "A2:I10000";
                var rangeCells = worksheet.Cells[range];
                object[,] Allvalues = rangeCells.Value as object[,];
                if (Allvalues != null)
                {
                    int rows = Allvalues.GetLength(0);
                    int columns = Allvalues.GetLength(1);

                    for (int i = 0; i <= rows; i++)
                    {
                        if (Allvalues[i, 0].ToString() == "end")
                            break;
                        if (Allvalues[i, 1].ToString() == PNR_Function && Allvalues[i, 0].ToString() == PNR_Name)
                        {
                            PNR_Func = Allvalues[i, 2].ToString();
                            flag = Allvalues[i, 8].ToString();
                            break;
                        }
                    }
                }
            }
        }
        public void NumberApart(IList<Element> ApartListElement, string PNR_Funс, string PNR_Building, string PNR_Section, string PNR_Floor, int number) //Метод, нумерующий квартиры
        {
            //Лист с каждым помещением и его кодировкой
            List<List<object>> ApartCodeList = new List<List<object>>();
            var path = new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\Имена помещений.xlsx"));
            using (var package = new ExcelPackage(path))
            {
                var count = package.Workbook.Worksheets.Count;
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Name"];
                string range = "A2:I10000";
                var rangeCells = worksheet.Cells[range];
                object[,] AllRoows = rangeCells.Value as object[,];
                if (AllRoows != null)
                {
                    int rows = AllRoows.GetLength(0);
                    int columns = AllRoows.GetLength(1);

                    //Собираем всю информацию - объект-помещение, кодировка
                    foreach (var aparts in ApartListElement)
                    {
                        for (int n = 0; n <= rows; n++)
                        {
                            int A = 0; //Имя 
                            int B = 1; //Функция
                            int C = 2; //Сокращение
                            int D = 3; //Здание
                            int E = 4; //Этаж
                            int F = 5; //Секция
                            int G = 6; //Номер квартиры
                            int H = 7; //Номер помещения
                            int I = 8; //Флаг
                            if (AllRoows[n, A].ToString() == "end")
                                break;
                            if (AllRoows[n, A].ToString() == aparts.LookupParameter("PNR_Имя помещения").AsString() && AllRoows[n, B].ToString() == aparts.LookupParameter("PNR_Функция помещения").AsString())
                            {
                                List<object> apart = new List<object>
                                {
                                    aparts,         //0
                                    AllRoows[n, A], //1
                                    AllRoows[n, B], //2
                                    AllRoows[n, C], //3
                                    AllRoows[n, D], //4
                                    AllRoows[n, E], //5
                                    AllRoows[n, F], //6
                                    AllRoows[n, G], //7
                                    AllRoows[n, H], //8
                                    AllRoows[n, I]  //9
                                };
                                ApartCodeList.Add(apart);
                                break;
                            }
                        }
                    }
                }
            }

            Transaction tr = new Transaction(_doc, "Заносим параметр");
            tr.Start();
            //Нумеруем помещения
            try
            {
                //Переменная-индекс для нумерации помещений
                int i = 1;
                //Пррверка существование жучка
                bool bug = false;
                //Номер квартиры
                string num = string.Empty;

                foreach (var apart in ApartCodeList)
                {
                    string apartCode = string.Empty;

                    //Удаляем все "-"
                    apart.RemoveAll(a => a.ToString() == "-");

                    //формируем объект-строку эксель с объектом комнатой по 1 варианту
                    if (apart.Last().ToString() == "1")
                    {
                        apart.RemoveAt(apart.Count() - 1);
                        for (int j = 0; j <= apart.Count() - 1; j++)
                        {
                            if (j > 2)
                            {
                                apartCode += apart[j].ToString();
                            }
                        }
                    }

                    //формируем объект-строку эксель с объектом комнатой по 0 варианту
                    if (apart.Last().ToString() == "0")
                    {
                        apart.RemoveAt(apart.Count() - 1);
                        string func = string.Empty;
                        for (int j = 0; j <= apart.Count() - 1; j++)
                        {
                            if (j == 3)
                            {
                                func += apart[j].ToString() + ".";
                            }
                            if (j == 4)
                            {
                                apartCode += apart[j].ToString();
                            }
                            if (j == 5)
                            {
                                apartCode += apart[j].ToString() + "." + func;
                            }
                            if (j == 6)
                            {
                                apartCode += apart[j].ToString() ;
                            }
                        }
                    }

                    //Добавляем нули к этажу
                    if (apartCode.Count(c => c == 'F') == 2 && PNR_Floor.Length == 1)
                    {
                        PNR_Floor = "0" + PNR_Floor;
                    }

                    //Добавляем нули к секции
                    if (PNR_Section.Length == 1)
                    {
                        PNR_Section = "0" + PNR_Section;
                    }

                    //Проверяем длину номера и добавляем нули при необходимости (для квартир он тройной, для остальных двойной)
                    if (apartCode.Count(c => c == 'N') == 3)
                    {
                        if (number.ToString().Length == 1)
                            num = "00" + number;

                        if (number.ToString().Length == 2)
                            num = "0" + number;
                    }

                    if (apartCode.Count(c => c == 'N') == 2 && (number.ToString().Length == 1) || apartCode.Count(c => c == 'B') == 2 && (number.ToString().Length == 1))
                    {
                        num = "0" + number;
                    }

                    string aparatNumberApart = apartCode.Replace("Y/", PNR_Building)
                               .Replace("FF", PNR_Floor)
                               .Replace("F", PNR_Floor)
                               .Replace("SS", PNR_Section)
                               .Replace("BB", num)
                               .Replace("NNN", num)
                               .Replace("NN", num)
                               .Replace("K", "");

                    string aparatNumberRoom = aparatNumberApart + "." + i.ToString();

                    if (apartCode.Count(c => c == 'K') == 3)
                    {
                        string index = i.ToString();
                        if(i.ToString().Length == 1)
                            index = "0" + i.ToString();
                        aparatNumberRoom = aparatNumberApart + index;
                    }

                    ((Element)apart[0]).LookupParameter("ADSK_Номер квартиры").Set(aparatNumberApart);
                    ((Element)apart[0]).LookupParameter("PNR_Номер помещения").Set(aparatNumberRoom);

                    if (!bug && PNR_Funс == "КВ-" || !bug && PNR_Funс == "АП-")
                    {
                        if (i == 1)
                        {
                            Bug((Element)apart[0]);
                            bug = true;
                        }
                    }

                    i++;
                }

            }
            catch (Exception ex)
            {
                TaskDialog.Show("Не все параметры существуют в модели", "Проверьте существование параметров: ADSK_Номер квартиры, PNR_Номер помещения. " +
                    "Если параметры в модели присутсвуют, проверьте наличие семейства метки квартиры (id: 138092)");
                tr.RollBack();
                return;
            }
            tr.Commit();
        }
        public void MaxNumberApart(IList<Element> AllRoomsWGroup, string PNR_Floor, string PNR_Funс, out int number, int countN, bool section) //Максимальный номер помещений
        {   //Высчитываем максимальный номер
            // try
            // {
            List<int> numberList = new List<int>();
            foreach (var room in AllRoomsWGroup)
            {
                //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО СЕКЦИИ
                if (room.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue && section == true)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                    {
                        continue;
                    }
                    string Func = room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(0, 3);
                    if (countN == 3 && Func == PNR_Funс)
                    {
                        int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                        numberList.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0')));
                        continue;
                    }
                }

                //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО ЭТАЖУ
                if (room.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor && section != true)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                    {
                        continue;
                    }

                    string Func = room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(0, 3);
                    if (!Func.Contains("-"))
                    {
                        int i = room.LookupParameter("ADSK_Номер квартиры").AsString().IndexOf(".");
                        Func = room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(i+1, 2);
                    }
                    if (countN == 3 && Func == PNR_Funс)
                    {
                        int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                        numberList.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0')));
                        continue;
                    }

                    if (countN == 2 && Func == PNR_Funс)
                    {
                        int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                        numberList.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0')));
                        continue;
                    }

                    if (countN == 1 && Func == PNR_Funс)
                    {
                        int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                        numberList.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 1, 1).Trim('0')));
                        continue;
                    }
                }
            }
            if (numberList.Count() != 0)
                number = numberList.Max();
            else number = 0;
            number++;
        }
        private void Bug(Element room) //Марка-жучок квартир
        {
            LinkElementId roomId = new LinkElementId(room.Id);
            LocationPoint roomLocation = room.Location as LocationPoint;
            UV uv = new UV(roomLocation.Point.X, roomLocation.Point.Y);
            RoomTag roomTag = _doc.Create.NewRoomTag(roomId, uv, _doc.ActiveView.Id);
            var type = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                //.Where(g => g.Id.ToString() == "138092");
                .Where(g => g.Family.Name == "050_Марка_Помещение_Квартира_2.5мм")
                .Where(g => g.Name == "Полная марка");
            roomTag.ChangeTypeId(type.First().Id);
            roomTag.HasLeader = true;
            List<Autodesk.Revit.DB.Grid> grids = new List<Autodesk.Revit.DB.Grid>(
                new FilteredElementCollector(_doc)
                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                .Cast<Autodesk.Revit.DB.Grid>()
                .Where(g => ((Autodesk.Revit.DB.Line)g.Curve).Direction.X == 1.0 || ((Autodesk.Revit.DB.Line)g.Curve).Direction.X == -1.0));

            Autodesk.Revit.DB.Grid gridHighest = grids[0];
            Autodesk.Revit.DB.Grid gridLowest = grids[0];
            Autodesk.Revit.DB.Line lineHighest = gridHighest.Curve as Autodesk.Revit.DB.Line;
            Autodesk.Revit.DB.Line lineLowest = gridLowest.Curve as Autodesk.Revit.DB.Line;
            foreach (var grid in grids)
            {
                Autodesk.Revit.DB.Line line = grid.Curve as Autodesk.Revit.DB.Line;
                if (lineHighest.Origin.Y < line.Origin.Y)
                {
                    gridHighest = grid;
                }
                if (lineLowest.Origin.Y > line.Origin.Y)
                {
                    gridLowest = grid;
                }
                lineHighest = gridHighest.Curve as Autodesk.Revit.DB.Line;
                lineLowest = gridLowest.Curve as Autodesk.Revit.DB.Line;
            }

            if ((lineHighest.Origin.Y - roomTag.TagHeadPosition.Y) < (roomTag.TagHeadPosition.Y - lineLowest.Origin.Y))
            {
                roomTag.TagHeadPosition = new XYZ(roomLocation.Point.X, lineHighest.Origin.Y, roomLocation.Point.Z) + new XYZ(0, 15, 0);
            }
            if ((lineHighest.Origin.Y - roomTag.TagHeadPosition.Y) > (roomTag.TagHeadPosition.Y - lineLowest.Origin.Y))
            {
                roomTag.TagHeadPosition = new XYZ(roomLocation.Point.X, lineLowest.Origin.Y, roomLocation.Point.Z) + new XYZ(0, -15, 0);
            }
        }
        public void ApartList() //Метод, считывающий все помещения для ListView
        {
            //Фильтр по всем квартирам - помещениям на данном виде
            var AllRooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .ToList();

            //Заносим все помещения на активном виде в ListView
            foreach (var room in AllRooms)
            {
                if (room.GroupId.ToString() == "-1")
                {
                    if (room.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue || (_SelectedSectionValue == "" && room.LookupParameter("ADSK_Номер секции").AsString() == null))
                    {
                        try
                        {
                            ClassApartAndRoom roomInList = new ClassApartAndRoom(room.LookupParameter("ADSK_Номер квартиры").AsString(), room.LookupParameter("PNR_Номер помещения").AsString(), room.Name, room.Id);
                            SelectedApartList.Add(roomInList);
                        }
                        catch (Exception e)
                        {
                            ClassApartAndRoom roomInList = new ClassApartAndRoom("ПАРАМЕТР НЕ НАЙДЕН", "ПАРАМЕТР НЕ НАЙДЕН", room.Name, room.Id);
                            SelectedApartList.Add(roomInList);
                        }
                    }
                }
            }

            //Сортировка по номеру квартиры и номеру помещения
            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(SelectedApartList);
            view.SortDescriptions.Add(new SortDescription("ADSK_Номер_квартиры", ListSortDirection.Ascending));

            CollectionView view2 = (CollectionView)CollectionViewSource.GetDefaultView(SelectedApartList);
            view.SortDescriptions.Add(new SortDescription("PNR_Номер_помещения", ListSortDirection.Ascending));
        }
        private void SelectAparts_2()
        {
            RaiseCloseRequest();
            var sectionWindow = new SelectSection(_uiapp, _uidoc, _doc, 2);
            sectionWindow.ShowDialog();
        }
        public void RenumberApart() //Перенумеровка квартир
        {
            RaiseCloseRequest();

            try
            { }
            catch (Exception e)
            {
                TaskDialog.Show("Ошибка", $"Выберите помещение");
                var apartWindow1 = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                apartWindow1.ShowDialog();
            }

            //Фильтр по всем квартирам - помещениям на данном виде
            AllRoomsRenumberNew = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() == SelectedRoom.ADSK_Номер_квартиры)
                .ToList();

            var renumber = new Renumber(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2, AllRoomsRenumberNew, 0);
            renumber.ShowDialog();
        }
        public void RenumberRoom() //Перенумеровка комнат
        {
            RaiseCloseRequest();

            try
            { }
            catch (Exception e)
            {
                TaskDialog.Show("Ошибка", $"Выберите помещение");
                var apartWindow1 = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                apartWindow1.ShowDialog();
            }

            //Фильтр по всем квартирам - помещениям на данном виде
            AllRoomsRenumberNew = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(g => g.Id == SelectedRoom.id)
                .ToList();

            var renumber = new Renumber(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2, AllRoomsRenumberNew, 1);
            renumber.ShowDialog();
        }
        public void EnterRenumber()
        {
            RaiseCloseRequest();

            if (index == 0)//Переименовываем квартиры
            {
                int length = 0;
                string ADSK_Номер_квартиры = string.Empty;
                string func = string.Empty;
                foreach (var a in AllRoomsRenumberNew)
                {
                    length = a.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                    func = a.LookupParameter("ADSK_Номер квартиры").AsString().Substring(0, 2);
                    break;
                }

                List<Element> rooms = new List<Element>();

                if (func != "КВ")
                {
                    rooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                   .WhereElementIsNotElementType()
                   .OfCategory(BuiltInCategory.OST_Rooms)
                   .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0') == NewValueNumberApart)
                   .ToList();
                }
                if (func == "КВ")
                {
                    rooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0') == NewValueNumberApart)
                    .ToList();
                }

                if (rooms.Count > 0)
                {
                    foreach (var room in rooms)
                    {
                        AllRoomsRenumberNew.Add(room);
                    }
                }

                Transaction tr = new Transaction(_doc, "Renumber");
                int i = 1;
                tr.Start();
                foreach (var room in AllRoomsRenumberNew)
                {
                    string str = room.LookupParameter("ADSK_Номер квартиры").AsString();
                    if (str.Length > 0)
                    {
                        if (func == "КВ")
                            str = str.Remove(str.Length - 3);
                        if (func != "КВ")
                            str = str.Remove(str.Length - 2);
                    };
                    if (NewValueNumberApart.Length == 1)
                    {
                        if (func == "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "00" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "00" + NewValueNumberApart + "." + i);
                        }
                        if (func != "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "." + i);
                        }
                        if (func == "ХК")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            if (i > 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + i);
                            if (i < 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "0" + i);
                        }
                    }
                    if (NewValueNumberApart.Length == 2)
                    {
                        if (func == "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "." + i);
                        }
                        if (func != "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "." + i);
                        }
                        if (func == "ХК")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                            if (i > 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + i);
                            if (i < 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "0" + i);
                        }
                    }
                    if (NewValueNumberApart.Length == 3)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                        room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "." + i);
                    }
                    i++;
                }
                i = 0;
                tr.Commit();
                var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                AllRoomsRenumberNew.Clear();
                apartWindow.ShowDialog();
            }


            if (index == 1) //Переименовываем помещения
            {
                int length = 0;
                string func = string.Empty;
                string ADSK_Номер_квартиры = string.Empty;
                string OldValueNumberApart = string.Empty;
                var selectedRoom = _doc.GetElement(SelectedRoom.id);

                if (selectedRoom != null)
                {
                    length = selectedRoom.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                    func = selectedRoom.LookupParameter("ADSK_Номер квартиры").AsString().Substring(0, 2);
                    if (func == "КВ")
                    {
                        OldValueNumberApart = selectedRoom.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0');
                    }
                    if (func != "КВ")
                    {
                        OldValueNumberApart = selectedRoom.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0');
                    }
                }

                List<Element> NewRooms = new List<Element>();
                List<Element> OldRooms = new List<Element>();

                if (func != "КВ")
                {
                    NewRooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                   .WhereElementIsNotElementType()
                   .OfCategory(BuiltInCategory.OST_Rooms)
                   .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0') == NewValueNumberApart)
                   .ToList();

                    OldRooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                   .WhereElementIsNotElementType()
                   .OfCategory(BuiltInCategory.OST_Rooms)
                   .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0') == OldValueNumberApart)
                   .ToList();
                }
                if (func == "КВ")
                {
                    NewRooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0') == NewValueNumberApart)
                    .ToList();

                    OldRooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                   .WhereElementIsNotElementType()
                   .OfCategory(BuiltInCategory.OST_Rooms)
                   .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                   .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0') == OldValueNumberApart)
                   .ToList();
                }

                OldRooms.RemoveAll(g => g.Id == selectedRoom.Id);
                NewRooms.Add(selectedRoom);

                /*                if (NewRooms.Count > 0)
                                {
                                    foreach (var room in NewRooms)
                                    {
                                        AllRoomsRenumberNew.Add(room);
                                    }
                                }*/

                Transaction tr = new Transaction(_doc, "Renumber");
                int i = 1;
                tr.Start();
                //Переименовываем новые квартиры с новым помещением
                foreach (var room in NewRooms)
                {
                    string str = room.LookupParameter("ADSK_Номер квартиры").AsString();
                    if (str.Length > 0)
                    {
                        if (func == "КВ")
                            str = str.Remove(str.Length - 3);
                        if (func != "КВ")
                            str = str.Remove(str.Length - 2);
                    }
                    if (NewValueNumberApart.Length == 1)
                    {
                        if (func == "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "00" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "00" + NewValueNumberApart + "." + i);
                        }
                        if (func != "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "." + i);
                        }
                        if (func == "ХК")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            if (i > 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + i);
                            if (i < 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "0" + i);
                        }
                    }
                    if (NewValueNumberApart.Length == 2)
                    {
                        if (func == "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + "0" + NewValueNumberApart + "." + i);
                        }
                        if (func != "КВ")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                            room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "." + i);
                        }
                        if (func == "ХК")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                            if (i > 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + i);
                            if (i < 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "0" + i);
                        }
                    }
                    if (NewValueNumberApart.Length == 3)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + NewValueNumberApart);
                        room.LookupParameter("PNR_Номер помещения").Set(str + NewValueNumberApart + "." + i);
                    }
                    i++;
                }

                i = 1;

                foreach (var room in OldRooms)
                {
                    string str = room.LookupParameter("ADSK_Номер квартиры").AsString();
                    if (str.Length > 0)
                    {
                        if (func == "ХК")
                        {
                            if (i > 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + i);
                            if (i < 9)
                                room.LookupParameter("PNR_Номер помещения").Set(str + "0" + i);
                        }
                        if (func != "ХК")
                            room.LookupParameter("PNR_Номер помещения").Set(str + "." + i);
                    }
                    i++;
                }
                tr.Commit();
                var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                AllRoomsRenumberNew.Clear();
                apartWindow.ShowDialog();
            }
        }
    }
    public class RoomPickFilter : ISelectionFilter //Фильтр для комнат
    {
        public bool AllowElement(Element e)
        {
            return (e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Rooms));
            //return true;
        }
        public bool AllowReference(Autodesk.Revit.DB.Reference r, XYZ p)
        {
            return false;
        }

    }
}
