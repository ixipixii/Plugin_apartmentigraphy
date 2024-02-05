﻿using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        public void RoomList() //******метод, считывающий все помещения для ListView
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
                /*                if(room.LookupParameter("PNR_Функция помещения").AsString() == "Помещение управляющей компании")
                                {
                                    Group group = _doc.GetElement(room.GroupId) as Group;
                                    try
                                    {
                                        ClassApartAndRoom room_УК_InList = new ClassApartAndRoom("0", group.LookupParameter("ADSK_Номер квартиры").AsString(), group.Name);
                                        if (!SelectedRoomList.Contains(room_УК_InList))
                                            SelectedRoomList.Add(room_УК_InList);
                                    }
                                    catch
                                    {
                                        ClassApartAndRoom room_УК_InList = new ClassApartAndRoom("0", "ПАРАМЕТР НЕ НАЙДЕН", group.Name);
                                        if (!SelectedRoomList.Contains(room_УК_InList))
                                            SelectedRoomList.Add(room_УК_InList);
                                    }
                                }*/
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
        public void SelectionRoom() //******метод, запускающий поиск помещений
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
                /*                if (roomElement.LookupParameter("PNR_Функция помещения").AsString() == "Помещение управляющей компании")
                                    if(OvverideContains(RoomListElement, _doc.GetElement(roomElement.GroupId)))
                                        RoomListElement.Add(_doc.GetElement(roomElement.GroupId));*/
            }

            //Нумеруем помещения
            NumberRoom(RoomListElement);

            //После завершения нумерации выводим список помещений
            var room = new Room(_uiapp, _uidoc, _doc);
            room.ShowDialog();
        }
        public void NumberRoom(IList<Element> RoomListElement) //*****метод, нумерующий комнаты
        {
            // Параметры помещений
            String PNR_Function = string.Empty;
            String PNR_Floor = string.Empty;
            String PNR_Building = string.Empty;

            //Переменные количества помещений определённой функции
            int number_МОП = 0;
            int number_СРВ = 0;
            int number_Т = 0;
            int number_А = 0;
            int number_УК = 0;

            String PNR_Funс = string.Empty; //Ссылка на функцию помещения
            String numberRoom = string.Empty; //Итоговый номер помещения

            foreach (var room in RoomListElement)
            {
                if ((BuiltInCategory)room.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups)
                {
                    Group group = room as Group;
                    Transaction tr = new Transaction(_doc, "Number room");
                    tr.Start();
                    var roomInGroup = group.UngroupMembers().ToList();
                    tr.RollBack();
                    PNR_Floor = _doc.GetElement(roomInGroup[0]).LookupParameter("ADSK_Этаж").AsString();
                    try
                    {
                        PNR_Building = _doc.GetElement(roomInGroup[0]).LookupParameter("ADSK_Номер здания").AsString();
                        if (PNR_Building == "" || PNR_Building == null)
                            PNR_Building = "-1";
                    }
                    catch { PNR_Building = "-1"; }
                    number_УК++;

                    tr.Start();
                    if (PNR_Building == "-1")
                    {
                        group.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Floor}.УК.{number_УК}");
                        foreach (var room_ in roomInGroup)
                        {
                            _doc.GetElement(room_).LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.УК.{number_УК}");
                        }
                    }
                    else
                    {
                        group.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Building}/{PNR_Floor}.УК.{number_УК}");
                        foreach (var room_ in roomInGroup)
                        {
                            _doc.GetElement(room_).LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.УК.{number_УК}");
                        }
                    }
                    tr.Commit();
                    continue;
                }

                PNR_Function = room.LookupParameter("PNR_Функция помещения").AsString();
                PNR_Floor = room.LookupParameter("ADSK_Этаж").AsString();
                try
                {
                    PNR_Building = room.LookupParameter("ADSK_Номер здания").AsString();
                    if (PNR_Building == "" || PNR_Building == null)
                        PNR_Building = "-1";
                }
                catch { PNR_Building = "-1"; }

                if (PNR_Function == "МОП входной группы 1 этажа" ||
                   PNR_Function == "МОП входной группы -1 этажа" ||
                   PNR_Function == "МОП типовых этажей" ||
                   PNR_Function == "Лестницы эвакуации  (с -1го до последнего этажа)" ||
                   PNR_Function == "Паркинг" ||
                   PNR_Function == "Помещения загрузки" ||
                   PNR_Function == "Помещения мусороудаления" ||
                   PNR_Function == "Инженерно-технические помещения" ||
                   PNR_Function == "Объединенный диспетчерский пункт" ||
                   PNR_Function == "Помещения линейного и обслуживающего персонала" ||
                   PNR_Function == "Помещения охраны" ||
                   PNR_Function == "Помещения Клининговых служб")
                {
                    switch (PNR_Function)
                    {
                        case "МОП входной группы 1 этажа":
                            PNR_Funс = "МОП";
                            break;

                        case "МОП входной группы -1 этажа":
                            PNR_Funс = "МОП";
                            break;

                        case "МОП типовых этажей":
                            PNR_Funс = "МОП";
                            break;

                        case "Лестницы эвакуации  (с -1го до последнего этажа)":
                            PNR_Funс = "МОП";
                            break;

                        case "Паркинг":
                            PNR_Funс = "А";
                            break;

                        case "Помещения загрузки":
                            PNR_Funс = "СРВ";
                            break;

                        case "Помещения мусороудаления":
                            PNR_Funс = "СРВ";
                            break;

                        case "Инженерно-технические помещения":
                            PNR_Funс = "Т";
                            break;

                        case "Объединенный диспетчерский пункт":
                            PNR_Funс = "СРВ";
                            break;

                        case "Помещения линейного и обслуживающего персонала":
                            PNR_Funс = "СРВ";
                            break;

                        case "Помещения охраны":
                            PNR_Funс = "СРВ";
                            break;

                        case "Помещения Клининговых служб":
                            PNR_Funс = "СРВ";
                            break;
                    }

                    if (PNR_Funс == "МОП")
                    {
                        number_МОП++;
                        Transaction tr = new Transaction(_doc, "Number room");
                        tr.Start();
                        if (PNR_Building == "-1")
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.{PNR_Funс}.{number_МОП}");
                        else
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_МОП}");
                        tr.Commit();
                    }

                    if (PNR_Funс == "А")
                    {
                        number_А++;
                        Transaction tr = new Transaction(_doc, "Number room");
                        tr.Start();
                        if (PNR_Building == "-1")
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.{PNR_Funс}.{number_А}");
                        else
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_А}");
                        tr.Commit();
                    }

                    if (PNR_Funс == "СРВ")
                    {
                        number_СРВ++;
                        Transaction tr = new Transaction(_doc, "Number room");
                        tr.Start();
                        if (PNR_Building == "-1")
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.{PNR_Funс}.{number_СРВ}");
                        else
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_СРВ}");
                        tr.Commit();
                    }

                    if (PNR_Funс == "Т")
                    {
                        number_Т++;
                        Transaction tr = new Transaction(_doc, "Number room");
                        tr.Start();
                        if (PNR_Building == "-1")
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.{PNR_Funс}.{number_Т}");
                        else
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_Т}");
                        tr.Commit();
                    }
                }
            }
        }


        //version 2
        public void SelectionApart()
        {
            RaiseCloseRequest();

            List<int> numberMax_КВ = new List<int>();
            List<int> numberMax_АП = new List<int>();
            List<int> numberMax_КМ = new List<int>();
            List<int> numberMax_ХК = new List<int>();
            List<int> numberMax_УК = new List<int>();

            String PNR_Function = string.Empty;
            String PNR_Section = string.Empty;
            String PNR_Building = string.Empty;
            String PNR_Funс = string.Empty;
            String PNR_Floor = string.Empty;

            var roomFilter = new RoomPickFilter();

            //Создаём листы: ссылки на элементы, лист помещений
            IList<Autodesk.Revit.DB.Reference> refrence = new List<Autodesk.Revit.DB.Reference>();
            IList<Element> ApartListElement = new List<Element>();

            var refElement = _uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите элементы");

            foreach (var element in refElement)
            {
                var apart = _doc.GetElement(element);
                ApartListElement.Add(apart);
                PNR_Function = apart.LookupParameter("PNR_Функция помещения").AsString();
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
            try
            {
                AllRooms = new FilteredElementCollector(_doc)
                             .WhereElementIsNotElementType()
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .Where(r => r.LookupParameter("ADSK_Номер секции").AsString() == PNR_Section || r.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor)
                            .ToList();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Не все параметры существуют в модели", "Проверьте существование параметров: ADSK_Номер секции, ADSK_Этаж");
                return;
            }

            List<Element> AllRoomsWGroup = new List<Element>();
            foreach (var room in AllRooms)
            {
                if (room.GroupId.IntegerValue == -1)
                    AllRoomsWGroup.Add(room);
            }

            //Находим максимальный номер у помещений
            MaxNumberApart(AllRoomsWGroup, PNR_Floor, numberMax_КВ, numberMax_ХК, numberMax_АП,
                                numberMax_КМ,
                                numberMax_УК);

            if (PNR_Function == "Квартиры бед доп. отделки" ||
                            PNR_Function == "Апартаменты без доп. Отделки" ||
                            PNR_Function == "Коммерческие помещения без доп. отделки" ||
                            PNR_Function == "Квартиры с отделкой" ||
                            PNR_Function == "Апартаменты с отделкой" ||
                            PNR_Function == "Коммерческие помещения с отделкой" ||
                            PNR_Function == "Помещения кладовых" ||
                            PNR_Function == "Помещение управляющей компании")
            {
                switch (PNR_Function)
                {
                    case "Квартиры бед доп. отделки":
                        PNR_Funс = "КВ";
                        break;

                    case "Квартиры с отделкой":
                        PNR_Funс = "КВ";
                        break;

                    case "Апартаменты без доп. Отделки":
                        PNR_Funс = "АП";
                        break;

                    case "Апартаменты с отделкой":
                        PNR_Funс = "АП";
                        break;

                    case "Коммерческие помещения без доп. отделки":
                        PNR_Funс = "КМ";
                        break;

                    case "Коммерческие помещения с отделкой":
                        PNR_Funс = "КМ";
                        break;

                    case "Помещения кладовых":
                        PNR_Funс = "ХК";
                        break;

                    case "Помещение управляющей компании":
                        PNR_Funс = "УК";
                        break;
                }
            }

            try
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
                            if (apart.LookupParameter("PNR_Номер помещения").AsString().Contains("Жилая") ||
                                apart.LookupParameter("PNR_Номер помещения").AsString().Contains("Гостиная") ||
                                apart.LookupParameter("PNR_Номер помещения").AsString().Contains("Спальня") ||
                                apart.LookupParameter("PNR_Номер помещения").AsString().Contains("Мастер-спальня"))
                            {
                                LifeArea += double.Parse(apart.get_Parameter(BuiltInParameter.ROOM_AREA).AsValueString());
                            }
                            GeneralArea += double.Parse(apart.get_Parameter(BuiltInParameter.ROOM_AREA).AsValueString());
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
            catch
            {
                TaskDialog.Show("Не все параметры существуют в модели", "Проверьте существование параметров: ADSK_Площадь квартиры жилая, ADSK_Площадь квартиры общая, " +
                    "ADSK_Тип квартиры, ADSK_Диапазон, ADSK_Количество комнат");
                return;
            }

            //Нумеруем помещения
            NumberApart(ApartListElement, PNR_Funс, PNR_Building, PNR_Section, PNR_Floor, numberMax_КВ, numberMax_ХК, numberMax_АП, numberMax_КМ);

            var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
            apartWindow.ShowDialog();
        }
        public void NumberApart(IList<Element> ApartListElement, string PNR_Funс, string PNR_Building, string PNR_Section, string PNR_Floor, 
                                List<int> numberMax_КВ,
                                List<int> numberMax_ХК,
                                List<int> numberMax_АП,
                                List<int> numberMax_КМ) //Нумерация помещений
        {
            //Переменная-индекс для нумерации помещений
            int i = 1;
            //Пррверка существование жучка
            bool bug = false;

            Transaction tr = new Transaction(_doc, "Заносим параметр");
            tr.Start();
            //Нумеруем помещения
            try
            {
                foreach (var room in ApartListElement)
                {
                    int number = 0;
                    //Проставляем в number максимальное значение квартиры в зависимости от функции (если ошибка - значит квартир нет и номер = 1)
                    try
                    {
                        if (PNR_Funс == "КВ")
                        {
                            number = numberMax_КВ.Max(); number++;
                        }

                        if (PNR_Funс == "ХК")
                        {
                            number = numberMax_ХК.Max(); number++;
                        }

                        if (PNR_Funс == "АП")
                        {
                            number = numberMax_АП.Max(); number++;
                        }

                        if (PNR_Funс == "КМ")
                        {
                            number = numberMax_КМ.Max(); number++;
                        }
                    }
                    catch { number = 1; }

                    //Формируем номера квартир
                    if (PNR_Funс == "КВ")
                    {
                        string num = number.ToString();
                        if (num.Length == 1)
                            num = "00" + num;
                        if (num.Length == 2)
                            num = "0" + num;

                        room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}{PNR_Section}{num}");
                        room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}{PNR_Section}{num}.{i}");
                        if (!bug)
                        {
                            if (ApartListElement.Count == i)
                            {
                                Bug(room);
                                bug = true;
                            }
                        }
                    }

                    //Формируем номера апартов, кладовых, коммерции
                    if (PNR_Funс == "ХК" || PNR_Funс == "АП" || PNR_Funс == "КМ")
                    {
                        string num = number.ToString();
                        if (num.Length == 1)
                            num = "0" + num;

                        if (PNR_Funс == "АП" || PNR_Funс == "КМ")
                        {
                            //Если это коммерция, проверяем на стрит-ритейл (если true то в номере убираем этаж)
                            if (PNR_Floor == "1" && PNR_Funс == "КМ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}{num}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}{num}.{i}");
                            }
                            else
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}{PNR_Floor}{num}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}{PNR_Floor}{num}.{i}");
                            }
                            if (!bug)
                            {
                                if (ApartListElement.Count == i)
                                {
                                    Bug(room);
                                    bug = true;
                                }
                            }
                        }

                        //Формируем номера кладовых
                        if (PNR_Funс == "ХК")
                        {
                            string iKX = i.ToString();
                            if (iKX.Length == 1)
                                iKX = "0" + iKX;
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}{num}{i}");
                            room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}{num}");
                        }
                    }

                    number = 0;
                    i++;
                    bug = false;
                }

            }
            catch (Exception ex)
            {
                TaskDialog.Show("Не все параметры существуют в модели", "Проверьте существование параметров: ADSK_Номер квартиры, PNR_Номер помещения. " +
                    "Если параметры в модели присутсвуют, проверьте наличие семейства метки квартиры (id: 138092)");
                return;
            }
            tr.Commit();
        }

        public void MaxNumberApart(IList<Element> AllRoomsWGroup, string PNR_Floor,
                                List<int> numberMax_КВ,
                                List<int> numberMax_ХК,
                                List<int> numberMax_АП,
                                List<int> numberMax_КМ,
                                List<int> numberMax_УК) //Максимальный номер помещений
        {   //Высчитываем максимальный номер
            try
            {
                foreach (var room in AllRoomsWGroup)
                {
                    //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО СЕКЦИИ ДЛЯ КВ
                    if (room.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                    {
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                        {
                            continue;
                        }
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                        {
                            int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                            numberMax_КВ.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0')));
                        }
                    }

                    //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО ЭТАЖУ ДЛЯ АП, КМ, ХК, УК
                    if (room.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor)
                    {
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                        {
                            continue;
                        }

                        //Высчитываем максимальный номер АП
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("АП"))
                        {
                            int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                            numberMax_АП.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0')));
                        }

                        //Высчитываем максимальный номер КМ
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КМ"))
                        {
                            int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                            numberMax_КМ.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 2, 2).Trim('0')));
                        }

                        //Высчитываем максимальный номер УК
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("УК"))
                        {
                            int length = room.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                            numberMax_УК.Add(int.Parse(room.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 1, 1).Trim('0')));
                        }

                        //Высчитываем максимальный номер ХК
                        if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("ХК"))
                        {
                            var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
                            if (numberString.Contains("/"))
                            {
                                int index = numberString.IndexOf("/");
                                numberMax_ХК.Add(int.Parse(numberString.Substring(index + 1, 2).TrimStart('0')));
                                continue;
                            }
                            if (numberString.Contains("/") == false)
                            {
                                int index = numberString.IndexOf("-");
                                numberMax_ХК.Add(int.Parse(numberString.Substring(index + 1, 2).TrimStart('0')));
                                continue;
                            }
                        }
                    }
                }
            }
            catch
            {
                TaskDialog.Show("Не все параметры существуют в модели", "Проверьте существование параметров: ADSK_Номер секции, ADSK_Этаж, ADSK_Номер квартиры");
                return;
            }
        }
        private void Bug(Element room)
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
            /*            Location loc = room.Location;
                        if(loc is LocationPoint locationPoint)
                        {
                            XYZ roomlocation = locationPoint.Point;
                            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(_doc);
                            SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry((SpatialElement)room);

                        }*/
        }
        public void ApartList()
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
        public void RenumberApart()
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
        public void RenumberRoom()
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
            //return (e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Rooms));
            return true;
        }
        public bool AllowReference(Autodesk.Revit.DB.Reference r, XYZ p)
        {
            return false;
        }

    }
    public class GroupPickFilter : ISelectionFilter //Фильтр для групп
    {
        Document _doc;
        public GroupPickFilter(Document doc)
        {
            _doc = doc;
        }
        public bool AllowElement(Element e)
        {
            if (e is Group && e.LookupParameter("ADSK_Номер квартиры") != null)
                return e is Group;
            return false;
        }
        public bool AllowReference(Autodesk.Revit.DB.Reference r, XYZ p)
        {
            return false;
        }

    }
}
