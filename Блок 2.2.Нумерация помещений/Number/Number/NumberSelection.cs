﻿using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Number
{
    internal class NumberSelection
    {
        UIApplication _uiapp;
        UIDocument _uidoc;
        Document _doc;

        public List<Group> AllGroupsApart = new List<Group>(); //Лист всех групп-помещений
        public DelegateCommand SelectCommandApart { get; } //Делегат кнопки "Пронумеровать" группы
        public DelegateCommand SelectCommandRoom { get; } //Делегат кнопки "Пронумеровать" помещения
        public DelegateCommand SelectApart { get; } //Делегат кнопки "Нумерация групп помещений"
        public DelegateCommand SelectRoom { get; } //Делегат кнопки "Нумерация помещений"
        public List<ClassApartAndRoom> SelectedApartList { get; set; } = new List<ClassApartAndRoom>(); //Вспомогательный лист для ListView для групп 
        public List<ClassApartAndRoom> SelectedRoomList { get; set; } = new List<ClassApartAndRoom>(); //Вспомогательный лист для ListView для помещений
        public String _SelectedSectionValue { get; set; } //выбранная секция в ComboBox

        public NumberSelection(UIApplication uiapp, UIDocument uidoc, Document doc, string SelectedSectionValue)
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            SelectCommandApart = new DelegateCommand(SelectionApart);
            SelectCommandRoom = new DelegateCommand(SelectionRoom);
            SelectApart = new DelegateCommand(SelectAparts);
            SelectRoom = new DelegateCommand(SelectRooms);
            _SelectedSectionValue = SelectedSectionValue;
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
            var sectionWindow = new SelectSection(_uiapp, _uidoc, _doc);
            sectionWindow.ShowDialog();
        }
        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
        public void RoomList() //метод, считывающий все помещения для ListView
        {
            //Фильтр по всем помещениям на данном виде
            var AllRooms = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .ToList();

            //Заносим все помещения на активном виде в ListView
            foreach(var room in AllRooms)
            {
                if (room.GroupId.ToString() == "-1")
                {
                    try
                    {
                        ClassApartAndRoom roomInList = new ClassApartAndRoom("0", room.LookupParameter("PNR_Номер помещения").AsString(), room.Name);
                        SelectedRoomList.Add(roomInList);
                    }
                    catch
                    {
                        ClassApartAndRoom roomInList = new ClassApartAndRoom("0", "ПАРАМЕТР НЕ НАЙДЕН", room.Name);
                        SelectedRoomList.Add(roomInList);
                    }

                }
                if(room.LookupParameter("PNR_Функция помещения").AsString() == "Помещение управляющей компании")
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
                }
            }
        }
        public void ApartList() //метод, считывающий все группы с помещениями для ListView
        {
            //Фильтр по всем группам модели
            var AllGroups = new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Group))
                .Cast<Group>()
                .ToList();

            //Цикл по всем группам
            Transaction transaction = new Transaction(_doc, "UnGroup");
            foreach (var ApartGroup in AllGroups)
            {
                transaction.Start();
                var elementApartGroup = ApartGroup.UngroupMembers().ToList();
                transaction.RollBack();
                
                //Если все элементы группы - помещения, заносим группу в ListView
                if (elementApartGroup.All(g => (BuiltInCategory)_doc.GetElement(g).Category.Id.IntegerValue == BuiltInCategory.OST_Rooms))
                {
                    if(elementApartGroup.All(g => _doc.GetElement(g).LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue))
                    {
                    AllGroupsApart.Add(ApartGroup);
                        try
                        {
                            ClassApartAndRoom apart = new ClassApartAndRoom(
                                                        ApartGroup.LookupParameter("ADSK_Номер квартиры").AsString(),
                                                        "0",
                                                        ApartGroup.Name);
                            SelectedApartList.Add(apart);
                        }
                        catch
                        {
                            ClassApartAndRoom apart = new ClassApartAndRoom("ПАРАМЕТР НЕ НАЙДЕН", "0", ApartGroup.Name);
                            SelectedApartList.Add(apart);
                        }
                    }
                }
            }
        }
        public void SelectionApart() //метод, запускающий выбор групп
        {
            RaiseCloseRequest();
            var groupFilter = new GroupPickFilter(_doc);

            //Создаём листы: ссылки на элементы, лист групп
            IList<Autodesk.Revit.DB.Reference> refrence = new List<Autodesk.Revit.DB.Reference>();
            IList<Element> ApartListElement = new List<Element>();

            //Выбор продолжается пока пользователь не нажмёт ESC
            try
            {
                while (true)
                    refrence.Add(_uidoc.Selection.PickObject(ObjectType.Element, groupFilter, "Выберите элементы, для выхода нажмите ESC"));
            }
            catch { }

            foreach (var _ref in refrence)
            {
                Element element = _doc.GetElement(_ref);
                ApartListElement.Add(element);
            }
            
            //Нумеруем группы
            NumberApart(ApartListElement);

            //После завершения нумерации выводим список квартир
            var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue);
            apartWindow.ShowDialog();
        }
        public void SelectionRoom() //метод, запускающий поиск помещений
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
                if(roomElement.GroupId.ToString() == "-1")
                    RoomListElement.Add(roomElement);
                if (_doc.GetElement(roomElement.GroupId).LookupParameter("PNR_Функция помещения").AsString() == "Помещение управляющей компании")
                    if(!RoomListElement.Contains(_doc.GetElement(roomElement.GroupId)))
                        RoomListElement.Add(_doc.GetElement(roomElement.GroupId));
            }

            //Нумеруем помещения
            NumberRoom(RoomListElement);

            //После завершения нумерации выводим список помещений
            var room = new Room(_uiapp, _uidoc, _doc);
            room.ShowDialog();
        }
        public void NumberRoom(IList<Element> RoomListElement) //метод, нумерующий комнаты
        {
            // Параметры помещений
            String PNR_Function = string.Empty;
            String PNR_Floor = string.Empty;
            String PNR_Building = string.Empty;

            //Переменные количества помещений определённой функции
            int number_МОП = 0;
            int number_СРВ = 0;
            int number_Т= 0;
            int number_А = 0;
            int number_УК = 0;

            String PNR_Funс = string.Empty; //Ссылка на функцию помещения
            String numberRoom = string.Empty; //Итоговый номер помещения

            foreach (var room in RoomListElement)
            {
                if((BuiltInCategory)room.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups)
                {
                    _doc.GetElement(room.GroupId)
                    number_УК++;
                    Transaction tr = new Transaction(_doc, "Number room");
                    tr.Start();
                    if (PNR_Building == "-1")
                        _doc.GetElement(room.GroupId).LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Floor}.{PNR_Funс}.{number_МОП}");
                    _doc.GetElement(room.GroupId).LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_МОП}");
                    tr.Commit();
                }

                PNR_Function = room.LookupParameter("PNR_Функция помещения").AsString();
                PNR_Floor = room.LookupParameter("PNR_Этаж").AsString();
                try { PNR_Building = room.LookupParameter("ADSK_Номер здания").AsString(); }
                catch { PNR_Building = "-1"; }

                if(PNR_Function == "МОП входной группы 1 этажа" ||
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
                    
                    if(PNR_Funс == "МОП")
                    {
                        number_МОП++;
                        Transaction tr = new Transaction(_doc, "Number room");
                        tr.Start();
                        if(PNR_Building == "-1")
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Floor}.{PNR_Funс}.{number_МОП}");
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
                        room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Floor}.{PNR_Funс}.{number_Т}");
                        tr.Commit();
                    }
                }
            }
        }
        public void NumberApart(IList<Element> ApartListElement) //метод, нумерующий группы
        {
            // Параметры помещений внутри группы
            String PNR_Function = string.Empty;
            String PNR_Section = string.Empty;
            String PNR_Building = string.Empty;

            String PNR_Funс = string.Empty; //Ссылка на функцию помещения
            int number = 0; //Номер квартиры по счёту
            String numberApart = string.Empty; //Итоговый номер квартиры

            foreach (var apart in ApartListElement)
            {
                Group group = apart as Group;
                Transaction transaction = new Transaction(_doc, "UnGroup");
                transaction.Start();
                var groupRooms = group.UngroupMembers().ToList();
                transaction.RollBack();
                foreach (var room in groupRooms)
                {
                    if (room != null)
                    {
                        Element element = _doc.GetElement(room);
                        PNR_Function = element.LookupParameter("PNR_Функция помещения").AsString();
                        if (PNR_Function == "Квартиры бед доп. отделки" ||
                            PNR_Function == "Апартаменты без доп. Отделки" ||
                            PNR_Function == "Коммерческие помещения без доп. отделки" ||
                            PNR_Function == "Квартиры с отделкой" ||
                            PNR_Function == "Апартаменты с отделкой" ||
                            PNR_Function == "Коммерческие помещения с отделкой" ||
                            PNR_Function == "Помещения кладовых")
                        {
                            try { PNR_Section = element.LookupParameter("ADSK_Номер секции").AsString(); }
                            catch { PNR_Section = "-1"; }
                            try { PNR_Building = element.LookupParameter("ADSK_Номер здания").AsString(); }
                            catch { PNR_Building = "-1"; }

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
                                    PNR_Funс = "КХ";
                                    break;                                
                            }
                            
                            if(PNR_Funс == "КВ" || PNR_Funс == "КХ" || PNR_Funс == "АП")
                            {
                                number = 0;
                                //С помощью метода Count высчитываем номер квартиры в зависимости от функции
                                number = CountApart(PNR_Function, PNR_Section);
                                if (number == -1)
                                    break;
                                number++;

                                //Формируем номер квартиры 
                                if (number > 0 && number <= 9)
                                {
                                    if(PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-{PNR_Section}00{number}";
                                    if(PNR_Section == "-1")
                                        numberApart = $"{PNR_Building}/{PNR_Funс}-00{number}";
                                    if (PNR_Section == "-1" && PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-00{number}";
                                    
                                    numberApart = $"{PNR_Building}/{PNR_Funс}-{PNR_Section}00{number}";
                                }
                                
                                if (number > 9 && number <= 99)
                                {
                                    if (PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-{PNR_Section}0{number}";
                                    if (PNR_Section == "-1")
                                        numberApart = $"{PNR_Building}/{PNR_Funс}-0{number}";
                                    if (PNR_Section == "-1" && PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-0{number}";

                                    numberApart = $"{PNR_Building}/{PNR_Funс}-{PNR_Section}0{number}";
                                }

                                if (number > 99 && number <= 999)
                                {
                                    if (PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-{PNR_Section}{number}";
                                    if (PNR_Section == "-1")
                                        numberApart = $"{PNR_Building}/{PNR_Funс}-{number}";
                                    if (PNR_Section == "-1" && PNR_Building == "-1")
                                        numberApart = $"{PNR_Funс}-{number}";

                                    numberApart = $"{PNR_Building}/{PNR_Funс}-{PNR_Section}{number}";
                                }

                                number = 0;
                            }
                            
                            if (PNR_Funс == "КМ")
                            {
                                number = 0;
                                //С помощью метода Count высчитываем номер квартиры в зависимости от функции
                                number = CountApart(PNR_Function, PNR_Section);
                                if (number == -1)
                                    break;
                                number++;

                                //Формируем номер квартиры 
                                if (number > 0 && number <= 9)
                                    numberApart = $"{PNR_Building}/{PNR_Funс}-00{number}";

                                if (number > 9 && number <= 99)
                                    numberApart = $"{PNR_Building}/{PNR_Funс}-0{number}";

                                if (number > 99 && number <= 999)
                                    numberApart = $"{PNR_Building}/{PNR_Funс}-{number}";
                                number = 0;
                            }
                                                       
                            break;
                        }
                    }
                }

                //Нумеруем помещения внутри группы
                if(number != -1)
                {
                    int numberRoom = 1;
                    transaction.Start();
                    foreach (var room in groupRooms)
                    {
                        Element element = _doc.GetElement(room);
                        element.LookupParameter("PNR_Номер помещения").Set($"{numberApart}.{numberRoom}");
                        numberRoom++;
                    }
                    transaction.Commit();
                }

                //Заносим в параметр номер группы
                if(number != -1)
                {
                    transaction.Start();
                    apart.LookupParameter("ADSK_Номер квартиры").Set(numberApart);
                    transaction.Commit();
                }
            }
        }
        private int CountApart(String PNR_Function, String PNR_Section) //метод, считывающий количество групп определённой категории 
        {
            List<int> numberMax_КВ = new List<int>();
            List<int> numberMax_АП = new List<int>();
            List<int> numberMax_КМ = new List<int>();
            List<int> numberMax_КХ = new List<int>();

            foreach (var groupApart in AllGroupsApart)
            {
                //Создаём лист элементов группы
                Transaction tr = new Transaction(_doc, "UnGroup");
                tr.Start();
                var elementsGroup = groupApart.UngroupMembers().ToList();
                tr.RollBack();
                var room = _doc.GetElement(elementsGroup[0]);
                
                if (groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ") ||
                    groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("АП") ||
                    groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КМ") ||
                    groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КХ"))
                {
                    //Высчитываем максимальный номер КВ
                    if(groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                    {
                        var numberString = groupApart.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 3);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_КВ.Add(number);
                    }

                    //Высчитываем максимальный номер АП
                    if (groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("АП"))
                    {
                        var numberString = groupApart.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 3);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_АП.Add(number);
                    }

                    //Высчитываем максимальный номер КМ
                    if (groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КМ"))
                    {
                        var numberString = groupApart.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 3);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_КМ.Add(number);
                    }

                    //Высчитываем максимальный номер КХ
                    if (groupApart.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КХ"))
                    {
                        var numberString = groupApart.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 3);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_КХ.Add(number);
                    }
                }
            }

            //Возваращаем количество квартир в зависимости от функции
            if (PNR_Function == "Квартиры бед доп. отделки" || PNR_Function == "Квартиры с отделкой")
                if (numberMax_КВ.Count > 0)
                    return numberMax_КВ.Max();
                else
                    return 0;
            if (PNR_Function == "Апартаменты без доп. Отделки" || PNR_Function == "Апартаменты с отделкой")
                if (numberMax_АП.Count > 0)
                    return numberMax_АП.Max();
                else
                    return 0;
            if (PNR_Function == "Коммерческие помещения без доп. отделки" || PNR_Function == "Коммерческие помещения с отделкой")
                if (numberMax_КМ.Count > 0)
                    return numberMax_КМ.Max();
                else
                    return 0;
            if (PNR_Function == "Помещения кладовых")
                if (numberMax_КХ.Count > 0)
                    return numberMax_КХ.Max();
                else
                    return 0;
            return -1;
        }

    }
    public class RoomPickFilter : ISelectionFilter //Фильтр для комнат
    {
        public bool AllowElement(Element e)
        {
            return e is Room;
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
