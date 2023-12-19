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
        public List<Element> AllRoomsRenumber = new List<Element>(); //Квартиры для переименовки
        public String _SelectedSectionValue { get; set; } //Выбранная секция в ComboBox
        private string text; //
        public int index = 0;
        public string Text
        {
            get => text;
            set
            {
                text = value;
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
            if(v == 2)
                SelectCommandApart = new DelegateCommand(SelectionApart_2);
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
            CloseRequest?.Invoke(this, EventArgs.Empty);        }
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
            foreach(var room in AllRooms)
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
                                                        ApartGroup.Name, ApartGroup.Id);
                            SelectedApartList.Add(apart);
                        }
                        catch
                        {
                            ClassApartAndRoom apart = new ClassApartAndRoom("ПАРАМЕТР НЕ НАЙДЕН", "0", ApartGroup.Name, ApartGroup.Id);
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
            var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 1);
            apartWindow.ShowDialog();
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
                if(roomElement.GroupId.ToString() == "-1")
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
            int number_Т= 0;
            int number_А = 0;
            int number_УК = 0;

            String PNR_Funс = string.Empty; //Ссылка на функцию помещения
            String numberRoom = string.Empty; //Итоговый номер помещения

            foreach (var room in RoomListElement)
            {
                if((BuiltInCategory)room.Category.Id.IntegerValue == BuiltInCategory.OST_IOSModelGroups)
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
                        foreach(var room_ in roomInGroup)
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

        //version 2
        public void SelectionApart_2()
        {
            RaiseCloseRequest();

            List<int> numberMax_КВ = new List<int>();
            List<int> numberMax_АП = new List<int>();
            List<int> numberMax_КМ = new List<int>();
            List<int> numberMax_КХ = new List<int>();
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
                try 
                { 
                    PNR_Section = apart.LookupParameter("ADSK_Номер секции").AsString();
                    if (PNR_Section == "")
                        PNR_Section = "-1";
                }
                catch { PNR_Section = "-1"; }
                try 
                { 
                    PNR_Building = apart.LookupParameter("ADSK_Номер здания").AsString(); 
                    if(PNR_Building == "")
                        PNR_Building= "-1";
                }
                catch { PNR_Building = "-1"; }
            }

            //Фильтр по всем помещениям на всей модели
            var AllRooms = new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(r => r.LookupParameter("ADSK_Номер секции").AsString() == PNR_Section || r.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor)
                .ToList();

            List<Element> AllRoomsWGroup = new List<Element> ();
            foreach(var room in AllRooms)
            {
                if(room.GroupId.IntegerValue == -1)
                    AllRoomsWGroup.Add(room);
            }

            //Высчитываем максимальный номер
            foreach (var room in AllRoomsWGroup)
            {
                //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО СЕКЦИИ ДЛЯ КВ
                if(room.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                    {
                        continue;
                    }
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                    {
                        var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
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
                }

                //ВЫСЧИТЫВАЕМ МАКСИМАЛЬНЫЙ НОМЕР ПО ЭТАЖУ ДЛЯ АП, КМ, КХ
                if (room.LookupParameter("ADSK_Этаж").AsString() == PNR_Floor)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString() == null || room.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                    {
                        continue;
                    }

                    //Высчитываем максимальный номер АП
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("АП"))
                    {
                        var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 2);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_АП.Add(number);
                    }

                    //Высчитываем максимальный номер КМ
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КМ"))
                    {
                        var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 2);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_КМ.Add(number);
                    }

                //Высчитываем максимальный номер КХ
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КХ"))
                    {
                        var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
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
                        PNR_Funс = "КХ";
                        break;
                    
                    case "Помещение управляющей компании":
                        PNR_Funс = "УК";
                        break;
                }
            }

            //Формируем ADSK_Номер квартиры для УК
            if (PNR_Funс == "УК")
            {
                var AllRoomsУК = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                                .WhereElementIsNotElementType()
                                .OfCategory(BuiltInCategory.OST_Rooms)
                                .ToList();

                //Высчитываем максимальный номер УК
                foreach(var room in AllRoomsУК)
                {
                    if (room.LookupParameter("ADSK_Номер квартиры").AsString().Contains("КВ"))
                    {
                        var numberString = room.LookupParameter("ADSK_Номер квартиры").AsString();
                        var numberStringArray = numberString.ToCharArray();
                        Array.Reverse(numberStringArray);
                        var numberStringReverse = new string(numberStringArray);
                        var numberRoomReverse = numberStringReverse.Substring(0, 3);
                        var numberRoomArray = numberRoomReverse.ToCharArray();
                        Array.Reverse(numberRoomArray);
                        var numberRoom = new string(numberRoomArray);
                        var numberStr = numberRoom.TrimStart('0');
                        var number = int.Parse(numberStr);
                        numberMax_УК.Add(number);
                    }
                }
            }

            //Переменная-индекс для нумерации помещений
            int i = 1;
            Transaction tr = new Transaction(_doc, "Заносим параметр");
            tr.Start();
            //Нумеруем помещения
            foreach (var room in ApartListElement)
            {
                int number = 0;
                if (PNR_Funс == "КВ")
                {
                    try { number = numberMax_КВ.Max(); number++; }
                    catch { number = 1; }
                }

                if (PNR_Funс == "КХ")
                {
                    try { number = numberMax_КХ.Max(); number++; }
                    catch { number = 1; }
                }

                if (PNR_Funс == "АП")
                {
                    try { number = numberMax_АП.Max(); number++; }
                    catch { number = 1; }
                }

                if (PNR_Funс == "КМ")
                {
                    try { number = numberMax_КМ.Max(); number++; }
                    catch { number = 1; }
                }

                //Формируем номер КВ, КХ, АП
                if (PNR_Funс == "КВ" || PNR_Funс == "КХ" || PNR_Funс == "АП")
                { 
                    if (number > 0 && number <= 9)
                    {
                        if (PNR_Building == "-1" && PNR_Section != "-1")
                        {
                            if(PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Section}00{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Section}00{number}.{i}");
                            }
                            
                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Floor}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Floor}0{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if(i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}{i}");
                            }
                        }
                        if (PNR_Building != "-1" && PNR_Section == "-1")
                        {
                            if(PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/00{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/00{number}.{i}");
                            }
                            
                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}{i}");
                            }
                        }
                        if (PNR_Section == "-1" && PNR_Building == "-1")
                        {
                            if(PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-00{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-00{number}.{i}");
                            }
                            
                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-0{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}{i}");
                            }
                        }
                        if (PNR_Section != "-1" && PNR_Building != "-1")
                        {
                            if(PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}00{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}00{number}.{i}");
                            }
                            
                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}{i}");
                            }
                        }
                    }
                    if (number > 9 && number <= 99)
                    {
                        if (PNR_Building == "-1" && PNR_Section != "-1")
                        {
                            if(PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Section}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Section}0{number}.{i}");
                            }
                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Floor}{number}.{i}");
                            }
                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}{i}");
                            }
                        }
                        if (PNR_Building != "-1" && PNR_Section == "-1")
                        {
                            if (PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/0{number}.{i}");
                            }

                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{number}{i}");
                            }

                        }
                        if (PNR_Section == "-1" && PNR_Building == "-1")
                        {
                            if (PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-0{number}.{i}");
                            }

                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Floor}{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}{i}");
                            }
                        }
                        if (PNR_Section != "-1" && PNR_Building != "-1")
                        {
                            if (PNR_Funс == "КВ") 
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}0{number}.{i}");
                            }

                            if (PNR_Funс == "АП")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}.{i}");
                            }

                            if (PNR_Funс == "КХ")
                            {
                                if (i > 0 && i <= 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{number}0{i}");
                                if (i > 9)
                                    room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{number}{i}");
                            }

                        }
                    }
                    if (number > 99 && number <= 999)
                    {
                        if (PNR_Building == "-1" && PNR_Section != "-1")
                        {
                            if (PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Section}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Section}{number}.{i}");
                            }
                        }
                        if (PNR_Building != "-1" && PNR_Section == "-1")
                        {
                            if (PNR_Funс == "КВ")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Building}/{PNR_Funс}-{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Building}/{PNR_Funс}-{number}.{i}");
                            }
                        }
                        if (PNR_Section == "-1" && PNR_Building == "-1")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}");
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{number}.{i}");
                        }
                        if (PNR_Section != "-1" && PNR_Building != "-1")
                        {
                            room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}{number}");
                            room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Section}{number}.{i}");
                        }
                    }                    
                }

                if (PNR_Funс == "КМ")
                {
                    //Формируем номер КМ 
                    if (number > 0 && number <= 9)
                    {
                        if(room.LookupParameter("ADSK_Этаж").AsString() == "1")
                        {
                            if(PNR_Building == "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-0{number}.{i}");

                            }
                            if(PNR_Building != "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/0{number}.{i}");
                            }
                        }
                        if (room.LookupParameter("ADSK_Этаж").AsString() != "1")
                        {
                            if(PNR_Building == "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Floor}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Floor}0{number}.{i}");
                            }                           
                            if(PNR_Building != "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}0{number}.{i}");
                            }
                        }
                    }
                    if (number > 9 && number <= 99)
                    {
                        if (room.LookupParameter("ADSK_Этаж").AsString() == "1")
                        {
                            if (PNR_Building == "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{number}.{i}");
                            }
                            if(PNR_Building != "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{number}.{i}");
                            }
                        }
                        if (room.LookupParameter("ADSK_Этаж").AsString() != "1")
                        {
                            if (PNR_Building == "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Floor}{number}.{i}");
                            }
                            if (PNR_Building != "-1")
                            {
                                room.LookupParameter("ADSK_Номер квартиры").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}");
                                room.LookupParameter("PNR_Номер помещения").Set($"{PNR_Funс}-{PNR_Building}/{PNR_Floor}{number}.{i}");
                            }
                        }
                    }
                }

                number = 0;
                i++;
            }
            tr.Commit();

            var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
            apartWindow.ShowDialog();
        }
        public void ApartList_2()
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
                    if(room.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
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
                    if(_SelectedSectionValue == "-1")
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
            {}
            catch(Exception e)
            {
                TaskDialog.Show("Ошибка", $"Выберите помещение");
                var apartWindow1 = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                apartWindow1.ShowDialog();
            }

            //Фильтр по всем квартирам - помещениям на данном виде
            AllRoomsRenumber = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() == SelectedRoom.ADSK_Номер_квартиры)
                .ToList();

            var renumber = new Renumber(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2, AllRoomsRenumber, 0);
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
            AllRoomsRenumber = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Where(g => g.Name == SelectedRoom.name)
                .ToList();

            var renumber = new Renumber(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2, AllRoomsRenumber, 1);
            renumber.ShowDialog();
        }
        public void EnterRenumber()
        {
            RaiseCloseRequest();

            if(index == 0) //Переименовываем квартиры
            {
                int length = 0;
                string ADSK_Номер_квартиры = string.Empty;
                foreach (var a in AllRoomsRenumber)
                {
                    length = a.LookupParameter("ADSK_Номер квартиры").AsString().Length;
                    break;
                }

                var rooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                                    .WhereElementIsNotElementType()
                                    .OfCategory(BuiltInCategory.OST_Rooms)
                                    .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != null)
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0') == text)
                                    .ToList();

                if (rooms.Count > 0)
                {
                    foreach (var room in rooms)
                    {
                        AllRoomsRenumber.Add(room);
                    }
                }

                Transaction tr = new Transaction(_doc, "Renumber");
                int i = 1;
                tr.Start();
                foreach (var room in AllRoomsRenumber)
                {
                    string str = room.LookupParameter("ADSK_Номер квартиры").AsString();
                    if (str.Length > 0)
                        str = str.Remove(str.Length - 3);
                    if (text.Length == 1)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + "00" + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + "00" + text + "." + i);
                    }
                    if (text.Length == 2)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + "0" + text + "." + i);
                    }
                    if (text.Length == 3)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + text + "." + i);
                    }
                    i++;
                }
                i = 0;
                tr.Commit();
                var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                apartWindow.ShowDialog();
            }
                        
            if (index == 1) //Переименовываем помещения
            {
                int length = 0;
                string ADSK_Номер_квартиры = string.Empty;
                var selectedRoom = _doc.GetElement(SelectedRoom.id);

                if (selectedRoom != null) { length = selectedRoom.LookupParameter("ADSK_Номер квартиры").AsString().Length; }

                var rooms = new FilteredElementCollector(_doc, _uidoc.ActiveView.Id)
                                    .WhereElementIsNotElementType()
                                    .OfCategory(BuiltInCategory.OST_Rooms)
                                    .Where(g => g.LookupParameter("ADSK_Номер секции").AsString() == _SelectedSectionValue)
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Length == length)
                                    .Where(g => g.LookupParameter("ADSK_Номер квартиры").AsString().Substring(length - 3, 3).Trim('0') == text)
                                    .ToList();

                if (rooms.Count > 0)
                {
                    foreach (var room in rooms)
                    {
                        AllRoomsRenumber.Add(room);
                    }
                }

                Transaction tr = new Transaction(_doc, "Renumber");
                int i = 1;
                tr.Start();
                foreach (var room in AllRoomsRenumber)
                {
                    string str = room.LookupParameter("ADSK_Номер квартиры").AsString();
                    if (str.Length > 0)
                        str = str.Remove(str.Length - 3);
                    if (text.Length == 1)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + "00" + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + "00" + text + "." + i);
                    }
                    if (text.Length == 2)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + "0" + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + "0" + text + "." + i);
                    }
                    if (text.Length == 3)
                    {
                        room.LookupParameter("ADSK_Номер квартиры").Set(str + text);
                        room.LookupParameter("PNR_Номер помещения").Set(str + text + "." + i);
                    }
                    i++;
                }
                i = 0;
                tr.Commit();
                var apartWindow = new Apart(_uiapp, _uidoc, _doc, _SelectedSectionValue, 2);
                apartWindow.ShowDialog();
            }                      
        }
    }
    public class RoomPickFilter : ISelectionFilter //Фильтр для комнат
    {
        public bool AllowElement(Element e)
        {
            return (e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Rooms));
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
