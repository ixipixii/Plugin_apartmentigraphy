using Autodesk.Revit.DB;
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
        public DelegateCommand SelectCommand { get; } //Делегат кнопки "Пронумеровать"
        public DelegateCommand SelectApart { get; } //Делегат кнопки "Нумерация групп помещений"
        public DelegateCommand SelectRoom { get; } //Делегат кнопки "Нумерация помещений"
        public List<ClassApartAndRoom> SelectedApartList { get; set; } = new List<ClassApartAndRoom>(); //Вспомогательный лист для ListView 
        public String _SelectedSectionValue { get; set; } //выбранная секция в ComboBox

        //Переменные количества групп определённой функции
        int number_КВ = 0;
        int number_АП = 0;
        int number_КМ = 0;
        int number_КХ = 0;

        public NumberSelection(UIApplication uiapp, UIDocument uidoc, Document doc, string SelectedSectionValue)
        {
            _uiapp = uiapp;
            _uidoc = uidoc;
            _doc = doc;
            SelectCommand = new DelegateCommand(Selection);
            SelectApart = new DelegateCommand(SelectAparts);
            SelectRoom = new DelegateCommand(SelectRooms);
            _SelectedSectionValue = SelectedSectionValue;
        }
        private void SelectRooms()
        {
            RaiseCloseRequest();
            TaskDialog.Show("a", "a");
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
                            ClassApartAndRoom apart = new ClassApartAndRoom(ApartGroup,
                                                        ApartGroup.LookupParameter("ADSK_Номер квартиры").AsString(),
                                                        "0",
                                                        ApartGroup.Name);
                            SelectedApartList.Add(apart);
                        }
                        catch
                        {
                            ClassApartAndRoom apart = new ClassApartAndRoom(ApartGroup, "ПАРАМЕТР НЕ НАЙДЕН", "0", ApartGroup.Name);
                            SelectedApartList.Add(apart);
                        }
                    }
                }
            }
        }
        public void Selection() //метод, запускающий выбор групп
        {
            RaiseCloseRequest();

            var roomFilter = new RoomPickFilter();
            var groupFilter = new GroupPickFilter(_doc);

            //Создаём листы: ссылки на элементы, лист групп, лист помещений
            IList<Autodesk.Revit.DB.Reference> refrence = new List<Autodesk.Revit.DB.Reference>();
            IList<Element> ApartListElement = new List<Element>();
            IList<Element> RoomListElement = new List<Element>();

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
        public void NumberRoom(IList<Element> RoomListElement) //метод, нумерующий комнаты
        {
            foreach (var element in RoomListElement)
            {
                String roomFunction = element.LookupParameter("PNR_Функция помещения").AsString();
                String roomSection = element.LookupParameter("ADSK_Номер секции").AsString();
                String roomBuilding = element.LookupParameter("ADSK_Номер здания").AsString();
            }
        }
        public void NumberApart(IList<Element> ApartListElement) //метод, нумерующий квартиры
        {
            // Параметры помещений внутри квартиры
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
                            PNR_Section = element.LookupParameter("ADSK_Номер секции").AsString();
                            PNR_Building = element.LookupParameter("ADSK_Номер здания").AsString();

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
                            number = 0;
                            number = CountApart(PNR_Function); //С помощью метода Count высчитываем номер квартиры в зависимости от функции
                            number++;
                            numberApart = $"{PNR_Building}/{PNR_Funс}-{PNR_Section}{number}"; //Формируем номер квартиры в зависимости от функции
                            number = 0;
                            
                            number_КВ = 0;
                            number_АП = 0;
                            number_КМ = 0;
                            number_КХ = 0;
                                                       
                            break;
                        }
                    }
                }
                //Формируем номер группы
                transaction.Start();
                if(apart.LookupParameter("ADSK_Номер квартиры").AsString() == "")
                    apart.LookupParameter("ADSK_Номер квартиры").Set(numberApart);
                transaction.Commit();
            }
        }
        private int CountApart(String PNR_Function) //метод, считывающий количество групп определённой категории 
        {
            foreach (var groupApart in AllGroupsApart)
            {
                //Создаём лист элементов группы
                Transaction tr = new Transaction(_doc, "UnGroup");
                tr.Start();
                var elementsGroup = groupApart.UngroupMembers().ToList();
                tr.RollBack();
                var room = _doc.GetElement(elementsGroup[0]);

                //Если в номере группы есть данные, увеличиваем количество нумерованных групп
                //В обратном случае количество = 0
                if(groupApart.LookupParameter("ADSK_Номер квартиры").AsString() != "")
                {
                    switch (room.LookupParameter("PNR_Функция помещения").AsString())
                    {
                        case "Квартиры бед доп. отделки":
                            number_КВ++;
                            break;

                        case "Квартиры с отделкой":
                            number_КВ++;
                            break;

                        case "Апартаменты без доп. Отделки":
                            number_АП++;
                            break;

                        case "Апартаменты с отделкой":
                            number_АП++;
                            break;

                        case "Коммерческие помещения без доп. отделки":
                            number_КМ++;
                            break;

                        case "Коммерческие помещения с отделкой":
                            number_КМ++;
                            break;

                        case "Помещения кладовых":
                            number_КХ++;
                            break;
                    }
                }
            } 
            //Возваращаем количество квартир в зависимости от функции
            if (PNR_Function == "Квартиры бед доп. отделки" || PNR_Function == "Квартиры с отделкой")
                return number_КВ;
            if (PNR_Function == "Апартаменты без доп. Отделки" || PNR_Function == "Апартаменты с отделкой")
                return number_АП;
            if (PNR_Function == "Коммерческие помещения без доп. отделки" || PNR_Function == "Коммерческие помещения с отделкой")
                return number_КМ;
            if (PNR_Function == "Помещения кладовых")
                return number_КХ;
            return 0;
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
