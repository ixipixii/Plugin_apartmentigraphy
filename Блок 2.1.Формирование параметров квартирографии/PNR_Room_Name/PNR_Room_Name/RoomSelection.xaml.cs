using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;

namespace PNR_Room_Name
{
    /// <summary>
    /// Логика взаимодействия для RoomSelection.xaml
    /// </summary>
    public partial class RoomSelection : Window
    {
        static public string RoomName = String.Empty;
        public RoomSelection(ExternalCommandData commandData)
        {
            InitializeComponent();
            Selection selection = new Selection(commandData);
            selection.CloseRequest += (s, e) => this.Close();
            DataContext = selection;

            FunctionRoom(Selection.functionRoom);
        }

        private void LVR_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = LVR.SelectedItem;
            RoomName = selected.ToString();
        }

        private void FunctionRoom(String functionRoom)
        {
            //Считываем файл наименований помещений
            using (var package = new ExcelPackage(new System.IO.FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Имена помещений.xlsx"))))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                string range = "A2:B172";
                var rangeCells = worksheet.Cells[range];
                object[,] Allvalues = rangeCells.Value as object[,];
                if (Allvalues != null)
                {
                    int rows = Allvalues.GetLength(0);
                    int columns = Allvalues.GetLength(1);
                    int start = 0;

                    for (int i = 0; i <= rows; i++)
                    {
                        int j = 1;
                        if (Allvalues[i, j].ToString() == functionRoom)
                        {
                            start = i;
                            break;
                        }
                    }
                    for (int i = start; i < rows; i++)
                    {
                        if (Allvalues[i, 1].ToString() != functionRoom)
                            break;
                        LVR.Items.Add(Allvalues[i, 0].ToString());
                    }
                }
            }
            /*switch (functionRoom)
            {
                case "Квартиры бед доп. отделки":
                    LVR.Items.Add("Жилая комната");
                    LVR.Items.Add("Гостиная");
                    LVR.Items.Add("Столовая");
                    LVR.Items.Add("Кухня");
                    LVR.Items.Add("Кухня-ниша");
                    LVR.Items.Add("Гардеробная");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("Постирочная");
                    LVR.Items.Add("Терраса");
                    LVR.Items.Add("Кабинет");
                    LVR.Items.Add("Балкон");
                    LVR.Items.Add("Лоджия");
                    break;              
                case "Коммерческие помещения без доп. отделки":
                    LVR.Items.Add("Помещения общественного назначения");
                    LVR.Items.Add("ПОН с возможностью размещения медицинских помещений");
                    LVR.Items.Add("Кафе");
                    LVR.Items.Add("Ресторан");
                    LVR.Items.Add("Супермаркет");
                    LVR.Items.Add("Детский сад");
                    LVR.Items.Add("Фитнес-центр");
                    break;
                case "МОП входной группы 1 этажа":
                    LVR.Items.Add("Вестибюль");
                    LVR.Items.Add("Тамбур");
                    LVR.Items.Add("Лифтовой холл");
                    LVR.Items.Add("Тамбур-шлюз");
                    LVR.Items.Add("Почтовая комната");
                    LVR.Items.Add("Инвентарная (помещение хранения детского оборудования)");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("ПУИ");
                    LVR.Items.Add("Сервисный коридор");
                    break;
                case "МОП входной группы -1 этажа":
                    LVR.Items.Add("Тамбур-шлюз");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("Лифтовой холл (тамбур-шлюз)");
                    break;
                case "МОП типовых этажей":
                    LVR.Items.Add("Лифтовой холл");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("Лифтовой холл (ПБЗ)");
                    LVR.Items.Add("Тамбур-шлюз");
                    LVR.Items.Add("Лестничная клетка");
                    LVR.Items.Add("Лестничная клетка/ПБЗ");
                    LVR.Items.Add("Помещение ревизии инженерных коммуникаций");
                    LVR.Items.Add("ПУИ");
                    break;
                case "Лестницы эвакуации  (с -1го до последнего этажа)":
                    LVR.Items.Add("НЛК (наземная лестничная клетка)");
                    LVR.Items.Add("ПЛК (подземная лестничная клетка)");
                    break;
                case "Паркинг":
                    LVR.Items.Add("Инвентарная(помещение хранения велосипедов)");
                    LVR.Items.Add("Помещение автостоянки");
                    LVR.Items.Add("Рампа");
                    break;
                case "Квартиры с отделкой":
                    LVR.Items.Add("Жилая комната");
                    LVR.Items.Add("Гостиная");
                    LVR.Items.Add("Столовая");
                    LVR.Items.Add("Кухня");
                    LVR.Items.Add("Кухня-ниша");
                    LVR.Items.Add("Гардеробная");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("Постирочная");
                    LVR.Items.Add("Терраса");
                    LVR.Items.Add("Кабинет");
                    LVR.Items.Add("Балкон");
                    LVR.Items.Add("Лоджия");
                    break;
                case "Коммерческие помещения с отделкой":
                    LVR.Items.Add("Помещения общественного назначения");
                    LVR.Items.Add("ПОН с возможностью размещения медицинских помещений");
                    LVR.Items.Add("Помещение ревизии инженерных коммуникаций");
                    LVR.Items.Add("Кафе");
                    LVR.Items.Add("Ресторан");
                    LVR.Items.Add("Супермаркет");
                    LVR.Items.Add("Детский сад");
                    LVR.Items.Add("Фитнес-центр");
                    break;
                case "Помещения загрузки":
                    LVR.Items.Add("Зона / коридор загрузки");
                    LVR.Items.Add("Помещение временного хранения");
                    LVR.Items.Add("Лифтовой холл сервисного лифта");
                    break;
                case "Помещения мусороудаления":
                    LVR.Items.Add("Буферная мусоросборная камера");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("Тамбур-шлюз мусоросборной камеры");
                    LVR.Items.Add("Помещение хранения контейнеров");
                    LVR.Items.Add("Центральная мусоросборная камера");
                    break;
                case "Инженерно-технические помещения":
                    LVR.Items.Add("Помещение венткамеры ОВ (принадлежность обслуживаемой зоны: автостоянка,ТП, ДОО, холодильного центра, ИТП...)");
                    LVR.Items.Add("Помещение венткамеры ПДВ (принадлежность обслуживаемой зоны: автостоянки, ТП, ДОО, холодильного центра, ИТП...)");
                    LVR.Items.Add("Насосная");
                    LVR.Items.Add("Насосная АУПТ");
                    LVR.Items.Add("Кроссовая (узел, связи, СС, А/с, АР1, помещение домофонной сети и пр.)");
                    LVR.Items.Add("Помещение ввода СС");
                    LVR.Items.Add("ИТП");
                    LVR.Items.Add("ВРУ (принадлежность обслуживаемой зоны: ТП, ДОО, холодильного центра, ИТП...)");
                    LVR.Items.Add("ВРУ жилья");
                    LVR.Items.Add("ВРУ автостоянки");
                    LVR.Items.Add("ВРУ ПОН");
                    LVR.Items.Add("ТП");
                    LVR.Items.Add("Трансформаторная камера");
                    LVR.Items.Add("РУНН");
                    LVR.Items.Add("РУВН");
                    LVR.Items.Add("Помещение кабельного ввода");
                    LVR.Items.Add("Холодильный центр");
                    LVR.Items.Add("Водомерный узел");
                    LVR.Items.Add("Серверная");
                    LVR.Items.Add("Помещение выпуска (ВК, ливневка и т.д)");
                    LVR.Items.Add("Помещение очистки воды");
                    LVR.Items.Add("КНС");
                    LVR.Items.Add("Помещение ревизии инженерных коммуникаций");
                    LVR.Items.Add("Техническое пространство");
                    break;
                case "Помещение управляющей компании":
                    LVR.Items.Add("Тамбур");
                    LVR.Items.Add("Помещение управляющей компании");
                    LVR.Items.Add("Вестибюль");
                    LVR.Items.Add("Переговорная комната для приема населения");
                    LVR.Items.Add("Переговорная комната рабочая");
                    LVR.Items.Add("Рабочее место менеджера по работе с клиентами");
                    LVR.Items.Add("Кабинет управляющего");
                    LVR.Items.Add("Зал менеджеров");
                    LVR.Items.Add("Комната для отдыха и приема пищи");
                    LVR.Items.Add("Архив");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("ПУИ");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("НКЛ (наземная лестничная клетка)");
                    break;
                case "Объединенный диспетчерский пункт":
                    LVR.Items.Add("Объединенный диспетчерский пункт");
                    LVR.Items.Add("Тамбур");
                    LVR.Items.Add("Кроссовая");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("ПУИ");
                    break;
                case "Помещения линейного и обслуживающего персонала":
                    LVR.Items.Add("Служебное помещение линейного персонала");
                    LVR.Items.Add("Помещение для размещения круглосуточного дежурного техника");
                    LVR.Items.Add("Комната приема пищи линейного персонала");
                    LVR.Items.Add("Служебное помещение обслуживающего персонала");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("Помещение для хранения расходных материалов и запчастей");
                    LVR.Items.Add("Помещение для временного хранения строительных и отделочных материалов");
                    LVR.Items.Add("Кладовая (уличная мебель)");
                    LVR.Items.Add("Душевая");
                    LVR.Items.Add("Коридор");
                    break;
                case "Помещения охраны":
                    LVR.Items.Add("Служебное помещение сотрудников охраны");
                    LVR.Items.Add("Кабинет руководителя службы охраны");
                    LVR.Items.Add("Комната приема пищи сотрудников охраны");
                    LVR.Items.Add("Душевая");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("Коридор");
                    break;
                case "Помещения Клининговых служб":
                    LVR.Items.Add("Служебное помещение Клининга");
                    LVR.Items.Add("Кабинет руководителя службы Клининга");
                    LVR.Items.Add("Комната для отдыха приема пищи сотрудников Клининга");
                    LVR.Items.Add("Помещение для хранения инвентаря и подзарядки уборочных машин");
                    LVR.Items.Add("Помещение для хранения запаса расходных материалов и бытовой химии");
                    LVR.Items.Add("Помещение для хранения расходных материалов для территории");
                    LVR.Items.Add("Душевая");
                    LVR.Items.Add("Санузел");
                    LVR.Items.Add("Коридор");
                    LVR.Items.Add("ПУИ");
                    break;
                case "Помещения кладовых":
                    LVR.Items.Add("Проход блока кладовых");
                    LVR.Items.Add("Кладовая");
                    break;
            }*/           
        }
    }
}
