using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Plugin_Kvartiry2
{
    public class Selection
    {
        private ExternalCommandData _commandData;
        public static List<Element> rooms = new List<Element>();
        public static Group group;
        public static String section = String.Empty;
        public static String level = String.Empty;
        public static List<String> index = new List<String>();
        public DelegateCommand SelectionLevel { get; }
        public DelegateCommand Appoint { get; }
        public DelegateCommand Сontinue { get; }

        public Selection(ExternalCommandData commandData)
        {
            _commandData = commandData;
            SelectionLevel = new DelegateCommand(OnSelectionLevel);
            Appoint = new DelegateCommand(OnAppoint);
            Сontinue = new DelegateCommand(OnContinue);
        }

        private void OnContinue()
        {
            RaiseCloseRequest();
            Select();
            var window = new CreatingApartment(_commandData);
            window.ShowDialog();
        }

        private void OnAppoint()
        {
            List<ElementId> elementsId = new List<ElementId>();
            foreach(var element in rooms)
            {
                elementsId.Add(element.Id);
            }

            Transaction transaction = new Transaction(_commandData.Application.ActiveUIDocument.Document, "Create group");
            transaction.Start();
            group = _commandData.Application.ActiveUIDocument.Document.Create.NewGroup(elementsId);

            if (group.LookupParameter("ADSK_Номер квартиры") == null)
            {
                CreateParameter();
            }
            
            if(index.Count > 0)
                group.LookupParameter("ADSK_Номер квартиры").Set($"{section}.{level}.{index[index.Count + 1]}");
            else
                group.LookupParameter("ADSK_Номер квартиры").Set($"{section}.{level}.{1}");

            rooms.Clear();
            elementsId.Clear();
            transaction.Commit();
        }

        private void CreateParameter()
        {
            DefinitionFile definitionFile = _commandData.Application.Application.OpenSharedParameterFile();
            if (definitionFile == null)
                TaskDialog.Show("Ошибка", "Не найден ФОП");

            Definition definition = definitionFile.Groups.SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals("ADSK_Номер квартиры"));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            CategorySet categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(_commandData.Application.ActiveUIDocument.Document, BuiltInCategory.OST_IOSModelGroups));

            Binding binding = _commandData.Application.Application.Create.NewInstanceBinding(categorySet);

            BindingMap map = _commandData.Application.ActiveUIDocument.Document.ParameterBindings;
            map.Insert(definition, binding, BuiltInParameterGroup.PG_IDENTITY_DATA);
        }

        private void OnSelectionLevel()
        {
            RaiseCloseRequest();
            Select();
            var window = new CreatingApartment(_commandData);
            window.ShowDialog();
        }

        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }

        public string GetExeDirectory()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);
            return path;
        }

        public void Select()
        {
            var uiapp = _commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Document doc = _commandData.Application.ActiveUIDocument.Document;

            uidoc.ActiveView = LevelSelection.selectedLevel;

            var roomFilter = new GroupPickFilter();
            var selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Выберите помещения");

            foreach(var selectedElement in selectedRef)
            {
                rooms.Add(doc.GetElement(selectedElement));
            }

            section = rooms[0].LookupParameter("ADSK_Корпус").AsString();
            level = rooms[0].LookupParameter("ADSK_Этаж").AsString();

            //Чтение файла
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "All files(*.*)|*.*";

            string filePath = GetExeDirectory();

            if (filePath is null)
                return;

            int rowIndex = 0; //количество строк в файле
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath);
                ISheet sheet = workbook.GetSheetAt(index: 0);

                int lastRowNum = sheet.LastRowNum; //последняя строка в файле

                if(lastRowNum != -1) //если -1, значит файл пустой. В другом случае считываем номера квартир
                {
                    while (sheet.GetRow(rowIndex) != null)
                    {
                        if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                            sheet.GetRow(rowIndex).GetCell(1) == null ||
                            sheet.GetRow(rowIndex).GetCell(2) == null)
                        {
                            rowIndex++;
                            continue;
                        }

                        //Считываем номера квартир 
                        index.Add(sheet.GetRow(rowIndex).GetCell(0).StringCellValue);
                        rowIndex++;
                    }
                }
            }

            List<Element> apartment = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();

            //Запись параметров в файл
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelPath = Path.Combine(desktopPath, "База_данных_квартир.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.GetSheetAt(0);

                foreach (var apart in apartment)
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, index);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, level);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, section);
                    rowIndex++;
                }

                workbook.Write(stream, false);
            }

            System.Diagnostics.Process.Start(excelPath);
        }

        public class GroupPickFilter : ISelectionFilter
        {
            public bool AllowElement(Element e)
            {
                return e is Room;
            }
            public bool AllowReference(Reference r, XYZ p)
            {
                return false;
            }

        }
    }
}
