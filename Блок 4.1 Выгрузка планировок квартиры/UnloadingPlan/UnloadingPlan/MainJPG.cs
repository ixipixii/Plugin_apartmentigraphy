using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;

namespace UnloadingPlan
{
    [Transaction(TransactionMode.Manual)]
    public class MainJPG : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            Transaction tr = new Transaction(doc, "Export view in JPG");
            tr.Start();

            var saveDialogImg = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "name.jpg",
                DefaultExt = ".jpg"
            };

            string selectedFilePathiMG = string.Empty;

            if (saveDialogImg.ShowDialog() == DialogResult.OK)
            {
                selectedFilePathiMG = saveDialogImg.FileName;
            }

            Autodesk.Revit.DB.View view = doc.ActiveView;

            IList<ElementId> viewIds = new List<ElementId> { view.Id };

            ImageExportOptions options = new ImageExportOptions
            {
                //ExportRange = ExportRange.VisibleRegionOfCurrentView, //Диапазон экспорта, определяющий, какие представления будут экспортированы=Экспортируйте набор представлений (набор в ViewsAndSheets).
                ZoomType = ZoomFitType.FitToPage, //Подогнать весь вид к определенному размеру изображения.
                PixelSize = 1200, //Размер изображения в пикселях в одном направлении. Используется, только если ZoomType равен FitToPage.
                FilePath = selectedFilePathiMG, // путь хранения
                FitDirection = FitDirectionType.Horizontal, //Подходящее направление. Используется, только если ZoomType равен FitToPage
                HLRandWFViewsFileType = ImageFileType.JPEGLossless, //Тип файла для экспортированных видов HLR и каркаса.
                ShadowViewsFileType = ImageFileType.JPEGLossless, //Тип файла для экспортированных теневых видов.
                ImageResolution = ImageResolution.DPI_600, //Разрешение изображения в точках на дюйм.
                ExportRange = ExportRange.CurrentView,
            };

            doc.ExportImage(options);
            tr.RollBack();
            return Result.Succeeded;
        }

    }
}