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
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            /*Transaction tr = new Transaction(doc, "g");
            tr.Start();

            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "paint.jpg",
                DefaultExt = ".jpg"
            };

            string selectedFilePath = string.Empty;

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }

            Autodesk.Revit.DB.View view = doc.ActiveView; 

            IList<ElementId> viewIds = new List<ElementId> { view.Id }; 

            ImageExportOptions options = new ImageExportOptions 
            {
                //ExportRange = ExportRange.VisibleRegionOfCurrentView, //Диапазон экспорта, определяющий, какие представления будут экспортированы=Экспортируйте набор представлений (набор в ViewsAndSheets).
                ZoomType = ZoomFitType.FitToPage, //Подогнать весь вид к определенному размеру изображения.
                PixelSize = 1200, //Размер изображения в пикселях в одном направлении. Используется, только если ZoomType равен FitToPage.
                FilePath = selectedFilePath, // путь хранения
                FitDirection = FitDirectionType.Horizontal, //Подходящее направление. Используется, только если ZoomType равен FitToPage
                HLRandWFViewsFileType = ImageFileType.JPEGLossless, //Тип файла для экспортированных видов HLR и каркаса.
                ShadowViewsFileType = ImageFileType.JPEGLossless, //Тип файла для экспортированных теневых видов.
                ImageResolution = ImageResolution.DPI_600, //Разрешение изображения в точках на дюйм.
                ExportRange = ExportRange.CurrentView,
            };           
            doc.ExportImage(options);
            tr.RollBack();*/

            Transaction tr = new Transaction(doc, "new list");
            tr.Start();

            ViewSheet viewSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);

            UV location = new UV((viewSheet.Outline.Max.U - viewSheet.Outline.Min.U) / 2,
                                    (viewSheet.Outline.Max.V - viewSheet.Outline.Min.V) / 2);
            
            Viewport.Create(doc, viewSheet.Id, doc.ActiveView.Id, new XYZ(location.U, location.V, 0));

            

            tr.Commit();

            return Result.Succeeded;
        }

        /*public void PDFExport(Document doc)
        {
            Transaction tr = new Transaction(doc);
            PDFExportOptions opt = new PDFExportOptions();
            opt.Combine = true;
            opt.FileName = "My House";
            opt.PaperFormat = ExportPaperFormat.ARCH_E;
            doc.Export(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .ToElementIds().ToList(),
                opt);
        }*/


    }
}