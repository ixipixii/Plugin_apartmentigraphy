using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnloadingPlan
{
    [Transaction(TransactionMode.Manual)]
    internal class MainPDF : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            Transaction tr = new Transaction(doc, "Export view in PDF");
            tr.Start();
            var saveDialogPdf = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "name.pdf",
                DefaultExt = ".pdf"
            };

            string selectedFilePathpDF = string.Empty;

            if (saveDialogPdf.ShowDialog() == DialogResult.OK)
            {
                selectedFilePathpDF = saveDialogPdf.FileName;
            }

            String fileName = selectedFilePathpDF.Substring(selectedFilePathpDF.LastIndexOf('\\'));
            String fileNameNew = fileName.Trim('\\').Remove(fileName.LastIndexOf('.') - 1);
            String folderName = selectedFilePathpDF.Trim(fileName.ToCharArray());

            PDFExportOptions opt = new PDFExportOptions
            {
                Combine = true,
                FileName = fileNameNew,
                ZoomType = (ZoomType)ZoomFitType.FitToPage,
                PaperFormat = ExportPaperFormat.ARCH_E,
            };

            List<ElementId> elementIds = new List<ElementId> { doc.ActiveView.Id };
            doc.Export(folderName, elementIds, opt);
            tr.RollBack();
            return Result.Succeeded;
        }
    }
}
