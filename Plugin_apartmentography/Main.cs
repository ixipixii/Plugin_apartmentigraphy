using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Windows;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Media.Imaging;
using System.Reflection;

namespace Plugin_apartmentography
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public string GetExeDirectory()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = Path.GetDirectoryName(path);
            return path;
        }
        public Result OnStartup(UIControlledApplication application)
        {

            string tabName = "Квартирография";
            application.CreateRibbonTab(tabName);

            string absPath = GetExeDirectory();

            string relPath1 = @"1\";
            string path1 = Path.Combine(absPath, relPath1);
            path1 = Path.GetFullPath(path1);

            string pathImg1 = Path.Combine(absPath, @"Resources\pm.png");

            string relPath2 = @"2\";
            string path2 = Path.Combine(absPath, relPath2);
            path2 = Path.GetFullPath(path2);

            string pathImg2 = Path.Combine(absPath, @"Resources\kv.png");

            string relPath3 = @"3\";
            string path3 = Path.Combine(absPath, relPath3);
            path3 = Path.GetFullPath(path3);

            string pathImg3 = Path.Combine(absPath, @"Resources\kv.png");

            string relPath4 = @"4\";
            string path4 = Path.Combine(absPath, relPath4);
            path4 = Path.GetFullPath(path4);

            string pathImg4 = Path.Combine(absPath, @"Resources\kv.png");

            var panel = application.CreateRibbonPanel(tabName, "Создание");

            var button_1 = new PushButtonData("Имя помещения", "Имя помещения",
                Path.Combine(path1, "Plugin_Kvartiry.dll"),
                "Plugin_Kvartiry.Main");

            var button_2 = new PushButtonData("Квартиры", "Создание квартир",
                Path.Combine(path2, "Plugin_Kvartiry2.dll"),
                "Plugin_Kvartiry2.Main");

            var button_3 = new PushButtonData("Этаж", "Этаж",
                Path.Combine(path3, "ADSK_Floor.dll"),
                "ADSK_Floor.Main");

            var button_4 = new PushButtonData("Секция", "Секция",
                Path.Combine(path4, "ADSK_Section.dll"),
                "ADSK_Section.Main");

/*            Uri uriImage1 = new Uri(pathImg1, UriKind.Absolute);
            BitmapImage largeImage1 = new BitmapImage(uriImage1);
            button_1.LargeImage = largeImage1;

            Uri uriImage2 = new Uri(pathImg2, UriKind.Absolute);
            BitmapImage largeImage2 = new BitmapImage(uriImage2);
            button_2.LargeImage = largeImage2;

            Uri uriImage3 = new Uri(pathImg3, UriKind.Absolute);
            BitmapImage largeImage3 = new BitmapImage(uriImage3);
            button_3.LargeImage = largeImage3;

            Uri uriImage4 = new Uri(pathImg4, UriKind.Absolute);
            BitmapImage largeImage4 = new BitmapImage(uriImage4);
            button_4.LargeImage = largeImage4;*/

            panel.AddItem(button_1);
            panel.AddItem(button_2);
            panel.AddItem(button_3);
            panel.AddItem(button_4);

            return Result.Succeeded;
        }
    }
}