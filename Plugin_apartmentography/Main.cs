﻿using Autodesk.Revit.ApplicationServices;
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

            string pathImg1 = Path.Combine(absPath, @"Resources\этажи.png");

            string relPath2 = @"2\";
            string path2 = Path.Combine(absPath, relPath2);
            path2 = Path.GetFullPath(path2);

            string pathImg2 = Path.Combine(absPath, @"Resources\секции.png");

            string relPath3 = @"3\";
            string path3 = Path.Combine(absPath, relPath3);
            path3 = Path.GetFullPath(path3);

            string pathImg3 = Path.Combine(absPath, @"Resources\номер.png");

            string relPath4 = @"4\";
            string path4 = Path.Combine(absPath, relPath4);
            path4 = Path.GetFullPath(path4);

            string pathImg4 = Path.Combine(absPath, @"Resources\функция.png");

            string relPath5 = @"5\";
            string path5 = Path.Combine(absPath, relPath5);
            relPath5 = Path.GetFullPath(path5);

            string pathImg5 = Path.Combine(absPath, @"Resources\имя.png");

            var panel = application.CreateRibbonPanel(tabName, "Квартирография");

            var button_1 = new PushButtonData("Этаж", "Этаж",
                Path.Combine(path1, "ADSK_Floor.dll"),
                "ADSK_Floor.Main");

            var button_2 = new PushButtonData("Секция", "Секция",
                Path.Combine(path2, "ADSK_Section.dll"),
                "ADSK_Section.Main");

            var button_3 = new PushButtonData("Номер здания", "Номер здания",
                Path.Combine(path3, "ADSK_Number_building.dll"),
                "ADSK_Number_building.Main");

            var button_4 = new PushButtonData("Функция\nпомещения", "Функция\nпомещения",
                Path.Combine(path4, "ADSK_Room_Function.dll"),
                "ADSK_Room_Function.Main");

            var button_5 = new PushButtonData("Имя\nпомещения", "Имя\nпомещения",
                Path.Combine(path5, "PNR_Room_Name.dll"),
                "PNR_Room_Name.Main");

            Uri uriImage1 = new Uri(pathImg1, UriKind.Absolute);
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
            button_4.LargeImage = largeImage4;

            Uri uriImage5 = new Uri(pathImg5, UriKind.Absolute);
            BitmapImage largeImage5 = new BitmapImage(uriImage5);
            button_5.LargeImage = largeImage5;

            panel.AddItem(button_1);
            panel.AddItem(button_2);
            panel.AddItem(button_3);
            panel.AddItem(button_4);
            panel.AddItem(button_5);

            return Result.Succeeded;
        }
    }
}