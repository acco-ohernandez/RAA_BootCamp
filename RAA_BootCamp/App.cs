#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using RAA_BootCamp;
using RAA_BootCamp.Common;

#endregion

namespace RAA_BootCamp
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            string myTabName = "Revit Add -in Bootcamp";
            // 1. Create ribbon tab
            try
            {
                app.CreateRibbonTab(myTabName);
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists.");
            }

            // 2. Create ribbon panel 
            RibbonPanel panel = Utils.CreateRibbonPanel(app, myTabName, "Revit Tools");

            // 3. Create button data instances 
            ButtonDataClass myBtn1Data = new ButtonDataClass("btn_Tool1", "Btn 1", Module04Challenge.GetMethod(), Properties.Resources.Green_32, Properties.Resources.Green_16, "This is a tooltip");
            ButtonDataClass myBtn2Data = new ButtonDataClass("btn_Tool2", "Btn 2", Module04Challenge.GetMethod(), Properties.Resources.Aqua_32, Properties.Resources.Aqua_16, "This is a tooltip");
            ButtonDataClass myBtn3Data = new ButtonDataClass("btn_Tool3", "Btn 3", Module04Challenge.GetMethod(), Properties.Resources.GBlue_32, Properties.Resources.GBlue_16, "This is a tooltip");
            ButtonDataClass myBtn4Data = new ButtonDataClass("btn_Tool4", "Btn 4", Module04Challenge.GetMethod(), Properties.Resources.LBlue_32, Properties.Resources.LBlue_16, "This is a tooltip");
            ButtonDataClass myBtn5Data = new ButtonDataClass("btn_Tool5", "Btn 5", Module04Challenge.GetMethod(), Properties.Resources.Yellow_32, Properties.Resources.Yellow_16, "This is a tooltip");
            ButtonDataClass myBtn6Data = new ButtonDataClass("btn_Tool5", "Btn 6\rSplit", Module04Challenge.GetMethod(), Properties.Resources.Brown_32, Properties.Resources.Brown_16, "This is a tooltip");
            ButtonDataClass myBtn7Data = new ButtonDataClass("btn_Tool7", "Btn 7\rSplit", Module04Challenge.GetMethod(), Properties.Resources.Pink_32, Properties.Resources.Pink_16, "This is a tooltip");
            ButtonDataClass myBtn8Data = new ButtonDataClass("btn_Tool8", "Btn 8\rPB", Module04Challenge.GetMethod(), Properties.Resources.Orange_32, Properties.Resources.Orange_16, "This is a tooltip");
            ButtonDataClass myBtn9Data = new ButtonDataClass("btn_Tool9", "Btn 9\rPB", Module04Challenge.GetMethod(), Properties.Resources.Gray_32, Properties.Resources.Gray_16, "This is a tooltip");
            ButtonDataClass myBtn10Data = new ButtonDataClass("btn_Tool10", "Btn 10\rPB", Module04Challenge.GetMethod(), Properties.Resources.Black_32, Properties.Resources.Black_16, "This is a tooltip");

            // 4. Create buttons
            // Push Buttons
            PushButton myBtn1 = panel.AddItem(myBtn1Data.Data) as PushButton;
            PushButton myBtn2 = panel.AddItem(myBtn2Data.Data) as PushButton;

            // Stacked Buttons
            panel.AddStackedItems(myBtn3Data.Data, myBtn4Data.Data, myBtn5Data.Data);

            // Split Buttons
            SplitButtonData splitButtonData = new SplitButtonData("split1", "Split\rButton");
            SplitButton splitButton = panel.AddItem(splitButtonData) as SplitButton;
            splitButton.AddPushButton(myBtn6Data.Data);
            splitButton.AddPushButton(myBtn7Data.Data);

            //Pulldown Button
            PulldownButtonData pullDownData = new PulldownButtonData("pulldown1", "More Tools\rButton");
            pullDownData.LargeImage = ButtonDataClass.BitmapToImageSource(Properties.Resources.CoolFace_32x32);
            pullDownData.Image = ButtonDataClass.BitmapToImageSource(Properties.Resources.CoolFace_16x16);
            PulldownButton pullDownButton = panel.AddItem(pullDownData) as PulldownButton;
            pullDownButton.AddPushButton(myBtn8Data.Data);
            pullDownButton.AddPushButton(myBtn9Data.Data);
            pullDownButton.AddPushButton(myBtn10Data.Data);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
