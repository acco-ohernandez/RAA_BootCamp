#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

//using Microsoft.Office.Interop.Excel;

using RAA_BootCamp.Common;

#endregion

namespace RAA_BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class Module03Challenge : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;


            TaskDialog.Show("Info", "\"Please importe the:\n RAB_Module 03_Furniture.xlsx\""); // Tell the user what to do next.

            var _excelSheets = MyUtils.UM_ImportExcelFileUsingEPPlus_GetAllSheetsAndDataAsDictionary(); // Import the Excel file
            var _furnitureSetsFromExcel = MyUtils.UM_GetsheetDataFromImportedExcelBySheetName(_excelSheets, "Furniture sets");  // Get "Furniture sets" sheet
            var _furnitureTypesFromExcel = MyUtils.UM_GetsheetDataFromImportedExcelBySheetName(_excelSheets, "Furniture types"); // Get "Furniture types" sheet

            List<FurnitureSet> furnitureSetsList = MyUtils.MU_GetListOfFurnitureSets(_furnitureSetsFromExcel);    // List of FurnitureSet Types
            List<FurnitureData> furnitureDataList = MyUtils.MU_GetListOfFurnitureTypes(doc, _furnitureTypesFromExcel);// List of FurnitureData Types

            int counter = 0;
            List<SpatialElement> roomList = MyUtils.MU_GetAllRooms(doc); // Get a list of all the rooms

            using (Transaction t = new Transaction(doc, "Inserted Furniture"))
            {
                t.Start();

                // Method 1
                //Utils.InsertFurnitureInRooms_Method1(roomList, furnitureSetsList, furnitureDataList, doc, ref counter);

                // Method 2
                //Utils.InsertFurnitureInRooms_Method2(roomList, furnitureSetsList, furnitureDataList, doc, ref counter);

                // Method 3
                MyUtils.InsertFurnitureInRooms_Method3(roomList, furnitureSetsList, furnitureDataList, doc, ref counter);

                t.Commit();
            }

            TaskDialog.Show("Done", $"Inserted {counter} families.");
            return Result.Succeeded;
        }

    }

}
