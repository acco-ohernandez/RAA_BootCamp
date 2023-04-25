#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RAA_BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class Module01Challenge : IExternalCommand
    {
        // Global Variables
        int levelsCount;
        int sheetCount;
        string globalFizzBuzzName;
        int floorPlanCount;
        int ceilingPlanCount;
        int FizzBuzzViewPortCounter;

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
         
            //Variable Declarations
            int _number = 250;
            double _elevation = 0;
            double _floorHeight = 15;

            //Reset global variables
            ResetGlobalVariables();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("FizzBuzz creation");

                for (int i = 0; i < _number; i++)
                {
                    string _FizzBuzzName = GetTheFizzBuzzName(i);                  // Get the FizzBuzzName
                    var _newLevel = CreateLevel(doc, _elevation, _FizzBuzzName);   // Call Create Level method and name it
                    _elevation = _elevation + _floorHeight;                        // Increment the elevation

                    if (globalFizzBuzzName == "FIZZBUZZ")
                    { 
                        var _newViewSheet = CreateViewSheet(doc, _FizzBuzzName);    // Create FIZZBUZZ_# Sheet Sheet

                        //Bonus Part 1
                        //creating a sheet, create a floor plan for each FIZZBUZZ.
                        var _newFizzBuzzFloorViewPlan = CreateFizzBuzzFloorViewPlan(doc, _FizzBuzzName, _newLevel.Id);
                        //Bonus Part 2
                        //add the floor plan to the sheet by creating a Viewport element
                        Viewport _newViewPort = addFloorPlanViewPortToSheet(doc, _newViewSheet, _newFizzBuzzFloorViewPlan);
                    }
                    else if (globalFizzBuzzName == "FIZZ")
                    {
                        var _newFloorPlan = CreateFloorPlan(doc, _FizzBuzzName, _newLevel.Id);    // Call method to create new floorPlan
                    }
                    else if (globalFizzBuzzName == "BUZZ")
                    {
                        var _newCeilingPlan = CreateCeilingPlan(doc, _FizzBuzzName, _newLevel.Id);    // Call method to create new floorPlan
                    }
                }
                tx.Commit();
            }

            TaskDialog.Show("INFO",$"{levelsCount} Levels created\n" +
                                   $"{sheetCount} Sheets Created\n" +
                                   $"{floorPlanCount} Floor Plan(s) Created\n" +
                                   $"{ceilingPlanCount} Ceiling Plan(s) Created\n" +
                                   $"{FizzBuzzViewPortCounter} ViewPort(s) Created");

            return Result.Succeeded;
        }

        private Viewport addFloorPlanViewPortToSheet(Autodesk.Revit.DB.Document doc, Element _viewSheetElem, Element _floorPlanElem)
        {
            XYZ insertPoint = new XYZ(2, 1, 0);
            Viewport _newViewPort = Viewport.Create(doc, _viewSheetElem.Id, _floorPlanElem.Id, insertPoint);
            FizzBuzzViewPortCounter++;
            return _newViewPort;
        }
        private ViewPlan CreateFizzBuzzFloorViewPlan(Autodesk.Revit.DB.Document doc, string fizzBuzzName, ElementId levelId)
        {
            // Get view family types
            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType floorPlanVFT = null;
            foreach (ViewFamilyType curViewFamType in vftCollector)
            {
                if (curViewFamType.ViewFamily == ViewFamily.FloorPlan)
                {
                    floorPlanVFT = curViewFamType;
                }
            }
            var newFloorPlan = ViewPlan.Create(doc, floorPlanVFT.Id, levelId);          // Create Floor Plan
            newFloorPlan.Name = fizzBuzzName;
            //floorPlanCount++;
            return newFloorPlan;
        }
        private void ResetGlobalVariables()
        {
            levelsCount = 0;
            sheetCount = 0;
            globalFizzBuzzName = "";
            floorPlanCount = 0;
            ceilingPlanCount = 0;
            FizzBuzzViewPortCounter = 0;
        }
        private string GetTheFizzBuzzName(int num)
        {
            string returnFizzBuzzName;
            if (num % 3 == 0 && num % 5 == 0)
            {
                returnFizzBuzzName = $"FIZZBUZZ_{num}";
                globalFizzBuzzName = "FIZZBUZZ";
                //CreateViewSheet(doc, FizzBuzzName);
            }
            else if (num % 3 == 0)
            {
                returnFizzBuzzName = $"FIZZ_{num}";
                globalFizzBuzzName = "FIZZ";
                //CreateCeilingPlan(doc, FizzBuzzName);                    // NOT DONE, NEED TO IMPLEMENT
            }
            else if (num % 5 == 0)
            {
                returnFizzBuzzName = $"BUZZ_{num}";
                globalFizzBuzzName = "BUZZ";
            }
            else
            {
                returnFizzBuzzName = $"{num} % (3 or 5) != 0";
                globalFizzBuzzName = "";
            }

            return returnFizzBuzzName;
        }
        private ViewSheet CreateViewSheet(Autodesk.Revit.DB.Document doc, string fizzBuzzName)
        {
            // Get an available title block from document
            FilteredElementCollector titleBlockCatCollector = new FilteredElementCollector(doc);
            titleBlockCatCollector.OfClass(typeof(FamilySymbol));
            titleBlockCatCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);

            ViewSheet _newViewSheet = ViewSheet.Create(doc, titleBlockCatCollector.FirstElementId());
            _newViewSheet.Name = fizzBuzzName;
            sheetCount++;
            return _newViewSheet;
        }
        private Level CreateLevel(Autodesk.Revit.DB.Document doc, double _elevation, string _levelName)
        {
            Level _newLevel = Level.Create(doc, _elevation);             // create a levels
            _newLevel.Name = _levelName;                                // Rename the level
            Debug.Print($"Level Created. ID:{_newLevel.Id}  Name:{_newLevel.Name}");
            levelsCount++;
            return _newLevel;
        }
        private object CreateFloorPlan(Autodesk.Revit.DB.Document doc, string fizzBuzzName, ElementId levelId)
        {
            // Get view family types
            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType floorPlanVFT = null;
            foreach (ViewFamilyType curViewFamType in vftCollector)
            {
                if (curViewFamType.ViewFamily == ViewFamily.FloorPlan)
                {
                    floorPlanVFT = curViewFamType;
                }
            }
            var newFloorPlan = ViewPlan.Create(doc, floorPlanVFT.Id, levelId);          // Create Floor Plan
            newFloorPlan.Name = fizzBuzzName;
            floorPlanCount++;
            return floorPlanVFT;

        }
        private object CreateCeilingPlan(Autodesk.Revit.DB.Document doc, string fizzBuzzName, ElementId levelId)
        {
            // Get view family types
            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType ceilingPlanVFT = null;
            foreach (ViewFamilyType curViewFamType in vftCollector)
            {
                if (curViewFamType.ViewFamily == ViewFamily.CeilingPlan)
                {
                    ceilingPlanVFT = curViewFamType;
                }
            }
            var newFloorPlan = ViewPlan.Create(doc, ceilingPlanVFT.Id, levelId);          // Create Ceiling Plan
            newFloorPlan.Name = fizzBuzzName;
            ceilingPlanCount++;
            return ceilingPlanVFT;
        }
    }
}
