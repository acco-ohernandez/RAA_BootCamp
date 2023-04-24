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

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("FizzBuzz Levels");

                //for (int i = 1; i < _number; i++)
                for (int i = 0; i < _number; i++)
                    {
                    Level _newLevel = Level.Create(doc, _elevation); // create a levels
                    string _levelName = GetTheFizzBuzzName(i);       // Get the FizzBuzzName
                    _newLevel.Name = _levelName;                     // Rename the level
                    _elevation = _elevation + _floorHeight;          // Increment the elevation
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }

        private string GetTheFizzBuzzName(double _floorHeight)
        {
            string FizzBuzzName;
            if (_floorHeight % 3 == 0 && _floorHeight % 5 == 0)
                FizzBuzzName = $"BUZZBUZZ_{_floorHeight}";
            else if (_floorHeight % 3 == 0)
                FizzBuzzName = $"FIZZ_{_floorHeight}";
            else if (_floorHeight % 5 == 0)
                FizzBuzzName = $"BUZZ{_floorHeight}";
            else
                FizzBuzzName = $"{_floorHeight} % (3 or 5) != 0";
            return FizzBuzzName;
        }
    }
}
