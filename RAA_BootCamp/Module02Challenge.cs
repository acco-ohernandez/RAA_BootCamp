#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

#endregion

#region classStart
namespace RAA_BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class ComModule02Challengemand : IExternalCommand
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
            #endregion

            // Get Level One which will be used to Create all the elements on it.
            Level _levelOne = _GetLevelByName(doc, "Level 1");

            // Get Supply Air Duct System type
            MEPSystemType _supplyAirMEPSystemType = _GetMEPSystemType(doc, "Supply Air");

            // Get The default duct type
            DuctType _defaultDuctType = _GetDefaultDuctType(doc, "Default");

            // Get the Pipe System Type
            MEPSystemType _domesticHotWaterPipeMEPSystemType = _GetPipeSystemType(doc, "Domestic Hot Water");

            // Get the Pipe Type
            PipeType _pipeType = _GetPipeSystemType(doc);


            // Tell the user what to do.
            TaskDialog.Show("INFO", "Select elements by drawing a rectangle over them.");
            // Pick Elements and filter them into list
            IList<Element> _selectedElements = uidoc.Selection.PickElementsByRectangle("Select elements by drawing a rectangle over them.");

            // Get the list of model curves
            List<CurveElement> _modelCurvesList = _GetModelCurvesList(_selectedElements);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Delete lines");
                // Loop through all the Model Curve Elements
                foreach (CurveElement _curCurveModelElem in _modelCurvesList)
                {
                    Curve _curve = _curCurveModelElem.GeometryCurve; // Retrieve the geometry curve of a model element

                    if (_curve != null && _curve.IsBound) // Check if the curve is bound and not null
                    {
                        XYZ _startPoint = _curve.GetEndPoint(0); // Get the start point of the curve
                        XYZ _endPoint = _curve.GetEndPoint(1); // Get the end point of the curve

                        GraphicsStyle _curStyle = _curCurveModelElem.LineStyle as GraphicsStyle; // Retrieve the graphics style of the model element

                        switch (_curStyle.Name) // Check the name of the graphics style
                        {
                            case "A-WALL": // Create a wall if the graphics style is "A-WALL"
                                Debug.Print($"Switch 1: {_curStyle.Name}\n Creating Wall");
                                var _wallTypeName = "Generic - 8\"";
                                Wall _wall = _CreateAWall(doc, _curve, _levelOne, _wallTypeName); // Call the method to create a wall
                                break;

                            case "M-DUCT": // Create a duct if the graphics style is "M-DUCT"
                                Debug.Print($"Switch 2: {_curStyle.Name}\n Creating Duct");
                                Duct _ductCreated = _CreateADuct(doc, _supplyAirMEPSystemType, _defaultDuctType, _levelOne, _startPoint, _endPoint); // Call the method to create a duct
                                break;

                            case "P-PIPE": // Create a pipe if the graphics style is "P-PIPE"
                                Debug.Print($"Switch 3: {_curStyle.Name}\n Creating Pipe");
                                var _pipeCreated = _CreateAPipe(doc, _domesticHotWaterPipeMEPSystemType, _pipeType, _levelOne, _startPoint, _endPoint); // Call the method to create a pipe
                                break;

                            case "A-GLAZ": // Create a glazed wall if the graphics style is "A-GLAZ"
                                Debug.Print($"Switch 4: {_curStyle.Name}\n Creating Glaz wall");
                                var _wallTypeName2 = "Storefront";
                                Wall _wall2 = _CreateAWall(doc, _curve, _levelOne, _wallTypeName2); // Call the method to create a glazed wall
                                break;

                            default: // Delete the model element if none of the above cases match
                                var e = _curCurveModelElem as Element;
                                doc.Delete(e.Id);
                                break;
                        }
                    }
                    else
                    {
                        // Detele the other elements not used
                        Debug.Print("Deleting Invalid or null curve.");

                        var e = _curCurveModelElem as Element;
                        doc.Delete(e.Id);

                    }
                }
                tx.Commit();

            }

            return Result.Succeeded;
        }

        private PipeType _GetPipeSystemType(Document doc)
        {
            FilteredElementCollector _pipeTypeCollector = new FilteredElementCollector(doc);
            _pipeTypeCollector.OfClass(typeof(PipeType));
            return _pipeTypeCollector.FirstElement() as PipeType;
        }

        private List<CurveElement> _GetModelCurvesList(IList<Element> _selectedElements)
        {
            var _modelCurvesList = new List<CurveElement>();
            // Filter the elements for model curves
            foreach (var _elem in _selectedElements)
            {
                if (_elem is CurveElement)
                {
                    var _curveElement = _elem as CurveElement;
                    if (_curveElement.CurveElementType == CurveElementType.ModelCurve)
                    {
                        _modelCurvesList.Add(_curveElement);
                        Debug.Print(_curveElement.Name);
                    }
                }

            }

            return _modelCurvesList;
        }

        private Level _GetLevelByName(Document doc, string _levelName)
        {
            // Search for Level element by name
            FilteredElementCollector levelsCollector = new FilteredElementCollector(doc).OfClass(typeof(Level));
            Level level = levelsCollector.FirstOrDefault(l => l.Name == _levelName) as Level;
            // Check if Level element exists and is not null
            if (level != null)
                return level;
            else
                return null;
        }

        private Wall _CreateAWall(Document doc, Curve curve, Level _level, string _wallTypeName)
        {
            FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
            wallTypes.OfClass(typeof(WallType));

            WallType myWallType = _GetWallTypeByName(doc, "Generic - 8\"");

            Wall _createdWall = Wall.Create(doc, curve, _level.Id, false);
            return _createdWall;
        }

        private WallType _GetWallTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (WallType curType in collector)
            {
                //if ((curType.Name).Contains(typeName))
                //{
                //    Debug.Print(curType.Name);

                //    return curType;
                //}
                if (curType.Name == typeName)
                {
                    Debug.Print($"_GetWallTypeByName(): {curType.Name}");
                    return curType;
                }
            }

            return null;
        }
        private DuctType _GetDefaultDuctType(Document doc, string _ductTypeName)
        {
            // Get the default duct type
            FilteredElementCollector ductTypesCollector = new FilteredElementCollector(doc).OfClass(typeof(DuctType));
            DuctType _defaultDuctType = ductTypesCollector.Cast<DuctType>().FirstOrDefault(dt => dt.Name == _ductTypeName);
            return _defaultDuctType;
        }

        private Duct _CreateADuct(Document doc, MEPSystemType supplyAirMEPSystemType, DuctType defaultDuctType, Level levelOne, XYZ startPoint, XYZ endPoint)
        {
            Duct _ductCreated = Duct.Create(doc, supplyAirMEPSystemType.Id, defaultDuctType.Id, levelOne.Id, startPoint, endPoint);
            return _ductCreated;
        }

        private Pipe _CreateAPipe(Document doc, Element systemType, Element pipeType, Element level, XYZ startPoint, XYZ endPoint)
        {
            Pipe _pipeCreated = Pipe.Create(doc, systemType.Id, pipeType.Id, level.Id, startPoint, endPoint);
            if (_pipeCreated != null)
                return _pipeCreated;

            return null;
        }

        // Retrieves the MEPSystemType object that matches the specified system type name.
        private MEPSystemType _GetMEPSystemType(Document doc, string _systemTypeName)
        {
            // Create a filtered element collector for all MEP system types in the document
            FilteredElementCollector _sytemTypeCollector = new FilteredElementCollector(doc);
            _sytemTypeCollector.OfClass(typeof(MEPSystemType));

            // Iterate through the MEP system types to find the one with the specified name
            foreach (MEPSystemType _curSysType in _sytemTypeCollector)
            {
                if (_curSysType.Name == _systemTypeName)
                {
                    // If MEP system type with the specified name is found, return it.
                    return _curSysType;
                }
            }
            // If no MEP system type with the specified name is found, return null
            return null;
        }

        // Retrieves the MEPSystemType for the given system type name.
        private MEPSystemType _GetPipeSystemType(Document doc, string _systemTypeName)
        {
            // Filter the element collector for all DuctTypes
            FilteredElementCollector ductTypesCollector = new FilteredElementCollector(doc).OfClass(typeof(MEPSystemType));

            // Iterate through each DuctType
            foreach (MEPSystemType _curSysType in ductTypesCollector)
            {
                // Check if the current DuctType has the same name as the system type we're searching for
                if (_curSysType.Name == _systemTypeName)
                {
                    // If the names match, return the current DuctType as the MEPSystemType
                    return _curSysType;
                }
            }

            // If no matching MEPSystemType is found, return null
            return null;
        }

    }
}