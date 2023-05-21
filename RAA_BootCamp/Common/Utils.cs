#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using OfficeOpenXml;

using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace RAA_BootCamp.Common
{
    public static class Utils
    {
        // Declare a static variable to store the last selected directory
        private static string lastSelectedDirectory = "C:\\";
        /// <summary>
        /// Prompts the user to select an Excel file.
        /// </summary>
        /// <returns>The selected Excel file path, or null if no file was selected.</returns>
        public static string UM_GetExcelFilePath()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();  // Create an instance of OpenFileDialog to browse for files

            openFileDialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";  // Set the file filter to display only Excel files
            openFileDialog.InitialDirectory = lastSelectedDirectory;  // Set the initial directory to the last selected directory
            openFileDialog.Multiselect = false;  // Allow selecting only a single file
            openFileDialog.CheckFileExists = false;  // Disable checking if the selected file exists
            openFileDialog.CheckPathExists = false;  // Disable checking if the selected path exists
            openFileDialog.RestoreDirectory = true;  // Restore the directory to the previously selected one

            string excelFile = "";  // Variable to store the selected Excel file path

            if (openFileDialog.ShowDialog() == DialogResult.OK)  // Show the file dialog and check if the user clicked OK
            {
                excelFile = openFileDialog.FileName;  // Get the selected file path
                                                      // Update the last selected directory to the current selected directory
                lastSelectedDirectory = Path.GetDirectoryName(excelFile);
                return excelFile;  // Return the selected file path
            }

            return null;  // Return null if no file was selected or the dialog was canceled
        }

        //
        /// <summary>
        ///  Reads Excel file using Excel Interop
        /// How to use it:
        /// List<List<string>> interopXlsx = Utils.UM_ImportExcelFileUsingExcelInterop();
        /// </summary>
        /// <returns>List<List<string>></returns>
        public static List<List<string>> UM_ImportExcelFileUsingExcelInterop()
        {
            string filePath = Utils.UM_GetExcelFilePath();  // Prompt the user to select an Excel file using the custom method UM_GetExcelFilePath()

            if (filePath == null)  // Check if no file was selected
            {
                TaskDialog.Show("Error", "Please select an Excel file");  // Show an error message to the user
                return null;  // Return null to indicate an error occurred
            }

            // Open Excel file
            Excel.Application excel = new Excel.Application();  // Create an instance of the Excel application
            Excel.Workbook workbook = excel.Workbooks.Open(filePath);  // Open the Excel workbook at the specified file path
            Excel.Worksheet worksheet = workbook.Worksheets[1] as Excel.Worksheet;  // Get the first worksheet in the workbook
            Excel.Range range = (Excel.Range)worksheet.UsedRange;  // Get the used range of the worksheet

            int rows = range.Rows.Count;  // Get the total number of rows in the worksheet
            int columns = range.Columns.Count;  // Get the total number of columns in the worksheet

            // Read Excel data into a list
            List<List<string>> excelData = new List<List<string>>();  // Create a list to store the Excel data
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();  // Create a list to store data for each row
                for (int j = 1; j <= columns; j++)
                {
                    var cellContent = (worksheet.Cells[i, j] as Excel.Range)?.Value2;  // Get the value of the cell as an object
                    string cellcon = cellContent?.ToString() ?? string.Empty;  // Convert the cell value to string
                    rowData.Add(cellcon);  // Add the cell value to the row data list
                }
                excelData.Add(rowData);  // Add the row data to the Excel data list
            }
            workbook.Save();
            excel.Quit();
            return excelData;
        }

        /// <summary>
        ///  Reads Excel file using EPPlus
        /// How to use it:
        /// List<List<string>> epplusXlsx = Utils.UM_ImportExcelFileUsingEPPlus();
        /// </summary>
        /// <returns>List<List<string>></returns>
        public static List<List<string>> UM_ImportExcelFileUsingEPPlus(int sheetNumber)
        {
            string filePath = Utils.UM_GetExcelFilePath();  // Prompt the user to select an Excel file using the custom method UM_GetExcelFilePath()

            if (filePath == null)  // Check if no file was selected
            {
                TaskDialog.Show("Error", "Please select an Excel file");  // Show an error message to the user
                return null;  // Return null to indicate an error occurred
            }

            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // Set the license context for EPPlus to NonCommercial

            // Open Excel file using EPPlus library
            ExcelPackage excel = new ExcelPackage(filePath);  // Create an instance of ExcelPackage by providing the file path
            ExcelWorkbook workbook = excel.Workbook;  // Get the workbook from the Excel package
            ExcelWorksheet worksheet = workbook.Worksheets[0];  // Get the first worksheet (index 0) from the workbook


            int rows = worksheet.Dimension.Rows;  // Get the total number of rows in the worksheet
            int columns = worksheet.Dimension.Columns;  // Get the total number of columns in the worksheet

            // Read Excel data into a list
            List<List<string>> excelData = new List<List<string>>();  // Create a list to store the Excel data
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();  // Create a list to store data for each row
                for (int j = 1; j <= columns; j++)
                {
                    // Get the value of the cell at the current row and column
                    string cellContent = worksheet.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);  // Add the cell value to the row data list
                }
                excelData.Add(rowData);  // Add the row data to the Excel data list
            }

            //Save and close excel file
            excel.Save();
            excel.Dispose();
            return excelData;
        }


        public static List<List<string>> UM_ImportExcelFileUsingEPPlusBySheetName(string sheetName)
        {
            string filePath = Utils.UM_GetExcelFilePath();  // Prompt the user to select an Excel file using the custom method UM_GetExcelFilePath()

            if (filePath == null)  // Check if no file was selected
            {
                TaskDialog.Show("Error", "Please select an Excel file");  // Show an error message to the user
                return null;  // Return null to indicate an error occurred
            }

            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // Set the license context for EPPlus to NonCommercial

            // Open Excel file using EPPlus library
            ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));  // Create an instance of ExcelPackage by providing the file path
            ExcelWorkbook workbook = excel.Workbook;  // Get the workbook from the Excel package

            ExcelWorksheet worksheet = workbook.Worksheets[sheetName];  // Get the worksheet with the specified sheet name

            if (worksheet == null)
            {
                TaskDialog.Show("Error", $"Sheet '{sheetName}' not found in the Excel file");
                excel.Dispose();
                return null;
            }

            int rows = worksheet.Dimension.Rows;  // Get the total number of rows in the worksheet
            int columns = worksheet.Dimension.Columns;  // Get the total number of columns in the worksheet

            // Read Excel data into a list
            List<List<string>> excelData = new List<List<string>>();  // Create a list to store the Excel data
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();  // Create a list to store data for each row
                for (int j = 1; j <= columns; j++)
                {
                    // Get the value of the cell at the current row and column
                    string cellContent = worksheet.Cells[i, j].Value?.ToString() ?? string.Empty;
                    rowData.Add(cellContent);  // Add the cell value to the row data list
                }
                excelData.Add(rowData);  // Add the row data to the Excel data list
            }

            // Save and close the Excel file
            excel.Dispose();

            return excelData;
        }



        public static Dictionary<string, List<List<string>>> UM_ImportExcelFileUsingEPPlus_GetAllSheetsAndDataAsDictionary()
        {
            string filePath = Utils.UM_GetExcelFilePath();  // Prompt the user to select an Excel file using the custom method UM_GetExcelFilePath()

            if (filePath == null)  // Check if no file was selected
            {
                TaskDialog.Show("Error", "Please select an Excel file");  // Show an error message to the user
                return null;  // Return null to indicate an error occurred
            }

            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // Set the license context for EPPlus to NonCommercial

            // Open Excel file using EPPlus library
            ExcelPackage excel = new ExcelPackage(new FileInfo(filePath));  // Create an instance of ExcelPackage by providing the file path
            ExcelWorkbook workbook = excel.Workbook;  // Get the workbook from the Excel package

            Dictionary<string, List<List<string>>> excelData = new Dictionary<string, List<List<string>>>();

            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                int rows = worksheet.Dimension.Rows;  // Get the total number of rows in the worksheet
                int columns = worksheet.Dimension.Columns;  // Get the total number of columns in the worksheet

                // Read Excel data into a list
                List<List<string>> sheetData = new List<List<string>>();  // Create a list to store the Excel data
                for (int i = 1; i <= rows; i++)
                {
                    List<string> rowData = new List<string>();  // Create a list to store data for each row
                    for (int j = 1; j <= columns; j++)
                    {
                        // Get the value of the cell at the current row and column
                        string cellContent = worksheet.Cells[i, j].Value?.ToString() ?? string.Empty;
                        rowData.Add(cellContent);  // Add the cell value to the row data list
                    }
                    sheetData.Add(rowData);  // Add the row data to the Excel data list
                }

                excelData.Add(worksheet.Name, sheetData);  // Add sheet name and its data to the dictionary
            }

            // Save and close the Excel file
            excel.Dispose();

            return excelData;
        }

        public static List<List<string>> UM_GetsheetDataFromImportedExcelBySheetName_1(Dictionary<string, List<List<string>>> _excelSheets, string excelSheetName)
        {
            List<List<string>> sheetData;

            if (_excelSheets.TryGetValue(excelSheetName, out sheetData))
            {
                // You could use sheetData here
                // It contains the data corresponding to the specified sheet name
                // For example:
                foreach (List<string> row in sheetData)
                {
                    foreach (string cellValue in row)
                    {
                        Debug.Print($"{cellValue}");
                    }
                }
            }
            else
            {
                // The specified sheet name was not found in the dictionary
            }

            return sheetData;
        }

        public static List<List<string>> UM_GetsheetDataFromImportedExcelBySheetName(Dictionary<string, List<List<string>>> _excelSheets, string excelSheetName)
        {
            List<List<string>> sheetData = null;
            List<List<string>> allRowsReturned = new List<List<string>>(); // Initialize the list to store all rows

            if (_excelSheets.TryGetValue(excelSheetName, out sheetData))
            {
                // The dictionary contains the data for the specified sheet name

                if (sheetData.Count > 0)
                {
                    // Iterate through the sheet data, starting from index 1 to skip the first row
                    for (int i = 1; i < sheetData.Count; i++)
                    {
                        // Get the current row
                        List<string> row = sheetData[i];
                        allRowsReturned.Add(row); // Add the row to the list of all rows

                        // Iterate through the cells in the row
                        foreach (string cellValue in row)
                        {
                            // Print the cell value
                            Debug.Print($"{cellValue}");
                        }
                    }
                }
            }
            else
            {
                // The specified sheet name was not found in the dictionary
                return null;
            }

            // Return the list of all rows
            return allRowsReturned;
        }


        //internal static List<FurnitureSet> MU_GetListOfFurnitureSets(List<List<string>> _furnitureSetsFromExcel)
        //{
        //    List<FurnitureSet> furniturSets = new List<FurnitureSet>();
        //    foreach (var curfurnitureSet in _furnitureSetsFromExcel)
        //    {
        //        FurnitureSet curFurSet = new FurnitureSet();
        //        curFurSet.Furniture_Set = curfurnitureSet[0];
        //        curFurSet.RoomType = curfurnitureSet[1];
        //        curFurSet.IncludedFurniture = curfurnitureSet[2];
        //        furniturSets.Add(curFurSet);
        //    }
        //    return furniturSets;  
        //}
        internal static List<FurnitureSet> MU_GetListOfFurnitureSets(List<List<string>> _furnitureSetsFromExcel)
        {
            List<FurnitureSet> furniturSets = new List<FurnitureSet>();
            foreach (var curfurnitureSet in _furnitureSetsFromExcel)
            {
                FurnitureSet curFurSet = new FurnitureSet(curfurnitureSet[0].Trim(), curfurnitureSet[1].Trim(), curfurnitureSet[2].Trim());

                furniturSets.Add(curFurSet);
            }
            return furniturSets;
        }

        internal static List<FurnitureData> MU_GetListOfFurnitureTypes(Document doc, List<List<string>> _furnitureSetsFromExcel)
        {
            List<FurnitureData> furnitureTypes = new List<FurnitureData>();
            foreach (var curfurnituretype in _furnitureSetsFromExcel)
            {
                FurnitureData curFurType = new FurnitureData(doc, curfurnituretype[0].Trim(), curfurnituretype[1].Trim(), curfurnituretype[2].Trim());
                furnitureTypes.Add(curFurType);
            }
            return furnitureTypes;
        }

        internal static List<SpatialElement> MU_GetAllRooms(Document doc)
        {
            FilteredElementCollector roomsCollector = new FilteredElementCollector(doc)
                                                        .OfCategory(BuiltInCategory.OST_Rooms)  // Filter by rooms category
                                                        .WhereElementIsNotElementType();  // Exclude room element types

            List<SpatialElement> allRooms = roomsCollector.Cast<SpatialElement>().ToList();
            return allRooms;
        }

        internal static string GetParamValue_Method1(Element _curElem, string paramName)
        {
            // Iterate through the parameters of the element
            foreach (Parameter curParam in _curElem.Parameters)
            {
                // Check if the parameter name matches the desired paramName
                if (curParam.Definition.Name == paramName)
                {
                    return curParam.AsString();  // Return the parameter value as a string
                }
            }

            return null;  // Return null if the parameter is not found
        }

        internal static string GetParamValue_Method2(Element _curElem, string paramName)
        {
            // Convert the ParameterSet to a list
            List<Parameter> parameters = _curElem.Parameters.Cast<Parameter>().ToList();

            // Find the first parameter that matches the given paramName
            Parameter matchingParam = parameters.FirstOrDefault(curParam => curParam.Definition.Name == paramName);

            // This could return a null value if no matchingParameter found
            return matchingParam?.AsString();
        }

        internal static FurnitureData GetFamilyInfo(string curFurn, List<FurnitureData> furnitureTypes)
        {
            foreach (FurnitureData furniture in furnitureTypes)
            {
                if (furniture.FurnitureName == curFurn)
                {
                    return furniture;
                }
            }
            return null;
        }

        internal static void SetParamValueAsInt(Element curElem, string paramName, int paramValue)
        {
            // Iterate through the parameters of the element
            foreach (Parameter curParam in curElem.Parameters)
            {
                // Check if the parameter name matches the desired paramName
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);  // Set the parameter value to the specified integer
                }
            }
        }


        public static void InsertFurnitureInRooms_Method2(List<SpatialElement> roomList, List<FurnitureSet> furnitureSetsList, List<FurnitureData> furnitureDataList, Document doc, ref int counter)
        {
            foreach (SpatialElement room in roomList)
            {
                LocationPoint roomPoint = room.Location as LocationPoint;
                XYZ insertionPoint = roomPoint?.Point; // Get the insertion point of the room

                string curRoomFurnSet = Utils.GetParamValue_Method2(room, "Furniture Set"); // Get the value of the "Furniture Set" parameter for the current room
                var matchedFurnitureSets = furnitureSetsList.Where(f => f.Furniture_Set == curRoomFurnSet); // Filter the furniture sets that match the current room's furniture set

                foreach (FurnitureSet furnSet in matchedFurnitureSets)
                {
                    var matchingFurnitureData = furnSet.IncludedFurniture
                        .Join(furnitureDataList, curFurn => curFurn, furnitureData => furnitureData.FurnitureName, (curFurn, furnitureData) => furnitureData) // Match the included furniture in the furniture set with the furniture data list based on their names
                        .ToList();

                    foreach (FurnitureData furnData in matchingFurnitureData)
                    {
                        furnData.familySymbol.Activate(); // Activate the family symbol
                        FamilyInstance newFamilyInstance = doc.Create.NewFamilyInstance(insertionPoint, furnData.familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural); // Create a new family instance in the room using the family symbol
                        counter++; // Increment the counter for the inserted families
                    }

                    Utils.SetParamValueAsInt(room, "Furniture Count", furnSet.FurnitureCount()); // Set the "Furniture Count" parameter value for the current room
                }
            }
        }

    }

    public class FurnitureSet
    {
        public string Furniture_Set { get; set; }
        public string RoomType { get; set; }
        public List<string> IncludedFurniture { get; private set; }

        public FurnitureSet(string furniture_Set, string roomType, string includedFurniture)
        {
            Furniture_Set = furniture_Set;
            RoomType = roomType;
            IncludedFurniture = GetFunitureSetFromString(includedFurniture);
        }
        private List<string> GetFunitureSetFromString(string furnitureList)
        {
            List<string> returnList = furnitureList.Split(',')  // Split the input string by comma
                                                  .Select(s => s.Trim())  // Trim each string before adding it to the list
                                                  .ToList();  // Convert the IEnumerable<string> to a List<string>
            return returnList;  // Return the resulting list
        }

        public int FurnitureCount()
        {
            return IncludedFurniture.Count;
        }
    }

    public class FurnitureData
    {
        public string FurnitureName { get; set; }
        public string RevitFamilyName { get; set; }
        public string RevitFamilyType { get; set; }
        public FamilySymbol familySymbol { get; set; }
        public Document doc { get; set; }
        public FurnitureData(Document _doc, string furnitureName, string revitFamilyName, string revitFamilyType)
        {
            FurnitureName = furnitureName;
            RevitFamilyName = revitFamilyName;
            RevitFamilyType = revitFamilyType;
            doc = _doc;
            familySymbol = GetFamilySymbol_Method2();
        }

        private FamilySymbol GetFamilySymbol_Method1()
        {
            FilteredElementCollector familyCollector = new FilteredElementCollector(doc);  // Create a FilteredElementCollector for all elements
            familyCollector.OfClass(typeof(Family));  // Filter the collector to include only Family elements

            foreach (Family curFam in familyCollector)  // Iterate through each Family element
            {
                if (curFam.Name == RevitFamilyName)  // Check if the Family name matches the specified RevitFamilyName
                {
                    ISet<ElementId> famSymbolList = curFam.GetFamilySymbolIds();  // Get the set of FamilySymbol element IDs for the current Family

                    foreach (ElementId curId in famSymbolList)  // Iterate through each FamilySymbol element ID in the set
                    {
                        FamilySymbol curFamSymbol = doc.GetElement(curId) as FamilySymbol;  // Retrieve the FamilySymbol element from the document

                        if (curFamSymbol.Name == RevitFamilyType)  // Check if the FamilySymbol name matches the specified RevitFamilyType
                        {
                            return curFamSymbol;  // Return the matching FamilySymbol
                        }
                    }
                }
            }

            return null;  // Return null if no matching FamilySymbol is found
        }


        private FamilySymbol GetFamilySymbol_Method2()
        {
            var familyCollector = new FilteredElementCollector(doc)  // Create a FilteredElementCollector for families
                .OfClass(typeof(Family))  // Filter by elements of type Family
                .Cast<Family>()  // Cast the elements to Family type
                .FirstOrDefault(family => family.Name == RevitFamilyName);  // Find the first family with matching name

            if (familyCollector != null)
            {
                var familySymbol = familyCollector.GetFamilySymbolIds()  // Get the set of family symbol IDs
                    .Select(id => doc.GetElement(id) as FamilySymbol)  // Convert element IDs to FamilySymbol objects
                    .FirstOrDefault(symbol => symbol != null && symbol.Name == RevitFamilyType);  // Find the first symbol with matching name

                return familySymbol;  // Return the found family symbol
            }

            return null;  // Return null if no matching family is found
        }

    }
}
