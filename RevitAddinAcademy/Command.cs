#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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


            // Use the @ symbol tell VS to ignore the \ symbol (used as a break symbol in VS code)
            string excelFile = @"C:\Users\cflessner\Desktop\Revit Addin Class\Session02_CombinationSheetList-220706-171323.xlsx";

            // Open Excel file; be sure to CLOSE excel AND QUIT at the end of the code
            Excel.Application excelApp = new Excel.Application();

            // Open particular Excel workbook from file in previous step
            Excel.Workbook excelWB = excelApp.Workbooks.Open(excelFile);

            // Selects the FIRST worksheet in the workbook (starts at 1; not 0); did not reference worksheet by name
            Excel.Worksheet excelWS = excelWB.Worksheets.Item[1];

            // Selects the used cells in a range in Excel
            Excel.Range excelRng = excelWS.UsedRange;

            // Counts the number of rows of date in the Excel worksheet
            int rowCount = excelRng.Rows.Count;


            // do some stuff in Excel

            // create a list of string arrays
            List<string[]> dataList = new List<string[]>();

            for(int i = 1; i <= rowCount; i++)
            {
                // Get data from the first cell of the (i'th)first row
                Excel.Range cell1 = excelWS.Cells[i, 1];

                // Get data from the second cell of the (i'th)first row
                Excel.Range cell2 = excelWS.Cells[i, 2];

                // list the value of the data in the cell as a string value
                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                // Need to create and use an array when working with Excel
                // When we create the array, we must define how many spaces are inside the array; in this case, 2 spaces
                string[] dataArray = new string[2];

                // first array value from Excel file
                dataArray[0] = data1;

                // second array value from Excel file
                dataArray[1] = data2;

                // Now that we've created our arrays, we can add them to our array list(dataList)
                dataList.Add(dataArray);
            }

            // Need to create a transaction
            using (Transaction t = new Transaction(doc))
            {
                // start transaction
                t.Start("Create some Revit Stuff");

                // Create levels; Level method requires 2 arguments, current document(doc) and elevation of level
                Level curLevel = Level.Create(doc, 100);

                // Need to filter out the titleblock family we will need for the sheet creation
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                // using an OfCatgory collectory collector will give you ALL INSTANCES AS WELL AS TYPES
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                // Using the WhereElementIsElementType to get the types(alternatively, you can use the WhereElementIsNotElementType)
                collector.WhereElementIsElementType();

                // Create SHeets; in Revit API, a sheet is called a ViewSheet
                // Using the collector from above with FirstElementId to choose the first titleblock type for use
                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());

                // once the sheets are created(above) we can reference them with SheetNumber and Name
                // curSheet.SheetNumber = "A1010101";
                // curSheet.Name = "Name Sheet";

                // Commit the transaction
                t.Commit();
            }


            // Close Excel
            excelWB.Close();
            // Quit Excel
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}
