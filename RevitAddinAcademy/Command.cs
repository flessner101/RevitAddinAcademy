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

            string excelFile = @"C:\Users\cflessner\Desktop\Revit Addin Class\Session02_CombinationSheetList-220706-171323.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(excelFile);
            Excel.Worksheet excelWSLevels = excelWB.Worksheets.Item[1];
            Excel.Worksheet excelWSSheets = excelWB.Worksheets.Item[2];
            Excel.Range excelRngLevels = excelWSLevels.UsedRange;
            Excel.Range excelRngSheets = excelWSSheets.UsedRange;


            int rowCountLevels = excelRngLevels.Rows.Count;
            int rowCountSheets = excelRngSheets.Rows.Count;


            List<string[]> dataListLevelName = new List<string[]>();
            List<double[]> dataListElevation = new List<double[]>();

            for(int i = 1; i <= rowCountLevels; i++)
            {
                Excel.Range cell1Levels = excelWSLevels.Cells[i, 1]; // Level Name
                Excel.Range cell2Levels = excelWSLevels.Cells[i, 2]; // Elevation (FT)

                Excel.Range cell1Sheets = excelWSSheets.Cells[i, 1]; // Sheet Number
                Excel.Range cell2Sheets = excelWSSheets.Cells[i, 2]; // Sheet Name

                string data1 = cell1Levels.Value.ToString(); // Level Name
                double data2 = cell2Levels.Value;            // Elevation (FT)
                
                string[] dataArray1 = new string[2]; // Level Name
                double[] dataArray2 = new double[2]; // Elevation (FT)

                dataArray1[0] = data1; // Level Name
                dataArray2[1] = data2; // Elevation

                dataListLevelName.Add(dataArray1); // Level Name
                dataListElevation.Add(dataArray2); // Elevation

            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create some Revit Stuff");


                Level curLevel = Level.Create(doc, 100);

                FilteredElementCollector collector = new FilteredElementCollector(doc);

                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());

                curSheet.SheetNumber = "A1010101";
                curSheet.Name = "Name Sheet";

                

                t.Commit();
            }

            excelWB.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}
