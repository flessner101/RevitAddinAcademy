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
    public class cmdProjectSetup : IExternalCommand
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

            string excelFile = @"J:\New Hire\Cameron Flessner\Revit Add-in Academy\Session 02\Resources\Session02_Challenge.xlsx";

            // open Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWs2 = excelWb.Worksheets.Item[2];

            Excel.Range excelRng1 = excelWs1.UsedRange;
            Excel.Range excelRng2 = excelWs2.UsedRange;

            int rowCount1 = excelRng1.Rows.Count;
            int rowCount2 = excelRng2.Rows.Count;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Setup project");
                
                for (int i = 2; i <= rowCount1; i++)
                {
                    Excel.Range levelData1 = excelWs1.Cells[i, 1];
                    Excel.Range levelData2 = excelWs1.Cells[i, 2];

                    string levelName = levelData1.Value.ToString();
                    double levelElev = levelData2.Value;

                    Level newLevel = Level.Create(doc, levelElev);
                    newLevel.Name = levelName;

                }

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                for (int j = 2; j <= rowCount2; j++)
                {
                    Excel.Range sheetData1 = excelWs2.Cells[j, 1];
                    Excel.Range sheetData2 = excelWs2.Cells[j, 2];

                    string sheetNum = sheetData1.Value.ToString();
                    string sheetName = sheetData2.Value.ToString();

                    ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    newSheet.SheetNumber = sheetNum;
                    newSheet.Name = sheetName;
                }
                t.Commit();
            }
            
            excelWb.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}