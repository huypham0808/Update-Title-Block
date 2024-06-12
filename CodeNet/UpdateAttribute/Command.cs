using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using OfficeOpenXml;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Runtime.InteropServices;
using _Excel = Microsoft.Office.Interop.Excel;

namespace UpdateAttribute
{
    public class Command
    {
        _Excel.Application xlApp;
        _Excel.Workbook xlWorkBook;
        _Excel.Worksheet xlWorkSheet;

        [CommandMethod("EXPTTBTOEXCEL")]
        public void ExportExcelFile()
        {
            ProcessForm pF = new ProcessForm();
            Application.ShowModelessDialog(pF);
            
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            List<string> layoutNameList = new List<string>();
            List<ObjectId> layoutIDList = new List<ObjectId>();
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutDic = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    foreach (DBDictionaryEntry entry in layoutDic)
                    {
                        layoutIDList.Add(entry.Value);
                        layoutIDList = layoutIDList.OrderBy(id => ((Layout)tr.GetObject(id, OpenMode.ForRead)).TabOrder).ToList();
                        Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                        if (layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                        //if (layout.LayoutName == "Model") continue;
                        layoutNameList.Add(layout.LayoutName);
                    }
                    tr.Commit();
                }
                // Create Excel package and worksheet
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Title Block");
                    //Header format
                    worksheet.Cells["A1:V1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //Color colFromHex = ColorTranslator.FromHtml("#B7DEE8");
                    worksheet.Cells["A1:V1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#df4907"));
                    worksheet.Cells["A1:V1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#fff"));
                    worksheet.Cells["A1:V1"].Style.Font.Bold = true;
                    worksheet.Cells["A1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["B1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["C1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["D1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["E1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["F1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["G1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["H1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["I1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["J1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["K1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["L1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["M1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["N1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["O1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["P1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["Q1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["R1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["S1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["T1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["U1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    worksheet.Cells["V1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);


                    worksheet.Cells["A1"].Value = "Layout Name";
                    worksheet.Cells["A1"].AutoFitColumns();

                    worksheet.Cells["B1"].Value = "Layout ID";


                    worksheet.Cells["C1"].Value = "PROJECT_TITLE1";


                    worksheet.Cells["D1"].Value = "SHEET_TITLE";
                    worksheet.Cells["D1:D100"].AutoFitColumns();

                    worksheet.Cells["E1"].Value = "PROJECT_TITLE1 Width factor";
                    worksheet.Cells["E1:E100"].AutoFitColumns();

                    worksheet.Cells["F1"].Value = "SHEET_TITLE Width factor";
                    worksheet.Cells["F1:F100"].AutoFitColumns();

                    worksheet.Cells["G1"].Value = "REV_LEVEL_1";
                    worksheet.Cells["G1:G100"].AutoFitColumns();

                    worksheet.Cells["H1"].Value = "REV_DATE1";
                    worksheet.Cells["H1:H100"].AutoFitColumns();

                    worksheet.Cells["I1"].Value = "REV_DESC1";
                    worksheet.Cells["I1:I100"].AutoFitColumns();

                    worksheet.Cells["J1"].Value = "BY1";

                    worksheet.Cells["K1"].Value = "REV_LEVEL_2";
                    worksheet.Cells["L1"].Value = "REV_DATE2";
                    worksheet.Cells["M1"].Value = "REV_DESC2";
                    worksheet.Cells["N1"].Value = "BY2";

                    worksheet.Cells["O1"].Value = "REV_LEVEL_3";
                    worksheet.Cells["P1"].Value = "REV_DATE3";
                    worksheet.Cells["Q1"].Value = "REV_DESC3";
                    worksheet.Cells["R1"].Value = "BY3";

                    worksheet.Cells["S1"].Value = "REV_LEVEL_4";
                    worksheet.Cells["T1"].Value = "REV_DATE4";
                    worksheet.Cells["U1"].Value = "REV_DESC4";
                    worksheet.Cells["V1"].Value = "BY4";

                    // Find the TitleBlock attribute values
                    string projectTitle1 = string.Empty;
                    string projectTitle2 = string.Empty;
                    double widthFactorTitle1 = 0.0;
                    double widthFactorTitle2 = 0.0;

                    string revLevel1 = string.Empty;
                    string revDate1 = string.Empty;
                    string revDesc1 = string.Empty;
                    string revBy1 = string.Empty;

                    string revLevel2 = string.Empty;
                    string revDate2 = string.Empty;
                    string revDesc2 = string.Empty;
                    string revBy2 = string.Empty;

                    string revLevel3 = string.Empty;
                    string revDate3 = string.Empty;
                    string revDesc3 = string.Empty;
                    string revBy3 = string.Empty;

                    string revLevel4 = string.Empty;
                    string revDate4 = string.Empty;
                    string revDesc4 = string.Empty;
                    string revBy4 = string.Empty;

                    for (int i = 0; i < layoutIDList.Count; i++)
                    {
                        ObjectId layoutId = layoutIDList[i];

                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                            // Switch to the layout and activate it
                            LayoutManager.Current.CurrentLayout = layout.LayoutName;
                            if (layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;
                            // Get the block table record of the current layout
                            BlockTableRecord layoutSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                            foreach (ObjectId entityId in layoutSpace)
                            {
                                Entity entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;

                                if (entity is BlockReference blockRef && blockRef.Name.Equals("STN_TITLE BOX 11x17", StringComparison.OrdinalIgnoreCase))
                                {
                                    foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                    {
                                        AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;
                                        if (attribute != null)
                                        {
                                            switch (attribute.Tag.ToUpper())
                                            {
                                                case "PROJECT_TITLE1":
                                                    projectTitle1 = attribute.TextString;
                                                    widthFactorTitle1 = attribute.WidthFactor;
                                                    break;
                                                case "SHEET_TITLE":
                                                    projectTitle2 = attribute.TextString;
                                                    widthFactorTitle2 = attribute.WidthFactor;
                                                    break;
                                                //Level 1
                                                case "REV_LEVEL1":
                                                    revLevel1 = attribute.TextString;
                                                    break;
                                                case "REV_DATE1":
                                                    revDate1 = attribute.TextString;
                                                    break;
                                                case "REV_DESC1":
                                                    revDesc1 = attribute.TextString;
                                                    break;
                                                case "REV_BY1":
                                                    revBy1 = attribute.TextString;
                                                    break;
                                                //Level 2
                                                case "REV_LEVEL2":
                                                    revLevel2 = attribute.TextString;
                                                    break;
                                                case "REV_DATE2":
                                                    revDate2 = attribute.TextString;
                                                    break;
                                                case "REV_DESC2":
                                                    revDesc2 = attribute.TextString;
                                                    break;
                                                case "REV_BY2":
                                                    revBy2 = attribute.TextString;
                                                    break;
                                                //Level 3
                                                case "REV_LEVEL3":
                                                    revLevel3 = attribute.TextString;
                                                    break;
                                                case "REV_DATE3":
                                                    revDate3 = attribute.TextString;
                                                    break;
                                                case "REV_DESC3":
                                                    revDesc3 = attribute.TextString;
                                                    break;
                                                case "REV_BY3":
                                                    revBy3 = attribute.TextString;
                                                    break;
                                                //Level 4
                                                case "REV_LEVEL4":
                                                    revLevel4 = attribute.TextString;
                                                    break;
                                                case "REV_DATE4":
                                                    revDate4 = attribute.TextString;
                                                    break;
                                                case "REV_DESC4":
                                                    revDesc4 = attribute.TextString;
                                                    break;
                                                case "REV_BY4":
                                                    revBy4 = attribute.TextString;
                                                    break;
                                            }
                                        }
                                    }
                                }
                            }
                            //DBDictionary layoutDic = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            // Write layout name, layout ID, and attribute values to Excel worksheet
                            worksheet.Cells[i + 1, 1].Value = layout.LayoutName;
                            worksheet.Cells[i + 1, 2].Value = "Handle: " + layoutId.Handle.ToString();
                            worksheet.Cells[i + 1, 3].Value = projectTitle1;
                            worksheet.Cells[i + 1, 4].Value = projectTitle2;
                            worksheet.Cells[i + 1, 5].Value = widthFactorTitle1.ToString();
                            worksheet.Cells[i + 1, 6].Value = widthFactorTitle2.ToString();

                            worksheet.Cells[i + 1, 7].Value = revLevel1.ToString();
                            worksheet.Cells[i + 1, 8].Value = revDate1.ToString();
                            worksheet.Cells[i + 1, 9].Value = revDesc1.ToString();
                            worksheet.Cells[i + 1, 10].Value = revBy1.ToString();

                            worksheet.Cells[i + 1, 11].Value = revLevel2.ToString();
                            worksheet.Cells[i + 1, 12].Value = revDate2.ToString();
                            worksheet.Cells[i + 1, 13].Value = revDesc2.ToString();
                            worksheet.Cells[i + 1, 14].Value = revBy2.ToString();

                            worksheet.Cells[i + 1, 15].Value = revLevel3.ToString();
                            worksheet.Cells[i + 1, 16].Value = revDate3.ToString();
                            worksheet.Cells[i + 1, 17].Value = revDesc3.ToString();
                            worksheet.Cells[i + 1, 18].Value = revBy3.ToString();

                            worksheet.Cells[i + 1, 19].Value = revLevel4.ToString();
                            worksheet.Cells[i + 1, 20].Value = revDate4.ToString();
                            worksheet.Cells[i + 1, 21].Value = revDesc4.ToString();
                            worksheet.Cells[i + 1, 22].Value = revBy4.ToString();

                            worksheet.Cells["B3:D100"].AutoFitColumns();
                            //Add border for cells
                            worksheet.Cells[i + 1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 9].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 10].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 11].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 12].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 13].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 14].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 15].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 16].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 17].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 18].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 19].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 20].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 21].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            worksheet.Cells[i + 1, 22].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                            tr.Commit();
                            tr.Dispose();

                        }
                    }
                    //var trans = db.TransactionManager.StartTransaction();
                    //DBDictionary layoutDic = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TitleBlock Information.xlsx");
                    FileInfo excelFile = new FileInfo(excelFilePath);
                    excelPackage.SaveAs(excelFile);
                    ed.WriteMessage($"\nLayout attributes exported to: {excelFilePath}");
                    //MessageBox.Show("Export successfully total " + ((layoutDic.Count) - 1).ToString() + " layouts", "Layout Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //trans.Dispose();
                    //System.Diagnostics.Process.Start(excelFilePath);
                }
            }
            catch
            {
                MessageBox.Show("Somethings wrong. Please try again", "Export data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        
        }
        [CommandMethod("IMPORTTTBEXCEL")]
        public void ImportExcelFile()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt for the Excel file to load
            PromptOpenFileOptions promptOptions = new PromptOpenFileOptions("Select Excel File");
            promptOptions.Filter = "Excel Files (*.xlsx)|*.xlsx";

            PromptFileNameResult promptResult = ed.GetFileNameForOpen(promptOptions);
            if (promptResult.Status != PromptStatus.OK)
                return;

            string excelFilePath = promptResult.StringResult;

            // Load the Excel file and update attribute values
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Title Block"];
                    if (worksheet == null)
                    {
                        ed.WriteMessage("\nThe 'Title Block' worksheet was not found in the Excel file.");
                        return;
                    }
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        foreach (var cell in worksheet.Cells["A2:V" + worksheet.Dimension.End.Row])
                        {
                            string layoutName = cell.Offset(0, 0).Text;
                            string layoutIdString = cell.Offset(0, 1).Text;
                            string projectTitle1 = cell.Offset(0, 2).Text;
                            string projectTitle2 = cell.Offset(0, 3).Text;
                            string widthFactorString1 = cell.Offset(0, 4).Text;
                            string widthFactorString2 = cell.Offset(0, 5).Text;

                            string revLevel1 = cell.Offset(0, 6).Text;
                            string revDate1 = cell.Offset(0, 7).Text;
                            string revDesc1 = cell.Offset(0, 8).Text;
                            string revBy1 = cell.Offset(0, 9).Text;

                            string revLevel2 = cell.Offset(0, 10).Text;
                            string revDate2 = cell.Offset(0, 11).Text;
                            string revDesc2 = cell.Offset(0, 12).Text;
                            string revBy2 = cell.Offset(0, 13).Text;

                            string revLevel3 = cell.Offset(0, 14).Text;
                            string revDate3 = cell.Offset(0, 15).Text;
                            string revDesc3 = cell.Offset(0, 16).Text;
                            string revBy3 = cell.Offset(0, 17).Text;

                            string revLevel4 = cell.Offset(0, 18).Text;
                            string revDate4 = cell.Offset(0, 19).Text;
                            string revDesc4 = cell.Offset(0, 20).Text;
                            string revBy4 = cell.Offset(0, 21).Text;

                            if (!string.IsNullOrEmpty(layoutName) && !string.IsNullOrEmpty(layoutIdString))
                            {
                                ObjectId layoutId = ObjectId.Null;

                                if (layoutIdString.StartsWith("Handle:"))
                                {
                                    // Convert Handle to ObjectId
                                    string handleValue = layoutIdString.Replace("Handle:", "").Trim();
                                    Handle handle = new Handle(Convert.ToInt64(handleValue, 16));
                                    layoutId = db.GetObjectId(false, handle, 0);
                                }
                                if (layoutId.IsValid && layoutId.ObjectClass == RXObject.GetClass(typeof(Layout)))
                                {
                                    Layout layout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;

                                    // Switch to the layout and activate it
                                    //LayoutManager.Current.CurrentLayout = layout.LayoutName;

                                    // Get the block table record of the current layout
                                    BlockTableRecord layoutSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                                    // Find the TitleBlock attribute values and update them
                                    foreach (ObjectId entityId in layoutSpace)
                                    {
                                        Entity entity = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;

                                        if (entity is BlockReference blockRef && blockRef.Name.Equals("STN_TITLE BOX 11x17", StringComparison.OrdinalIgnoreCase))
                                        {
                                            foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                            {
                                                AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;

                                                if (attribute != null)
                                                {
                                                    //double widthFactor1a;
                                                    switch (attribute.Tag.ToUpper())
                                                    {
                                                        case "PROJECT_TITLE1":
                                                            attribute.TextString = projectTitle1.Trim();
                                                            //attribute.WidthFactor = widthFactor1a;
                                                            break;
                                                        case "SHEET_TITLE":
                                                            attribute.TextString = projectTitle2.Trim();
                                                            attribute.WidthFactor = Double.Parse(widthFactorString2);
                                                            break;
                                                        case "REV_LEVEL1":
                                                            attribute.TextString = revLevel1.Trim();
                                                            break;
                                                        case "REV_DATE1":
                                                            attribute.TextString = revDate1.Trim();
                                                            break;
                                                        case "REV_DESC1":
                                                            attribute.TextString = revDesc1.Trim().ToUpper();
                                                            break;
                                                        case "REV_BY1":
                                                            attribute.TextString = revBy1.Trim().ToUpper();
                                                            break;
                                                        //Level 2
                                                        case "REV_LEVEL2":
                                                            attribute.TextString = revLevel2;
                                                            break;
                                                        case "REV_DATE2":
                                                            attribute.TextString = revDate2;
                                                            break;
                                                        case "REV_DESC2":
                                                            attribute.TextString = revDesc2.Trim().ToUpper();
                                                            break;
                                                        case "REV_BY2":
                                                            attribute.TextString = revBy2.Trim().ToUpper();
                                                            break;
                                                        //Level 3
                                                        case "REV_LEVEL3":
                                                            attribute.TextString = revLevel3;
                                                            break;
                                                        case "REV_DATE3":
                                                            attribute.TextString = revDate3;
                                                            break;
                                                        case "REV_DESC3":
                                                            attribute.TextString = revDesc3.Trim().ToUpper();
                                                            break;
                                                        case "REV_BY3":
                                                            attribute.TextString = revBy3.Trim().ToUpper();
                                                            break;
                                                        //Level 4
                                                        case "REV_LEVEL4":
                                                            attribute.TextString = revLevel4;
                                                            break;
                                                        case "REV_DATE4":
                                                            attribute.TextString = revDate4;
                                                            break;
                                                        case "REV_DESC4":
                                                            attribute.TextString = revDesc4.Trim().ToUpper();
                                                            break;
                                                        case "REV_BY4":
                                                            attribute.TextString = revBy4.Trim().ToUpper();
                                                            break;
                                                        default:
                                                            break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        tr.Commit();
                   
                    }
                }
                UnloadExcel();
                MessageBox.Show("Update successfully");
            }
            catch
            {
                UnloadExcel();
                ed.WriteMessage("Fail");
                return;
            }
          
        }
        private void UnloadExcel()
        {
            try
            {
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch { }
        }
        [CommandMethod("CallUpdateTTB")]
        public void CallForm()
        {
            mainForm mf = new mainForm();
            Application.ShowModelessDialog(mf);
        }
        [CommandMethod ("CallExportTTB")]
        public void CallExportForm()
        {
            ProcessForm pcf = new ProcessForm();
            pcf.ShowDialog();
            //Application.ShowModelessDialog(pcf);
        }

    }
    
}
