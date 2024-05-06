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

namespace UpdateAttribute
{
    public class Command
    {
        [CommandMethod("EXPTTBTOEXCEL")]
        public void ExportExcelFile()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            List<string> layoutNameList = new List<string>();
            List<ObjectId> layoutIDList = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary layoutDic = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                foreach (DBDictionaryEntry entry in layoutDic)
                {
                    layoutIDList.Add(entry.Value);
                    layoutIDList = layoutIDList.OrderBy(id => ((Layout)tr.GetObject(id, OpenMode.ForRead)).TabOrder).ToList();
                    Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase)) continue;

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
                worksheet.Cells["A1:U1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Color colFromHex = ColorTranslator.FromHtml("#B7DEE8");
                worksheet.Cells["A1:U1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#df4907"));
                worksheet.Cells["A1:U1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#fff"));
                worksheet.Cells["A1:U1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["B1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["C1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["D1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["E1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["F1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["G1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["H1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["I1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["J1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["K1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["L1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["M1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["N1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["O1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["P1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["Q1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["R1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["S1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["T1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["U1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);


                worksheet.Cells["A1"].Value = "Layout Name";
                worksheet.Cells["A1"].AutoFitColumns();

                worksheet.Cells["B1"].Value = "Layout ID";
                

                worksheet.Cells["C1"].Value = "PROJECT_TITLE1";
                

                worksheet.Cells["D1"].Value = "SHEET_TITLE";
                worksheet.Cells["D1:D1000"].AutoFitColumns();

                worksheet.Cells["E1"].Value = "PROJECT_TITLE1 Width factor";
                worksheet.Cells["E1:E1000"].AutoFitColumns();

                worksheet.Cells["F1"].Value = "SHEET_TITLE Width factor";
                worksheet.Cells["F1:F1000"].AutoFitColumns();

                worksheet.Cells["G1"].Value = "REV_LEVEL_1";
                worksheet.Cells["H1"].Value = "REV_DATE1";
                worksheet.Cells["I1"].Value = "REV_DESC1";
                worksheet.Cells["J1"].Value = "BY";

                worksheet.Cells["K1"].Value = "REV_LEVEL_2";
                worksheet.Cells["L1"].Value = "REV_DATE2";
                worksheet.Cells["M1"].Value = "REV_DESC2";
                worksheet.Cells["N1"].Value = "BY";

                worksheet.Cells["O1"].Value = "REV_LEVEL_3";
                worksheet.Cells["P1"].Value = "REV_DATE3";
                worksheet.Cells["Q1"].Value = "REV_DESC3";
                worksheet.Cells["R1"].Value = "BY";

                worksheet.Cells["S1"].Value = "REV_LEVEL_4";
                worksheet.Cells["T1"].Value = "REV_DATE4";
                worksheet.Cells["U1"].Value = "REV_DESC4";
                worksheet.Cells["V1"].Value = "BY";

                for (int i = 0; i < layoutIDList.Count; i++)
                {
                    ObjectId layoutId = layoutIDList[i];

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                        // Switch to the layout and activate it
                        LayoutManager.Current.CurrentLayout = layout.LayoutName;

                        // Get the block table record of the current layout
                        BlockTableRecord layoutSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                        // Find the TitleBlock attribute values
                        string projectTitle1 = string.Empty;
                        string projectTitle2 = string.Empty;
                        double widthFactorTitle1 = 0.0;
                        double widthFactorTitle2 = 0.0;

                        foreach (ObjectId entityId in layoutSpace)
                        {
                            Entity entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;

                            if (entity is BlockReference blockRef && blockRef.Name.Equals("STN_TITLE BOX 11x17", StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                {
                                    AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;

                                    if (attribute != null && attribute.Tag.ToUpper() == "PROJECT_TITLE1")
                                    {
                                        projectTitle1 = attribute.TextString;
                                        widthFactorTitle1 = attribute.WidthFactor;
                                    }
                                    else if (attribute != null && attribute.Tag.ToUpper() == "SHEET_TITLE")
                                    {
                                        projectTitle2 = attribute.TextString;
                                        widthFactorTitle2= attribute.WidthFactor;
                                    }
                                }
                            }
                        }
                        //DBDictionary layoutDic = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        // Write layout name, layout ID, and attribute values to Excel worksheet
                        worksheet.Cells[i + 2, 1].Value = layout.LayoutName;
                        worksheet.Cells[i + 2, 2].Value = "Handle: " + layoutId.Handle.ToString();
                        worksheet.Cells[i + 2, 3].Value = projectTitle1;
                        worksheet.Cells[i + 2, 4].Value = projectTitle2;
                        worksheet.Cells[i + 2, 5].Value = widthFactorTitle1.ToString();
                        worksheet.Cells[i + 2, 6].Value = widthFactorTitle2.ToString();

                        worksheet.Cells["B3:D1000"].AutoFitColumns();
                        //Add border for cells
                        worksheet.Cells[i + 3, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 9].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 10].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 11].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 12].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 13].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 14].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 15].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 16].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 17].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 18].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 19].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 20].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 21].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[i + 3, 22].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);




                        tr.Commit();
                    }
                }
                var trans = db.TransactionManager.StartTransaction();
                DBDictionary layoutDic = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TitleBlock.xlsx");
                FileInfo excelFile = new FileInfo(excelFilePath);
                excelPackage.SaveAs(excelFile);
                ed.WriteMessage($"\nLayout attributes exported to: {excelFilePath}");
                MessageBox.Show("Export successfully total " +((layoutDic.Count) - 1).ToString() + " layouts" );
                //FileInfo fi = new FileInfo(excelFilePath);
                //if(fi.Exists)
                //{
                //    System.Diagnostics.Process.Start(excelFilePath);
                //}    
                //else
                //{
                //    MessageBox.Show("File doesn't exist");
                //}
                System.Diagnostics.Process.Start(excelFilePath);
            }           
        }
        [CommandMethod("IMPORTTTBEXCEL")]
        public void ImportExcelFile()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Prompt for the Excel file to load
            PromptOpenFileOptions promptOptions = new PromptOpenFileOptions("Select Excel File");
            promptOptions.Filter = "Excel Files (*.xlsx)|*.xlsx";

            PromptFileNameResult promptResult = ed.GetFileNameForOpen(promptOptions);
            if (promptResult.Status != PromptStatus.OK)
                return;

            string excelFilePath = promptResult.StringResult;

            // Load the Excel file and update attribute values
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Title Block"];
                if (worksheet == null)
                {
                    ed.WriteMessage("\nThe 'LayoutAttributes' worksheet was not found in the Excel file.");
                    return;
                }
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var cell in worksheet.Cells["A2:D" + worksheet.Dimension.End.Row])
                    {
                        string layoutName = cell.Offset(0, 0).Text;
                        string layoutIdString = cell.Offset(0, 1).Text;
                        string projectTitle1 = cell.Offset(0, 2).Text;
                        string projectTitle2 = cell.Offset(0, 3).Text;
                        string widthFactorString1 = cell.Offset(0, 4).Text;
                        string widthFactorString2 = cell.Offset(0, 5).Text;

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
                            else if (layoutIdString.StartsWith("ObjectId:"))
                            {
                                // Convert ObjectId string to ObjectId
                                string objectIdValue = layoutIdString.Replace("ObjectId:", "").Trim();
                                layoutId = new ObjectId((IntPtr)Convert.ToInt64(objectIdValue, 16));
                            }

                            if (layoutId.IsValid && layoutId.ObjectClass == RXObject.GetClass(typeof(Layout)))
                            {
                                Layout layout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;

                                // Switch to the layout and activate it
                                LayoutManager.Current.CurrentLayout = layout.LayoutName;

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
                                                if (attribute.Tag.ToUpper() == "PROJECT_TITLE1")
                                                {
                                                    attribute.TextString = projectTitle1;
                                                    if (!string.IsNullOrEmpty(widthFactorString1))
                                                    {
                                                        double widthFactor1;
                                                        if (double.TryParse(widthFactorString1, out widthFactor1))
                                                        {
                                                            attribute.WidthFactor = widthFactor1;
                                                        }
                                                        else
                                                        {
                                                            ed.WriteMessage("\nInvalid width factor value for attribute 'PROJECT_TITLE2'.");
                                                        }
                                                    }
                                                }
                                                else if (attribute.Tag.ToUpper() == "SHEET_TITLE")
                                                {
                                                    attribute.TextString = projectTitle2;
                                                    if (!string.IsNullOrEmpty(widthFactorString2))
                                                    {
                                                        double widthFactor2;
                                                        if (double.TryParse(widthFactorString2, out widthFactor2))
                                                        {
                                                            attribute.WidthFactor = widthFactor2;
                                                        }
                                                        else
                                                        {
                                                            ed.WriteMessage("\nInvalid width factor value for attribute 'PROJECT_TITLE2'.");
                                                        }
                                                    }
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
            //ed.WriteMessage("\nLayout attributes updated from Excel file.");
            MessageBox.Show("Update successfully for");
        }
        [CommandMethod("CallUpdateTTB")]
        public void CallForm()
        {
            mainForm mf = new mainForm();
            mf.ShowDialog();
        }
    }
    
}
