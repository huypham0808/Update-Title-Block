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

namespace UpdateAttribute
{
    public class Command
    {
        [CommandMethod("TESTEXPTOEXCEL")]
        public void Test()
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
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("LayoutAttributes");
                worksheet.Cells["A1"].Value = "Layout Name";
                worksheet.Cells["B1"].Value = "Layout ID";
                worksheet.Cells["C1"].Value = "PROJECT_TITLE1";
                worksheet.Cells["D1"].Value = "PROJECT_TITLE2";

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

                        foreach (ObjectId entityId in layoutSpace)
                        {
                            Entity entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;

                            if (entity is BlockReference blockRef && blockRef.Name.Equals("STN_TITLE BOX 24x36", StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                {
                                    AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForRead) as AttributeReference;

                                    if (attribute != null && attribute.Tag.ToUpper() == "PROJECT_TITLE1")
                                    {
                                        projectTitle1 = attribute.TextString;
                                    }
                                    else if (attribute != null && attribute.Tag.ToUpper() == "PROJECT_TITLE2")
                                    {
                                        projectTitle2 = attribute.TextString;
                                    }
                                }
                            }
                        }

                        // Write layout name, layout ID, and attribute values to Excel worksheet
                        worksheet.Cells[i + 2, 1].Value = layout.LayoutName;
                        worksheet.Cells[i + 2, 2].Value = layoutId.Handle.ToString();
                        worksheet.Cells[i + 2, 3].Value = projectTitle1;
                        worksheet.Cells[i + 2, 4].Value = projectTitle2;

                        tr.Commit();
                    }
                }
                string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "LayoutAttributes.xlsx");
                FileInfo excelFile = new FileInfo(excelFilePath);
                excelPackage.SaveAs(excelFile);
                ed.WriteMessage($"\nLayout attributes exported to: {excelFilePath}");
            }
            
        }
        [CommandMethod("TESTIMPORTEXCEL")]
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
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["LayoutAttributes"];
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

                                    if (entity is BlockReference blockRef && blockRef.Name.Equals("STN_TITLE BOX 24x36", StringComparison.OrdinalIgnoreCase))
                                    {
                                        foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                        {
                                            AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;

                                            if (attribute != null)
                                            {
                                                if (attribute.Tag.ToUpper() == "PROJECT_TITLE1")
                                                {
                                                    attribute.TextString = projectTitle1;
                                                }
                                                else if (attribute.Tag.ToUpper() == "PROJECT_TITLE2")
                                                {
                                                    attribute.TextString = projectTitle2;
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

            ed.WriteMessage("\nLayout attributes updated from Excel file.");
        }
    }
    
}
