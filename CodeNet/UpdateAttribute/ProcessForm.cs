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
    public partial class ProcessForm : Form
    {
        public ProcessForm()
        {
            InitializeComponent();
            timerFormProcess.Enabled = false;
        }

        private void timerFormProcess_Tick(object sender, EventArgs e)
        {
            //var doc = AcAp.DocumentManager.MdiActiveDocument;
            //var db = doc.Database;
            //var ed = doc.Editor;
            //var trans = db.TransactionManager.StartTransaction();
            //DBDictionary layoutDic = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            //progressBarForm.Value = 1;
            //s++;
            //lblPercent.Text = s + "%";
            progressBarForm.Increment(1);
            lblPercent.Text = progressBarForm.Value.ToString() + "%";
            if(progressBarForm.Value == 100)
            {
                //timerFormProcess.Enabled = false;
                timerFormProcess.Stop();
                //lblLoadingStatus.Text = "Export Successfully " + ((layoutDic.Count) - 1).ToString() + " layouts";
                lblLoadingStatus.Text = "Export Successfully ";
                lblLoadingStatus.ForeColor = Color.Green;
            }
            //this.Hide();
        }
        private void btnCloseProcess_Click(object sender, EventArgs e)
        {
            string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TitleBlock Information.xlsx"); 
            DialogResult confirmPopUp = MessageBox.Show("Export Successfully", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (confirmPopUp == DialogResult.OK)
            {
                System.Diagnostics.Process.Start(excelFilePath);
                this.Close();
            }
            else
            {
                this.Close();
            }
            //MessageBox.Show("Export Successfully", "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //System.Diagnostics.Process.Start(excelFilePath);
            //this.Close();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {       
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            List<string> layoutNameList = new List<string>();
            List<ObjectId> layoutIDList = new List<ObjectId>();
            try
            {
                timerFormProcess.Start();
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
    }
}
