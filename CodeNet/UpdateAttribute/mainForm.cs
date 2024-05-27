using System;
using System.Windows.Forms;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace UpdateAttribute
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofD = new OpenFileDialog();
            ofD.Filter = "Excel Files|*.xlsx";
            ofD.Multiselect = false;
            if (ofD.ShowDialog() != DialogResult.OK) return;
            string filePath = ofD.FileName;
            txtExcelPath.Text = filePath;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Load the Excel file and update attribute values
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(txtExcelPath.Text)))
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
                                //else if (layoutIdString.StartsWith("ObjectId:"))
                                //{
                                //    // Convert ObjectId string to ObjectId
                                //    string objectIdValue = layoutIdString.Replace("ObjectId:", "").Trim();
                                //    layoutId = new ObjectId((IntPtr)Convert.ToInt64(objectIdValue, 16));
                                //}

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
                                                    switch (attribute.Tag.ToUpper())
                                                    {
                                                        case "PROJECT_TITLE1":
                                                            attribute.TextString = projectTitle1;
                                                            attribute.WidthFactor = Double.Parse(widthFactorString1);
                                                            break;
                                                        case "SHEET_TITLE":
                                                            attribute.TextString = projectTitle2;
                                                            attribute.WidthFactor = Double.Parse(widthFactorString2);
                                                            break;
                                                        //Level 1
                                                        case "REV_LEVEL1":
                                                            attribute.TextString = revLevel1;
                                                            break;
                                                        case "REV_DATE1":
                                                            attribute.TextString = revDate1;
                                                            break;
                                                        case "REV_DESC1":
                                                            attribute.TextString = revDesc1;
                                                            break;
                                                        case "REV_BY1":
                                                            attribute.TextString = revBy1;
                                                            break;
                                                        //Level 2
                                                        case "REV_LEVEL2":
                                                            attribute.TextString = revLevel2;
                                                            break;
                                                        case "REV_DATE2":
                                                            attribute.TextString = revDate2;
                                                            break;
                                                        case "REV_DESC2":
                                                            attribute.TextString = revDesc2;
                                                            break;
                                                        case "REV_BY2":
                                                            attribute.TextString = revBy2;
                                                            break;
                                                        //Level 3
                                                        case "REV_LEVEL3":
                                                            attribute.TextString = revLevel3;
                                                            break;
                                                        case "REV_DATE3":
                                                            attribute.TextString = revDate3;
                                                            break;
                                                        case "REV_DESC3":
                                                            attribute.TextString = revDesc3;
                                                            break;
                                                        case "REV_BY3":
                                                            attribute.TextString = revBy3;
                                                            break;
                                                        //Level 4
                                                        case "REV_LEVEL4":
                                                            attribute.TextString = revLevel4;
                                                            break;
                                                        case "REV_DATE4":
                                                            attribute.TextString = revDate4;
                                                            break;
                                                        case "REV_DESC4":
                                                            attribute.TextString = revDesc4;
                                                            break;
                                                        case "REV_BY4":
                                                            attribute.TextString = revBy4;
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
                MessageBox.Show("Update successfully!");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
