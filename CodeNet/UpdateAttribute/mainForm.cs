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
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(txtExcelPath.Text)))
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
                                                else if (attribute.Tag.ToUpper() == "PROJECT_TITLE2")
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
            MessageBox.Show("Update successfully for");
        }
    }
}
