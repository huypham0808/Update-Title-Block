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
using System.Data.OleDb;

namespace UpdateAttribute
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();
            dataDrawingGrid.AllowUserToAddRows = false;
            dataDrawingGrid.AllowUserToDeleteRows = false;
            dataDrawingGrid.RowHeadersVisible = false;
            dataDrawingGrid.EnableHeadersVisualStyles = false;
            dataDrawingGrid.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Orange;
        }
        public void LoadDataFromExcel(string fpath, string ext, string hdr)
        {
            try
            {
                string con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
                con = String.Format(con, fpath, hdr);
                OleDbConnection excelcon = new OleDbConnection(con);
                excelcon.Open();
                System.Data.DataTable excelData = excelcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string exsheetName = excelData.Rows[0]["TABLE_NAME"].ToString();
                OleDbCommand com = new OleDbCommand("Select * from [" + exsheetName + "]", excelcon);
                OleDbDataAdapter oda = new OleDbDataAdapter(com);
                System.Data.DataTable dt = new System.Data.DataTable();
                oda.Fill(dt);
                excelcon.Close();
                dataDrawingGrid.DataSource = dt;
                dataDrawingGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
            catch
            {
                MessageBox.Show("Something went wrong. Please try again!","Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }
        public string TitleBlockName()
        {
            string titleBlockName;
            string gender = cbbTitleBlockName.SelectedValue?.ToString();
            if(!string.IsNullOrEmpty(gender))
            {
                titleBlockName = gender;
            }
            else
            {
                titleBlockName = "STN_TITLE BOX 11x17";
            }
            return titleBlockName;
        }
        private void btnSelectExcel_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ofD = new OpenFileDialog();
            ofD.Filter = "Excel Files(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            ofD.Multiselect = false;
            if (ofD.ShowDialog() != DialogResult.OK) return;
            string filePath = ofD.FileName;
            txtExcelPath.Text = filePath;
            LoadDataFromExcel(filePath, ".xlsx", "yes");
        }
        private void btnUpdateInfor_Click(object sender, EventArgs e)
        {
            timerUpdateInfor.Start();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (DataGridViewRow row in dataDrawingGrid.Rows)
                    {
                        string layoutName = row.Cells["Layout Name"].Value.ToString();
                        string layoutIdString = row.Cells["Layout ID"].Value.ToString();
                        string projectTitle1 = row.Cells["PROJECT_TITLE1"].Value.ToString();
                        string projectTitle2 = row.Cells["SHEET_TITLE"].Value.ToString();
                        string widthFactorString1 = row.Cells["PROJECT_TITLE1 Width factor"].Value.ToString();
                        string widthFactorString2 = row.Cells["SHEET_TITLE Width factor"].Value.ToString();

                        string revLevel1 = row.Cells["REV_LEVEL_1"].Value.ToString();
                        string revDate1 = row.Cells["REV_DATE1"].Value.ToString();
                        string revDesc1 = row.Cells["REV_DESC1"].Value.ToString();
                        string revBy1 = row.Cells["BY1"].Value.ToString();

                        string revLevel2 = row.Cells["REV_LEVEL_2"].Value.ToString();
                        string revDate2 = row.Cells["REV_DATE2"].Value.ToString();
                        string revDesc2 = row.Cells["REV_DESC2"].Value.ToString();
                        string revBy2 = row.Cells["BY2"].Value.ToString();

                        string revLevel3 = row.Cells["REV_LEVEL_3"].Value.ToString();
                        string revDate3 = row.Cells["REV_DATE3"].Value.ToString();
                        string revDesc3 = row.Cells["REV_DESC3"].Value.ToString();
                        string revBy3 = row.Cells["BY3"].Value.ToString();

                        string revLevel4 = row.Cells["REV_LEVEL_4"].Value.ToString();
                        string revDate4 = row.Cells["REV_DATE4"].Value.ToString();
                        string revDesc4 = row.Cells["REV_DESC4"].Value.ToString();
                        string revBy4 = row.Cells["BY4"].Value.ToString();

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
                                doc.LockDocument();
                                Layout layout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;

                                // Switch to the layout and activate it
                                LayoutManager.Current.CurrentLayout = layout.LayoutName;

                                // Get the block table record of the current layout
                                BlockTableRecord layoutSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                                // Find the TitleBlock attribute values and update them
                                foreach (ObjectId entityId in layoutSpace)
                                {
                                    Entity entity = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;

                                    if (entity is BlockReference blockRef && blockRef.Name.Equals(TitleBlockName(), StringComparison.OrdinalIgnoreCase))
                                    {
                                        foreach (ObjectId attributeId in blockRef.AttributeCollection)
                                        {
                                            AttributeReference attribute = tr.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;

                                            if (attribute != null)
                                            {
                                                switch (attribute.Tag.ToUpper())
                                                {
                                                    case "PROJECT_TITLE1":
                                                        attribute.TextString = projectTitle1.Trim();
                                                        attribute.WidthFactor = Double.Parse(widthFactorString1);
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
            finally
            {
                MessageBox.Show("Update successfully !", "Update TitleBlock", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //this.Close();
            }
        }

        private void timerUpdateInfor_Tick(object sender, EventArgs e)
        {
            progressBarUpdateInFor.PerformStep();
        }
        //private Task ProcessDate
    }
}
