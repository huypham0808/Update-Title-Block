using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace UpdateAttribute
{
    public partial class ProcessForm : Form
    {
        public ProcessForm()
        {
            InitializeComponent();
        }

        private void timerFormProcess_Tick(object sender, EventArgs e)
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var trans = db.TransactionManager.StartTransaction();
            DBDictionary layoutDic = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

            progressBarForm.PerformStep();
            if(progressBarForm.Maximum == 100)
            {
                lblLoadingStatus.Text = "Export Successfully " + ((layoutDic.Count) - 1).ToString() + " layouts";
                lblLoadingStatus.ForeColor = Color.Green;
            }
            //this.Hide();
        }

        private void ProcessForm_Load(object sender, EventArgs e)
        {
            timerFormProcess.Start();
        }

        private void btnCloseProcess_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
