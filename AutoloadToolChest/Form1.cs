using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace AutoLoadToolChest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string vbscriptPath = @"C:\Users\ad\Desktop\MY TOOL\Test3.vbs"; // Path to your VBScript file

            if (System.IO.File.Exists(vbscriptPath))
            {
                Process.Start(vbscriptPath);
            }
            else
            {
                MessageBox.Show("VBScript file not found at the specified location.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
