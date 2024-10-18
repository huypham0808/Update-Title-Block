namespace UpdateAttribute
{
    partial class mainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.lblLoading = new System.Windows.Forms.Label();
            this.progressBarUpdateInFor = new System.Windows.Forms.ProgressBar();
            this.btnUpdateInfor = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataDrawingGrid = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.cbbTitleBlockName = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelectExcel = new System.Windows.Forms.Button();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.timerUpdateInfor = new System.Windows.Forms.Timer(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataDrawingGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1046, 517);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Transparent;
            this.tabPage1.Controls.Add(this.lblLoading);
            this.tabPage1.Controls.Add(this.progressBarUpdateInFor);
            this.tabPage1.Controls.Add(this.btnUpdateInfor);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.cbbTitleBlockName);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnSelectExcel);
            this.tabPage1.Controls.Add(this.txtExcelPath);
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1038, 486);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Update Information";
            this.tabPage1.ToolTipText = "Update information of project";
            // 
            // lblLoading
            // 
            this.lblLoading.AutoSize = true;
            this.lblLoading.Location = new System.Drawing.Point(805, 437);
            this.lblLoading.Name = "lblLoading";
            this.lblLoading.Size = new System.Drawing.Size(72, 18);
            this.lblLoading.TabIndex = 17;
            this.lblLoading.Text = "Loading...";
            // 
            // progressBarUpdateInFor
            // 
            this.progressBarUpdateInFor.Location = new System.Drawing.Point(18, 430);
            this.progressBarUpdateInFor.MarqueeAnimationSpeed = 1000;
            this.progressBarUpdateInFor.Name = "progressBarUpdateInFor";
            this.progressBarUpdateInFor.Size = new System.Drawing.Size(761, 33);
            this.progressBarUpdateInFor.TabIndex = 15;
            // 
            // btnUpdateInfor
            // 
            this.btnUpdateInfor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdateInfor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnUpdateInfor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdateInfor.Location = new System.Drawing.Point(897, 424);
            this.btnUpdateInfor.Name = "btnUpdateInfor";
            this.btnUpdateInfor.Size = new System.Drawing.Size(124, 45);
            this.btnUpdateInfor.TabIndex = 14;
            this.btnUpdateInfor.Text = "UPDATE";
            this.btnUpdateInfor.UseVisualStyleBackColor = false;
            this.btnUpdateInfor.Click += new System.EventHandler(this.btnUpdateInfor_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataDrawingGrid);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 62);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1015, 348);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "TitleBlock Data";
            // 
            // dataDrawingGrid
            // 
            this.dataDrawingGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataDrawingGrid.Location = new System.Drawing.Point(6, 21);
            this.dataDrawingGrid.Name = "dataDrawingGrid";
            this.dataDrawingGrid.Size = new System.Drawing.Size(1003, 322);
            this.dataDrawingGrid.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(704, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(121, 18);
            this.label2.TabIndex = 12;
            this.label2.Text = "TitleBlock Name:";
            // 
            // cbbTitleBlockName
            // 
            this.cbbTitleBlockName.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbbTitleBlockName.FormattingEnabled = true;
            this.cbbTitleBlockName.Items.AddRange(new object[] {
            "STN_TITLE BOX 11x17",
            "STN_TITLE BOX 8.5x11",
            "STN_TITLE BOX 36x42"});
            this.cbbTitleBlockName.Location = new System.Drawing.Point(846, 14);
            this.cbbTitleBlockName.Name = "cbbTitleBlockName";
            this.cbbTitleBlockName.Size = new System.Drawing.Size(175, 26);
            this.cbbTitleBlockName.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(7, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 18);
            this.label1.TabIndex = 10;
            this.label1.Text = "Excel path";
            // 
            // btnSelectExcel
            // 
            this.btnSelectExcel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectExcel.BackColor = System.Drawing.Color.LightGreen;
            this.btnSelectExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnSelectExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectExcel.Location = new System.Drawing.Point(557, 13);
            this.btnSelectExcel.Name = "btnSelectExcel";
            this.btnSelectExcel.Size = new System.Drawing.Size(124, 30);
            this.btnSelectExcel.TabIndex = 9;
            this.btnSelectExcel.Text = "Select Excel";
            this.btnSelectExcel.UseVisualStyleBackColor = false;
            this.btnSelectExcel.Click += new System.EventHandler(this.btnSelectExcel_Click_1);
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.Enabled = false;
            this.txtExcelPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExcelPath.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.txtExcelPath.Location = new System.Drawing.Point(88, 14);
            this.txtExcelPath.Multiline = true;
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(445, 27);
            this.txtExcelPath.TabIndex = 8;
            this.txtExcelPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 27);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1038, 486);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Update Revision";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // timerUpdateInfor
            // 
            this.timerUpdateInfor.Tick += new System.EventHandler(this.timerUpdateInfor_Tick);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1042, 517);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "mainForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Update Title Block Tool - Simpson Strong Tie - CSS Team - 2024";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataDrawingGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btnUpdateInfor;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dataDrawingGrid;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbbTitleBlockName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelectExcel;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ProgressBar progressBarUpdateInFor;
        private System.Windows.Forms.Label lblLoading;
        private System.Windows.Forms.Timer timerUpdateInfor;
    }
}