namespace UpdateAttribute
{
    partial class ProcessForm
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
            this.progressBarForm = new System.Windows.Forms.ProgressBar();
            this.timerFormProcess = new System.Windows.Forms.Timer(this.components);
            this.btnCloseProcess = new System.Windows.Forms.Button();
            this.lblLoadingStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBarForm
            // 
            this.progressBarForm.Location = new System.Drawing.Point(28, 46);
            this.progressBarForm.Maximum = 50;
            this.progressBarForm.Name = "progressBarForm";
            this.progressBarForm.Size = new System.Drawing.Size(441, 29);
            this.progressBarForm.Step = 5;
            this.progressBarForm.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBarForm.TabIndex = 0;
            // 
            // timerFormProcess
            // 
            this.timerFormProcess.Interval = 1000;
            this.timerFormProcess.Tick += new System.EventHandler(this.timerFormProcess_Tick);
            // 
            // btnCloseProcess
            // 
            this.btnCloseProcess.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnCloseProcess.Location = new System.Drawing.Point(189, 81);
            this.btnCloseProcess.Name = "btnCloseProcess";
            this.btnCloseProcess.Size = new System.Drawing.Size(118, 35);
            this.btnCloseProcess.TabIndex = 1;
            this.btnCloseProcess.Text = "OK";
            this.btnCloseProcess.UseVisualStyleBackColor = false;
            this.btnCloseProcess.Click += new System.EventHandler(this.btnCloseProcess_Click);
            // 
            // lblLoadingStatus
            // 
            this.lblLoadingStatus.AutoSize = true;
            this.lblLoadingStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoadingStatus.Location = new System.Drawing.Point(24, 14);
            this.lblLoadingStatus.Name = "lblLoadingStatus";
            this.lblLoadingStatus.Size = new System.Drawing.Size(88, 20);
            this.lblLoadingStatus.TabIndex = 2;
            this.lblLoadingStatus.Text = "Loading...";
            // 
            // ProcessForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(492, 128);
            this.Controls.Add(this.lblLoadingStatus);
            this.Controls.Add(this.btnCloseProcess);
            this.Controls.Add(this.progressBarForm);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProcessForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Loading ";
            this.Load += new System.EventHandler(this.ProcessForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBarForm;
        private System.Windows.Forms.Timer timerFormProcess;
        private System.Windows.Forms.Button btnCloseProcess;
        private System.Windows.Forms.Label lblLoadingStatus;
    }
}