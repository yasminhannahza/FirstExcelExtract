namespace FirstExcelExtract
{
    partial class frmFirstExcel
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
            this.txtConsole = new System.Windows.Forms.TextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.pbProgress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // txtConsole
            // 
            this.txtConsole.Location = new System.Drawing.Point(12, 219);
            this.txtConsole.Multiline = true;
            this.txtConsole.Name = "txtConsole";
            this.txtConsole.ReadOnly = true;
            this.txtConsole.Size = new System.Drawing.Size(882, 271);
            this.txtConsole.TabIndex = 0;
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(478, 78);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(147, 60);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Jeng Jeng";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // pbProgress
            // 
            this.pbProgress.Location = new System.Drawing.Point(12, 511);
            this.pbProgress.Name = "pbProgress";
            this.pbProgress.Size = new System.Drawing.Size(882, 33);
            this.pbProgress.TabIndex = 2;
            this.pbProgress.Visible = false;
            // 
            // frmFirstExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(906, 556);
            this.Controls.Add(this.pbProgress);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtConsole);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "frmFirstExcel";
            this.Text = "My First Excel Manipulation Program";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtConsole;
        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.ProgressBar pbProgress;
    }
}

