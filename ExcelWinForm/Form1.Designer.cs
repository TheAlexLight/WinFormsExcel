namespace ExcelWinForm
{
    partial class Form1
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
            this.button_Run = new System.Windows.Forms.Button();
            this.btnOpenExcFile = new System.Windows.Forms.Button();
            this.txtCurrPath = new System.Windows.Forms.TextBox();
            this.txtGroupCount = new System.Windows.Forms.TextBox();
            this.labelExcPath = new System.Windows.Forms.Label();
            this.labelGroupCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button_Run
            // 
            this.button_Run.Location = new System.Drawing.Point(489, 337);
            this.button_Run.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.button_Run.Name = "button_Run";
            this.button_Run.Size = new System.Drawing.Size(266, 132);
            this.button_Run.TabIndex = 0;
            this.button_Run.Text = "Run";
            this.button_Run.UseVisualStyleBackColor = true;
            this.button_Run.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnOpenExcFile
            // 
            this.btnOpenExcFile.Location = new System.Drawing.Point(106, 66);
            this.btnOpenExcFile.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.btnOpenExcFile.Name = "btnOpenExcFile";
            this.btnOpenExcFile.Size = new System.Drawing.Size(187, 83);
            this.btnOpenExcFile.TabIndex = 1;
            this.btnOpenExcFile.Text = "Excel file path";
            this.btnOpenExcFile.UseVisualStyleBackColor = true;
            this.btnOpenExcFile.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtCurrPath
            // 
            this.txtCurrPath.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtCurrPath.Location = new System.Drawing.Point(314, 93);
            this.txtCurrPath.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.txtCurrPath.Name = "txtCurrPath";
            this.txtCurrPath.ReadOnly = true;
            this.txtCurrPath.Size = new System.Drawing.Size(727, 25);
            this.txtCurrPath.TabIndex = 2;
            // 
            // txtGroupCount
            // 
            this.txtGroupCount.Location = new System.Drawing.Point(513, 229);
            this.txtGroupCount.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.txtGroupCount.Name = "txtGroupCount";
            this.txtGroupCount.Size = new System.Drawing.Size(242, 29);
            this.txtGroupCount.TabIndex = 3;
            this.txtGroupCount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGroupCount_KeyPress);
            // 
            // labelExcPath
            // 
            this.labelExcPath.AutoSize = true;
            this.labelExcPath.Location = new System.Drawing.Point(554, 51);
            this.labelExcPath.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelExcPath.Name = "labelExcPath";
            this.labelExcPath.Size = new System.Drawing.Size(166, 22);
            this.labelExcPath.TabIndex = 4;
            this.labelExcPath.Text = "Current Excel path";
            // 
            // labelGroupCount
            // 
            this.labelGroupCount.AutoSize = true;
            this.labelGroupCount.Location = new System.Drawing.Point(553, 186);
            this.labelGroupCount.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelGroupCount.Name = "labelGroupCount";
            this.labelGroupCount.Size = new System.Drawing.Size(159, 22);
            this.labelGroupCount.TabIndex = 5;
            this.labelGroupCount.Text = "Enter group count";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 22F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 634);
            this.Controls.Add(this.labelGroupCount);
            this.Controls.Add(this.labelExcPath);
            this.Controls.Add(this.txtGroupCount);
            this.Controls.Add(this.txtCurrPath);
            this.Controls.Add(this.btnOpenExcFile);
            this.Controls.Add(this.button_Run);
            this.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_Run;
        private System.Windows.Forms.Button btnOpenExcFile;
        private System.Windows.Forms.TextBox txtCurrPath;
        private System.Windows.Forms.TextBox txtGroupCount;
        private System.Windows.Forms.Label labelExcPath;
        private System.Windows.Forms.Label labelGroupCount;
    }
}

