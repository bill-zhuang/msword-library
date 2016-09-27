namespace wordLibrary
{
    partial class formWordLibrary
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
            this.label1 = new System.Windows.Forms.Label();
            this.textFolderPath = new System.Windows.Forms.TextBox();
            this.textBoxProcess = new System.Windows.Forms.TextBox();
            this.progressBarConvert = new System.Windows.Forms.ProgressBar();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 81);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 12);
            this.label1.TabIndex = 15;
            this.label1.Text = "已选择目录/文件：";
            // 
            // textFolderPath
            // 
            this.textFolderPath.Enabled = false;
            this.textFolderPath.Location = new System.Drawing.Point(135, 78);
            this.textFolderPath.Name = "textFolderPath";
            this.textFolderPath.Size = new System.Drawing.Size(384, 21);
            this.textFolderPath.TabIndex = 14;
            // 
            // textBoxProcess
            // 
            this.textBoxProcess.Location = new System.Drawing.Point(24, 145);
            this.textBoxProcess.Multiline = true;
            this.textBoxProcess.Name = "textBoxProcess";
            this.textBoxProcess.Size = new System.Drawing.Size(495, 170);
            this.textBoxProcess.TabIndex = 13;
            // 
            // progressBarConvert
            // 
            this.progressBarConvert.Location = new System.Drawing.Point(24, 116);
            this.progressBarConvert.Name = "progressBarConvert";
            this.progressBarConvert.Size = new System.Drawing.Size(495, 23);
            this.progressBarConvert.TabIndex = 12;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(361, 12);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(86, 47);
            this.btnRun.TabIndex = 11;
            this.btnRun.Text = "运行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.Location = new System.Drawing.Point(24, 12);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(145, 47);
            this.btnOpenFolder.TabIndex = 10;
            this.btnOpenFolder.Text = "打开文件夹";
            this.btnOpenFolder.UseVisualStyleBackColor = true;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(188, 12);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(145, 47);
            this.btnSelectFile.TabIndex = 16;
            this.btnSelectFile.Text = "选择文件";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // formWordLibrary
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(541, 327);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textFolderPath);
            this.Controls.Add(this.textBoxProcess);
            this.Controls.Add(this.progressBarConvert);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnOpenFolder);
            this.MaximizeBox = false;
            this.Name = "formWordLibrary";
            this.Text = "Word Library";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textFolderPath;
        private System.Windows.Forms.TextBox textBoxProcess;
        private System.Windows.Forms.ProgressBar progressBarConvert;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnOpenFolder;
        private System.Windows.Forms.Button btnSelectFile;
    }
}

