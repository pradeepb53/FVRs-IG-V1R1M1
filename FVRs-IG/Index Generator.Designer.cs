namespace FVRs_IG
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.labelWordList = new System.Windows.Forms.Label();
            this.labelSelectFile = new System.Windows.Forms.Label();
            this.textBoxSelectFile = new System.Windows.Forms.TextBox();
            this.buttonCreate = new System.Windows.Forms.Button();
            this.buttonClear = new System.Windows.Forms.Button();
            this.openFileDialogSelectFile = new System.Windows.Forms.OpenFileDialog();
            this.progressBarCoreOps = new System.Windows.Forms.ProgressBar();
            this.timerCoreOps = new System.Windows.Forms.Timer(this.components);
            this.textBoxExcludedWords = new System.Windows.Forms.TextBox();
            this.buttonSaveWords = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelWordList
            // 
            this.labelWordList.AutoSize = true;
            this.labelWordList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelWordList.Location = new System.Drawing.Point(18, 19);
            this.labelWordList.Name = "labelWordList";
            this.labelWordList.Size = new System.Drawing.Size(110, 15);
            this.labelWordList.TabIndex = 1;
            this.labelWordList.Text = "Excluded Words";
            // 
            // labelSelectFile
            // 
            this.labelSelectFile.AutoSize = true;
            this.labelSelectFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSelectFile.Location = new System.Drawing.Point(239, 144);
            this.labelSelectFile.Name = "labelSelectFile";
            this.labelSelectFile.Size = new System.Drawing.Size(131, 15);
            this.labelSelectFile.TabIndex = 3;
            this.labelSelectFile.Text = "Select a Transcription :";
            // 
            // textBoxSelectFile
            // 
            this.textBoxSelectFile.Location = new System.Drawing.Point(386, 143);
            this.textBoxSelectFile.Name = "textBoxSelectFile";
            this.textBoxSelectFile.Size = new System.Drawing.Size(268, 20);
            this.textBoxSelectFile.TabIndex = 4;
            this.textBoxSelectFile.Click += new System.EventHandler(this.textBoxSelectFile_Click);
            // 
            // buttonCreate
            // 
            this.buttonCreate.BackColor = System.Drawing.SystemColors.Control;
            this.buttonCreate.Location = new System.Drawing.Point(386, 396);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(75, 23);
            this.buttonCreate.TabIndex = 5;
            this.buttonCreate.Text = "Create Index";
            this.buttonCreate.UseVisualStyleBackColor = false;
            this.buttonCreate.Click += new System.EventHandler(this.buttonCreate_Click);
            // 
            // buttonClear
            // 
            this.buttonClear.Location = new System.Drawing.Point(523, 395);
            this.buttonClear.Name = "buttonClear";
            this.buttonClear.Size = new System.Drawing.Size(75, 23);
            this.buttonClear.TabIndex = 6;
            this.buttonClear.Text = "Clear File";
            this.buttonClear.UseVisualStyleBackColor = true;
            this.buttonClear.Click += new System.EventHandler(this.buttonClear_Click);
            // 
            // openFileDialogSelectFile
            // 
            this.openFileDialogSelectFile.FileName = "openFileDialogSelectFile";
            // 
            // progressBarCoreOps
            // 
            this.progressBarCoreOps.Location = new System.Drawing.Point(203, 273);
            this.progressBarCoreOps.Name = "progressBarCoreOps";
            this.progressBarCoreOps.Size = new System.Drawing.Size(451, 23);
            this.progressBarCoreOps.TabIndex = 7;
            // 
            // timerCoreOps
            // 
            this.timerCoreOps.Tick += new System.EventHandler(this.timerCoreOps_Tick);
            // 
            // textBoxExcludedWords
            // 
            this.textBoxExcludedWords.Location = new System.Drawing.Point(21, 53);
            this.textBoxExcludedWords.Multiline = true;
            this.textBoxExcludedWords.Name = "textBoxExcludedWords";
            this.textBoxExcludedWords.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxExcludedWords.Size = new System.Drawing.Size(100, 337);
            this.textBoxExcludedWords.TabIndex = 8;
            this.textBoxExcludedWords.TextChanged += new System.EventHandler(this.textBoxExcludedWords_TextChanged);
            // 
            // buttonSaveWords
            // 
            this.buttonSaveWords.Enabled = false;
            this.buttonSaveWords.Location = new System.Drawing.Point(33, 396);
            this.buttonSaveWords.Name = "buttonSaveWords";
            this.buttonSaveWords.Size = new System.Drawing.Size(75, 23);
            this.buttonSaveWords.TabIndex = 9;
            this.buttonSaveWords.Text = "Save-Words";
            this.buttonSaveWords.UseVisualStyleBackColor = true;
            this.buttonSaveWords.Click += new System.EventHandler(this.buttonSaveWords_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 445);
            this.Controls.Add(this.buttonSaveWords);
            this.Controls.Add(this.textBoxExcludedWords);
            this.Controls.Add(this.progressBarCoreOps);
            this.Controls.Add(this.buttonClear);
            this.Controls.Add(this.buttonCreate);
            this.Controls.Add(this.textBoxSelectFile);
            this.Controls.Add(this.labelSelectFile);
            this.Controls.Add(this.labelWordList);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "FVRs Index Generator - V1R1M0                                                    " +
    "                    Copyright(2017-FVRs)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label labelWordList;
        private System.Windows.Forms.Label labelSelectFile;
        private System.Windows.Forms.TextBox textBoxSelectFile;
        private System.Windows.Forms.Button buttonCreate;
        private System.Windows.Forms.Button buttonClear;
        private System.Windows.Forms.OpenFileDialog openFileDialogSelectFile;
        private System.Windows.Forms.ProgressBar progressBarCoreOps;
        private System.Windows.Forms.Timer timerCoreOps;
        private System.Windows.Forms.TextBox textBoxExcludedWords;
        private System.Windows.Forms.Button buttonSaveWords;
    }
}

