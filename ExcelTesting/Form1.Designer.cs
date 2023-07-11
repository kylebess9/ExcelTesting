namespace ExcelTesting
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.homeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.howToToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSharePoint = new System.Windows.Forms.Button();
            this.btnLoadFile2 = new System.Windows.Forms.Button();
            this.btnLoadFile1 = new System.Windows.Forms.Button();
            this.btnSubmitFile1 = new System.Windows.Forms.Button();
            this.btnSubmitFile2 = new System.Windows.Forms.Button();
            this.lbxOutputFile1 = new System.Windows.Forms.ListBox();
            this.lbxOutputFile2 = new System.Windows.Forms.ListBox();
            this.openFile = new System.Windows.Forms.OpenFileDialog();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.homeToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1308, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // homeToolStripMenuItem
            // 
            this.homeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.printToolStripMenuItem});
            this.homeToolStripMenuItem.Name = "homeToolStripMenuItem";
            this.homeToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.homeToolStripMenuItem.Text = "File";
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.saveToolStripMenuItem.Text = "Save";
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.saveAsToolStripMenuItem.Text = "Save As";
            // 
            // printToolStripMenuItem
            // 
            this.printToolStripMenuItem.Name = "printToolStripMenuItem";
            this.printToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.printToolStripMenuItem.Text = "Print";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.howToToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // howToToolStripMenuItem
            // 
            this.howToToolStripMenuItem.Name = "howToToolStripMenuItem";
            this.howToToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.howToToolStripMenuItem.Text = "How to";
            // 
            // btnSharePoint
            // 
            this.btnSharePoint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(154)))), ((int)(((byte)(5)))));
            this.btnSharePoint.FlatAppearance.BorderSize = 0;
            this.btnSharePoint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSharePoint.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnSharePoint.ForeColor = System.Drawing.Color.White;
            this.btnSharePoint.Location = new System.Drawing.Point(549, 27);
            this.btnSharePoint.Name = "btnSharePoint";
            this.btnSharePoint.Size = new System.Drawing.Size(149, 44);
            this.btnSharePoint.TabIndex = 1;
            this.btnSharePoint.Text = "Share Point";
            this.btnSharePoint.UseVisualStyleBackColor = false;
            this.btnSharePoint.Click += new System.EventHandler(this.btnSharePoint_Click);
            // 
            // btnLoadFile2
            // 
            this.btnLoadFile2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(154)))), ((int)(((byte)(5)))));
            this.btnLoadFile2.FlatAppearance.BorderSize = 0;
            this.btnLoadFile2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLoadFile2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnLoadFile2.ForeColor = System.Drawing.Color.White;
            this.btnLoadFile2.Location = new System.Drawing.Point(899, 90);
            this.btnLoadFile2.Name = "btnLoadFile2";
            this.btnLoadFile2.Size = new System.Drawing.Size(149, 44);
            this.btnLoadFile2.TabIndex = 2;
            this.btnLoadFile2.Text = "File 2 ";
            this.btnLoadFile2.UseVisualStyleBackColor = false;
            // 
            // btnLoadFile1
            // 
            this.btnLoadFile1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(154)))), ((int)(((byte)(5)))));
            this.btnLoadFile1.FlatAppearance.BorderSize = 0;
            this.btnLoadFile1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLoadFile1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnLoadFile1.ForeColor = System.Drawing.Color.White;
            this.btnLoadFile1.Location = new System.Drawing.Point(231, 90);
            this.btnLoadFile1.Name = "btnLoadFile1";
            this.btnLoadFile1.Size = new System.Drawing.Size(149, 44);
            this.btnLoadFile1.TabIndex = 3;
            this.btnLoadFile1.Text = "File 1";
            this.btnLoadFile1.UseVisualStyleBackColor = false;
            // 
            // btnSubmitFile1
            // 
            this.btnSubmitFile1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(154)))), ((int)(((byte)(5)))));
            this.btnSubmitFile1.FlatAppearance.BorderSize = 0;
            this.btnSubmitFile1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmitFile1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnSubmitFile1.ForeColor = System.Drawing.Color.White;
            this.btnSubmitFile1.Location = new System.Drawing.Point(12, 867);
            this.btnSubmitFile1.Name = "btnSubmitFile1";
            this.btnSubmitFile1.Size = new System.Drawing.Size(149, 43);
            this.btnSubmitFile1.TabIndex = 4;
            this.btnSubmitFile1.Text = "Submit";
            this.btnSubmitFile1.UseVisualStyleBackColor = false;
            // 
            // btnSubmitFile2
            // 
            this.btnSubmitFile2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(154)))), ((int)(((byte)(5)))));
            this.btnSubmitFile2.FlatAppearance.BorderSize = 0;
            this.btnSubmitFile2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmitFile2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnSubmitFile2.ForeColor = System.Drawing.Color.White;
            this.btnSubmitFile2.Location = new System.Drawing.Point(651, 867);
            this.btnSubmitFile2.Name = "btnSubmitFile2";
            this.btnSubmitFile2.Size = new System.Drawing.Size(149, 43);
            this.btnSubmitFile2.TabIndex = 5;
            this.btnSubmitFile2.Text = "Submit";
            this.btnSubmitFile2.UseVisualStyleBackColor = false;
            // 
            // lbxOutputFile1
            // 
            this.lbxOutputFile1.FormattingEnabled = true;
            this.lbxOutputFile1.ItemHeight = 15;
            this.lbxOutputFile1.Location = new System.Drawing.Point(12, 152);
            this.lbxOutputFile1.Name = "lbxOutputFile1";
            this.lbxOutputFile1.Size = new System.Drawing.Size(634, 709);
            this.lbxOutputFile1.TabIndex = 6;
            // 
            // lbxOutputFile2
            // 
            this.lbxOutputFile2.FormattingEnabled = true;
            this.lbxOutputFile2.ItemHeight = 15;
            this.lbxOutputFile2.Location = new System.Drawing.Point(651, 152);
            this.lbxOutputFile2.Name = "lbxOutputFile2";
            this.lbxOutputFile2.Size = new System.Drawing.Size(645, 709);
            this.lbxOutputFile2.TabIndex = 7;
            // 
            // openFile
            // 
            this.openFile.FileName = "openFile";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1308, 931);
            this.Controls.Add(this.lbxOutputFile2);
            this.Controls.Add(this.lbxOutputFile1);
            this.Controls.Add(this.btnSubmitFile2);
            this.Controls.Add(this.btnSubmitFile1);
            this.Controls.Add(this.btnLoadFile1);
            this.Controls.Add(this.btnLoadFile2);
            this.Controls.Add(this.btnSharePoint);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Excel Compare";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem homeToolStripMenuItem;
        private ToolStripMenuItem saveToolStripMenuItem;
        private ToolStripMenuItem saveAsToolStripMenuItem;
        private ToolStripMenuItem printToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem howToToolStripMenuItem;
        private Button btnSharePoint;
        private Button btnLoadFile2;
        private Button btnLoadFile1;
        private Button btnSubmitFile1;
        private Button btnSubmitFile2;
        private ListBox lbxOutputFile1;
        private ListBox lbxOutputFile2;
        private OpenFileDialog openFile;
    }
}