namespace Test_Planner
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
            this.Spara_button = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fIleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadCMWithDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeWindowsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.projectsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.setupConfigToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.portConfigToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utillityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sparaPPTPlannerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadCMSheetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.specCompareToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Table_Make = new System.Windows.Forms.Button();
            this.BTN_Build_PPT = new System.Windows.Forms.Button();
            this.BTN_GrabSnp = new System.Windows.Forms.Button();
            this.BTN_Create_Worst = new System.Windows.Forms.Button();
            this.TxTest_button = new System.Windows.Forms.Button();
            this.BTN_INSERT_SPARA = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // Spara_button
            // 
            this.Spara_button.Enabled = false;
            this.Spara_button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Spara_button.Location = new System.Drawing.Point(643, 41);
            this.Spara_button.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Spara_button.Name = "Spara_button";
            this.Spara_button.Size = new System.Drawing.Size(188, 59);
            this.Spara_button.TabIndex = 0;
            this.Spara_button.Text = "Build Spara Plan";
            this.Spara_button.UseVisualStyleBackColor = true;
            this.Spara_button.Click += new System.EventHandler(this.button1_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fIleToolStripMenuItem,
            this.projectsToolStripMenuItem,
            this.settingsToolStripMenuItem,
            this.utillityToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1951, 28);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fIleToolStripMenuItem
            // 
            this.fIleToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadCMWithDataToolStripMenuItem,
            this.closeWindowsToolStripMenuItem});
            this.fIleToolStripMenuItem.Name = "fIleToolStripMenuItem";
            this.fIleToolStripMenuItem.Size = new System.Drawing.Size(46, 24);
            this.fIleToolStripMenuItem.Text = "File";
            // 
            // loadCMWithDataToolStripMenuItem
            // 
            this.loadCMWithDataToolStripMenuItem.Name = "loadCMWithDataToolStripMenuItem";
            this.loadCMWithDataToolStripMenuItem.Size = new System.Drawing.Size(225, 26);
            this.loadCMWithDataToolStripMenuItem.Text = "Load CM with Data";
            this.loadCMWithDataToolStripMenuItem.Click += new System.EventHandler(this.loadCMWithDataToolStripMenuItem_Click);
            // 
            // closeWindowsToolStripMenuItem
            // 
            this.closeWindowsToolStripMenuItem.Name = "closeWindowsToolStripMenuItem";
            this.closeWindowsToolStripMenuItem.Size = new System.Drawing.Size(225, 26);
            this.closeWindowsToolStripMenuItem.Text = "Close";
            this.closeWindowsToolStripMenuItem.Click += new System.EventHandler(this.closeWindowsToolStripMenuItem_Click);
            // 
            // projectsToolStripMenuItem
            // 
            this.projectsToolStripMenuItem.Name = "projectsToolStripMenuItem";
            this.projectsToolStripMenuItem.Size = new System.Drawing.Size(76, 24);
            this.projectsToolStripMenuItem.Text = "Projects";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.setupConfigToolStripMenuItem,
            this.portConfigToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(77, 24);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // setupConfigToolStripMenuItem
            // 
            this.setupConfigToolStripMenuItem.Name = "setupConfigToolStripMenuItem";
            this.setupConfigToolStripMenuItem.Size = new System.Drawing.Size(205, 26);
            this.setupConfigToolStripMenuItem.Text = "CM sheet config";
            this.setupConfigToolStripMenuItem.Click += new System.EventHandler(this.setupConfigToolStripMenuItem_Click);
            // 
            // portConfigToolStripMenuItem
            // 
            this.portConfigToolStripMenuItem.Name = "portConfigToolStripMenuItem";
            this.portConfigToolStripMenuItem.Size = new System.Drawing.Size(205, 26);
            this.portConfigToolStripMenuItem.Text = "Port Config";
            this.portConfigToolStripMenuItem.Click += new System.EventHandler(this.portConfigToolStripMenuItem_Click);
            // 
            // utillityToolStripMenuItem
            // 
            this.utillityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sparaPPTPlannerToolStripMenuItem});
            this.utillityToolStripMenuItem.Name = "utillityToolStripMenuItem";
            this.utillityToolStripMenuItem.Size = new System.Drawing.Size(67, 24);
            this.utillityToolStripMenuItem.Text = "Utillity";
            // 
            // sparaPPTPlannerToolStripMenuItem
            // 
            this.sparaPPTPlannerToolStripMenuItem.Name = "sparaPPTPlannerToolStripMenuItem";
            this.sparaPPTPlannerToolStripMenuItem.Size = new System.Drawing.Size(218, 26);
            this.sparaPPTPlannerToolStripMenuItem.Text = "Spara PPT Planner";
            this.sparaPPTPlannerToolStripMenuItem.Click += new System.EventHandler(this.sparaPPTPlannerToolStripMenuItem_Click);
            // 
            // loadCMSheetToolStripMenuItem
            // 
            this.loadCMSheetToolStripMenuItem.Name = "loadCMSheetToolStripMenuItem";
            this.loadCMSheetToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // specCompareToolStripMenuItem
            // 
            this.specCompareToolStripMenuItem.Name = "specCompareToolStripMenuItem";
            this.specCompareToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 318);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1924, 506);
            this.dataGridView1.TabIndex = 2;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 302);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1924, 12);
            this.progressBar1.TabIndex = 3;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 41);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.ShortcutsEnabled = false;
            this.textBox1.Size = new System.Drawing.Size(615, 254);
            this.textBox1.TabIndex = 4;
            // 
            // Table_Make
            // 
            this.Table_Make.Enabled = false;
            this.Table_Make.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Table_Make.Location = new System.Drawing.Point(864, 41);
            this.Table_Make.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Table_Make.Name = "Table_Make";
            this.Table_Make.Size = new System.Drawing.Size(188, 59);
            this.Table_Make.TabIndex = 6;
            this.Table_Make.Text = "Build Summary Report";
            this.Table_Make.UseVisualStyleBackColor = true;
            this.Table_Make.Click += new System.EventHandler(this.Table_Make_Click);
            // 
            // BTN_Build_PPT
            // 
            this.BTN_Build_PPT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_Build_PPT.Location = new System.Drawing.Point(864, 127);
            this.BTN_Build_PPT.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BTN_Build_PPT.Name = "BTN_Build_PPT";
            this.BTN_Build_PPT.Size = new System.Drawing.Size(188, 59);
            this.BTN_Build_PPT.TabIndex = 7;
            this.BTN_Build_PPT.Text = "Build PPT plan";
            this.BTN_Build_PPT.UseVisualStyleBackColor = true;
            this.BTN_Build_PPT.Click += new System.EventHandler(this.BTN_Build_PPT_Click);
            // 
            // BTN_GrabSnp
            // 
            this.BTN_GrabSnp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_GrabSnp.Location = new System.Drawing.Point(864, 209);
            this.BTN_GrabSnp.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BTN_GrabSnp.Name = "BTN_GrabSnp";
            this.BTN_GrabSnp.Size = new System.Drawing.Size(188, 59);
            this.BTN_GrabSnp.TabIndex = 8;
            this.BTN_GrabSnp.Text = "Grab SNPs";
            this.BTN_GrabSnp.UseVisualStyleBackColor = true;
            this.BTN_GrabSnp.Click += new System.EventHandler(this.BTN_GrabSnp_Click);
            // 
            // BTN_Create_Worst
            // 
            this.BTN_Create_Worst.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_Create_Worst.Location = new System.Drawing.Point(643, 209);
            this.BTN_Create_Worst.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BTN_Create_Worst.Name = "BTN_Create_Worst";
            this.BTN_Create_Worst.Size = new System.Drawing.Size(188, 59);
            this.BTN_Create_Worst.TabIndex = 9;
            this.BTN_Create_Worst.Text = "Create Worst";
            this.BTN_Create_Worst.UseVisualStyleBackColor = true;
            this.BTN_Create_Worst.Click += new System.EventHandler(this.BTN_Create_Worst_Click);
            // 
            // TxTest_button
            // 
            this.TxTest_button.Enabled = false;
            this.TxTest_button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.TxTest_button.Location = new System.Drawing.Point(1082, 41);
            this.TxTest_button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxTest_button.Name = "TxTest_button";
            this.TxTest_button.Size = new System.Drawing.Size(188, 59);
            this.TxTest_button.TabIndex = 10;
            this.TxTest_button.Text = "Build TX Plan";
            this.TxTest_button.UseVisualStyleBackColor = true;
            this.TxTest_button.Click += new System.EventHandler(this.TxTest_button_Click);
            // 
            // BTN_INSERT_SPARA
            // 
            this.BTN_INSERT_SPARA.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_INSERT_SPARA.Location = new System.Drawing.Point(643, 127);
            this.BTN_INSERT_SPARA.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.BTN_INSERT_SPARA.Name = "BTN_INSERT_SPARA";
            this.BTN_INSERT_SPARA.Size = new System.Drawing.Size(188, 59);
            this.BTN_INSERT_SPARA.TabIndex = 11;
            this.BTN_INSERT_SPARA.Text = "Insert Spara Data";
            this.BTN_INSERT_SPARA.UseVisualStyleBackColor = true;
            this.BTN_INSERT_SPARA.Click += new System.EventHandler(this.BTN_INSERT_SPARA_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1951, 837);
            this.Controls.Add(this.BTN_INSERT_SPARA);
            this.Controls.Add(this.TxTest_button);
            this.Controls.Add(this.BTN_Create_Worst);
            this.Controls.Add(this.BTN_GrabSnp);
            this.Controls.Add(this.BTN_Build_PPT);
            this.Controls.Add(this.Table_Make);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.Spara_button);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Gentle Breed Auto planner";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Spara_button;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fIleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem setupConfigToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadCMSheetToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem specCompareToolStripMenuItem;
        private System.Windows.Forms.Button Table_Make;
        private System.Windows.Forms.ToolStripMenuItem loadCMWithDataToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeWindowsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem portConfigToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem projectsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem utillityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sparaPPTPlannerToolStripMenuItem;
        private System.Windows.Forms.Button BTN_Build_PPT;
        private System.Windows.Forms.Button BTN_GrabSnp;
        private System.Windows.Forms.Button BTN_Create_Worst;
        private System.Windows.Forms.Button TxTest_button;
        private System.Windows.Forms.Button BTN_INSERT_SPARA;
    }
}

