
namespace S_para_planner
{
    partial class Build_PPT_wfm
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
            this.Show_Path_TCF = new System.Windows.Forms.TextBox();
            this.BTN_LoadTCF = new System.Windows.Forms.Button();
            this.Show_Path_Unit = new System.Windows.Forms.TextBox();
            this.BTN_LoadUnits = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.CBox_UnitPath = new System.Windows.Forms.CheckedListBox();
            this.CBox_Bands = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.BTN_Build_PPT_Plan = new System.Windows.Forms.Button();
            this.SelectClear = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // Show_Path_TCF
            // 
            this.Show_Path_TCF.Location = new System.Drawing.Point(12, 63);
            this.Show_Path_TCF.MaximumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_TCF.MinimumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_TCF.Multiline = true;
            this.Show_Path_TCF.Name = "Show_Path_TCF";
            this.Show_Path_TCF.ReadOnly = true;
            this.Show_Path_TCF.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.Show_Path_TCF.Size = new System.Drawing.Size(650, 40);
            this.Show_Path_TCF.TabIndex = 0;
            // 
            // BTN_LoadTCF
            // 
            this.BTN_LoadTCF.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_LoadTCF.Location = new System.Drawing.Point(670, 63);
            this.BTN_LoadTCF.Name = "BTN_LoadTCF";
            this.BTN_LoadTCF.Size = new System.Drawing.Size(118, 40);
            this.BTN_LoadTCF.TabIndex = 1;
            this.BTN_LoadTCF.Text = "Load Plan";
            this.BTN_LoadTCF.UseVisualStyleBackColor = true;
            this.BTN_LoadTCF.Click += new System.EventHandler(this.BTN_LoadTCF_Click);
            // 
            // Show_Path_Unit
            // 
            this.Show_Path_Unit.Location = new System.Drawing.Point(12, 109);
            this.Show_Path_Unit.MaximumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_Unit.MinimumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_Unit.Multiline = true;
            this.Show_Path_Unit.Name = "Show_Path_Unit";
            this.Show_Path_Unit.ReadOnly = true;
            this.Show_Path_Unit.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.Show_Path_Unit.Size = new System.Drawing.Size(650, 40);
            this.Show_Path_Unit.TabIndex = 2;
            // 
            // BTN_LoadUnits
            // 
            this.BTN_LoadUnits.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_LoadUnits.Location = new System.Drawing.Point(670, 109);
            this.BTN_LoadUnits.Name = "BTN_LoadUnits";
            this.BTN_LoadUnits.Size = new System.Drawing.Size(118, 40);
            this.BTN_LoadUnits.TabIndex = 3;
            this.BTN_LoadUnits.Text = "Set SNP Path";
            this.BTN_LoadUnits.UseVisualStyleBackColor = true;
            this.BTN_LoadUnits.Click += new System.EventHandler(this.BTN_LoadUnits_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.MaximumSize = new System.Drawing.Size(650, 30);
            this.label1.MinimumSize = new System.Drawing.Size(650, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(650, 30);
            this.label1.TabIndex = 4;
            this.label1.Text = "Please select S-paratest plan and SNP file paths (need to be matched exactly)";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // CBox_UnitPath
            // 
            this.CBox_UnitPath.CheckOnClick = true;
            this.CBox_UnitPath.FormattingEnabled = true;
            this.CBox_UnitPath.Location = new System.Drawing.Point(17, 221);
            this.CBox_UnitPath.Name = "CBox_UnitPath";
            this.CBox_UnitPath.Size = new System.Drawing.Size(316, 169);
            this.CBox_UnitPath.TabIndex = 5;
            // 
            // CBox_Bands
            // 
            this.CBox_Bands.CheckOnClick = true;
            this.CBox_Bands.FormattingEnabled = true;
            this.CBox_Bands.Location = new System.Drawing.Point(346, 221);
            this.CBox_Bands.Name = "CBox_Bands";
            this.CBox_Bands.Size = new System.Drawing.Size(316, 169);
            this.CBox_Bands.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(14, 179);
            this.label2.MaximumSize = new System.Drawing.Size(300, 30);
            this.label2.MinimumSize = new System.Drawing.Size(300, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(300, 30);
            this.label2.TabIndex = 7;
            this.label2.Text = "DUT *.Snp selection";
            this.label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(343, 179);
            this.label3.MaximumSize = new System.Drawing.Size(300, 30);
            this.label3.MinimumSize = new System.Drawing.Size(300, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(300, 30);
            this.label3.TabIndex = 8;
            this.label3.Text = "Band selection";
            this.label3.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // BTN_Build_PPT_Plan
            // 
            this.BTN_Build_PPT_Plan.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_Build_PPT_Plan.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_Build_PPT_Plan.Location = new System.Drawing.Point(668, 221);
            this.BTN_Build_PPT_Plan.Name = "BTN_Build_PPT_Plan";
            this.BTN_Build_PPT_Plan.Size = new System.Drawing.Size(118, 169);
            this.BTN_Build_PPT_Plan.TabIndex = 9;
            this.BTN_Build_PPT_Plan.Text = "Generate PPT plan";
            this.BTN_Build_PPT_Plan.UseVisualStyleBackColor = true;
            this.BTN_Build_PPT_Plan.Click += new System.EventHandler(this.BTN_Build_PPT_Plan_Click);
            // 
            // SelectClear
            // 
            this.SelectClear.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.SelectClear.Location = new System.Drawing.Point(532, 186);
            this.SelectClear.Name = "SelectClear";
            this.SelectClear.Size = new System.Drawing.Size(130, 29);
            this.SelectClear.TabIndex = 10;
            this.SelectClear.Text = "Select / Clear all";
            this.SelectClear.UseVisualStyleBackColor = true;
            this.SelectClear.Click += new System.EventHandler(this.SelectClear_Click);
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Location = new System.Drawing.Point(203, 186);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(130, 29);
            this.button1.TabIndex = 11;
            this.button1.Text = "Select / Clear all";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 160);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(776, 16);
            this.progressBar1.TabIndex = 12;
            // 
            // Build_PPT_wfm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.SelectClear);
            this.Controls.Add(this.BTN_Build_PPT_Plan);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CBox_Bands);
            this.Controls.Add(this.CBox_UnitPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BTN_LoadUnits);
            this.Controls.Add(this.Show_Path_Unit);
            this.Controls.Add(this.BTN_LoadTCF);
            this.Controls.Add(this.Show_Path_TCF);
            this.Name = "Build_PPT_wfm";
            this.Text = "Build PPT plan from S-para test plan (TCF)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Show_Path_TCF;
        private System.Windows.Forms.Button BTN_LoadTCF;
        private System.Windows.Forms.TextBox Show_Path_Unit;
        private System.Windows.Forms.Button BTN_LoadUnits;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox CBox_UnitPath;
        private System.Windows.Forms.CheckedListBox CBox_Bands;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BTN_Build_PPT_Plan;
        private System.Windows.Forms.Button SelectClear;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}