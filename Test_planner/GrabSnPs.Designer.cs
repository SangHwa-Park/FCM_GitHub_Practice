
namespace S_para_planner
{
    partial class GrabSnPs
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
            this.BTN_selectPath = new System.Windows.Forms.Button();
            this.TBox_FilePath = new System.Windows.Forms.TextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Text_TestNum = new System.Windows.Forms.TextBox();
            this.Check_Data_erase = new System.Windows.Forms.CheckBox();
            this.BTN_Grab_SNP = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.Control;
            this.label1.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(422, 24);
            this.label1.TabIndex = 7;
            this.label1.Text = "SNP 파일 위치 (*.snp 파일이 있는 상위 폴더의 경로)";
            // 
            // BTN_selectPath
            // 
            this.BTN_selectPath.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BTN_selectPath.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_selectPath.Location = new System.Drawing.Point(647, 14);
            this.BTN_selectPath.Name = "BTN_selectPath";
            this.BTN_selectPath.Size = new System.Drawing.Size(127, 41);
            this.BTN_selectPath.TabIndex = 8;
            this.BTN_selectPath.Text = "폴더 선택";
            this.BTN_selectPath.UseVisualStyleBackColor = true;
            this.BTN_selectPath.Click += new System.EventHandler(this.BTN_selectPath_Click);
            // 
            // TBox_FilePath
            // 
            this.TBox_FilePath.Location = new System.Drawing.Point(16, 61);
            this.TBox_FilePath.Name = "TBox_FilePath";
            this.TBox_FilePath.Size = new System.Drawing.Size(758, 20);
            this.TBox_FilePath.TabIndex = 9;
            // 
            // listBox1
            // 
            this.listBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(16, 87);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(758, 249);
            this.listBox1.TabIndex = 10;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 374);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(174, 29);
            this.label2.TabIndex = 11;
            this.label2.Text = "테스트 번호 검색 :";
            // 
            // Text_TestNum
            // 
            this.Text_TestNum.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Text_TestNum.Location = new System.Drawing.Point(192, 374);
            this.Text_TestNum.MaximumSize = new System.Drawing.Size(350, 60);
            this.Text_TestNum.MinimumSize = new System.Drawing.Size(350, 50);
            this.Text_TestNum.Name = "Text_TestNum";
            this.Text_TestNum.Size = new System.Drawing.Size(350, 50);
            this.Text_TestNum.TabIndex = 12;
            // 
            // Check_Data_erase
            // 
            this.Check_Data_erase.AutoSize = true;
            this.Check_Data_erase.Checked = true;
            this.Check_Data_erase.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Check_Data_erase.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Check_Data_erase.Location = new System.Drawing.Point(192, 345);
            this.Check_Data_erase.Name = "Check_Data_erase";
            this.Check_Data_erase.Size = new System.Drawing.Size(242, 23);
            this.Check_Data_erase.TabIndex = 13;
            this.Check_Data_erase.Text = "매 검색시 기존 파일 삭제 (Optional)";
            this.Check_Data_erase.UseVisualStyleBackColor = true;
            // 
            // BTN_Grab_SNP
            // 
            this.BTN_Grab_SNP.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BTN_Grab_SNP.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_Grab_SNP.Location = new System.Drawing.Point(559, 374);
            this.BTN_Grab_SNP.Name = "BTN_Grab_SNP";
            this.BTN_Grab_SNP.Size = new System.Drawing.Size(215, 50);
            this.BTN_Grab_SNP.TabIndex = 14;
            this.BTN_Grab_SNP.Text = "검색 및 파일 모음";
            this.BTN_Grab_SNP.UseVisualStyleBackColor = true;
            this.BTN_Grab_SNP.Click += new System.EventHandler(this.BTN_Grab_SNP_Click);
            // 
            // GrabSnPs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.BTN_Grab_SNP);
            this.Controls.Add(this.Check_Data_erase);
            this.Controls.Add(this.Text_TestNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.TBox_FilePath);
            this.Controls.Add(this.BTN_selectPath);
            this.Controls.Add(this.label1);
            this.Name = "GrabSnPs";
            this.Text = "Grab_SNPs";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BTN_selectPath;
        private System.Windows.Forms.TextBox TBox_FilePath;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Text_TestNum;
        private System.Windows.Forms.CheckBox Check_Data_erase;
        private System.Windows.Forms.Button BTN_Grab_SNP;
    }
}