
namespace Test_Planner
{
    partial class TxPlan_builder
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
            this.BTN_LoadTxSet = new System.Windows.Forms.Button();
            this.Show_Path_TxSet = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // BTN_LoadTxSet
            // 
            this.BTN_LoadTxSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BTN_LoadTxSet.Location = new System.Drawing.Point(670, 12);
            this.BTN_LoadTxSet.Name = "BTN_LoadTxSet";
            this.BTN_LoadTxSet.Size = new System.Drawing.Size(118, 40);
            this.BTN_LoadTxSet.TabIndex = 3;
            this.BTN_LoadTxSet.Text = "Load Setting";
            this.BTN_LoadTxSet.UseVisualStyleBackColor = true;
            this.BTN_LoadTxSet.Click += new System.EventHandler(this.BTN_LoadTxSet_Click);
            // 
            // Show_Path_TxSet
            // 
            this.Show_Path_TxSet.Location = new System.Drawing.Point(12, 12);
            this.Show_Path_TxSet.MaximumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_TxSet.MinimumSize = new System.Drawing.Size(650, 40);
            this.Show_Path_TxSet.Multiline = true;
            this.Show_Path_TxSet.Name = "Show_Path_TxSet";
            this.Show_Path_TxSet.ReadOnly = true;
            this.Show_Path_TxSet.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.Show_Path_TxSet.Size = new System.Drawing.Size(650, 40);
            this.Show_Path_TxSet.TabIndex = 2;
            // 
            // TxPlan_builder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.BTN_LoadTxSet);
            this.Controls.Add(this.Show_Path_TxSet);
            this.Name = "TxPlan_builder";
            this.Text = "Tx Plan Builder";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BTN_LoadTxSet;
        private System.Windows.Forms.TextBox Show_Path_TxSet;
    }
}