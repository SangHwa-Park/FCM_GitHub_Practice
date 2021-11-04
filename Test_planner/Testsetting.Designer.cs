using System;
using System.Windows.Forms;

namespace Test_Planner
{
    partial class Testsetting
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
            this.Extream_Condition = new System.Windows.Forms.CheckedListBox();
            this.Test_Port_Input = new System.Windows.Forms.CheckedListBox();
            this.Test_Field = new System.Windows.Forms.CheckedListBox();
            this.Test_Port_Output = new System.Windows.Forms.CheckedListBox();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.listView1 = new System.Windows.Forms.ListView();
            this.Generate_Plan = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Extream_Condition
            // 
            this.Extream_Condition.CheckOnClick = true;
            this.Extream_Condition.FormattingEnabled = true;
            this.Extream_Condition.Location = new System.Drawing.Point(771, 429);
            this.Extream_Condition.Name = "Extream_Condition";
            this.Extream_Condition.Size = new System.Drawing.Size(151, 229);
            this.Extream_Condition.TabIndex = 2;
            this.Extream_Condition.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.Temp_Field_Select);
            // 
            // Test_Port_Input
            // 
            this.Test_Port_Input.CheckOnClick = true;
            this.Test_Port_Input.FormattingEnabled = true;
            this.Test_Port_Input.Location = new System.Drawing.Point(928, 429);
            this.Test_Port_Input.Name = "Test_Port_Input";
            this.Test_Port_Input.Size = new System.Drawing.Size(151, 229);
            this.Test_Port_Input.TabIndex = 3;
            // 
            // Test_Field
            // 
            this.Test_Field.CheckOnClick = true;
            this.Test_Field.FormattingEnabled = true;
            this.Test_Field.Location = new System.Drawing.Point(613, 429);
            this.Test_Field.Name = "Test_Field";
            this.Test_Field.Size = new System.Drawing.Size(152, 229);
            this.Test_Field.TabIndex = 5;
            this.Test_Field.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.Test_Field_Select);
            // 
            // Test_Port_Output
            // 
            this.Test_Port_Output.CheckOnClick = true;
            this.Test_Port_Output.FormattingEnabled = true;
            this.Test_Port_Output.Location = new System.Drawing.Point(1085, 429);
            this.Test_Port_Output.Name = "Test_Port_Output";
            this.Test_Port_Output.Size = new System.Drawing.Size(152, 229);
            this.Test_Port_Output.TabIndex = 6;
            // 
            // treeView1
            // 
            this.treeView1.CheckBoxes = true;
            this.treeView1.Location = new System.Drawing.Point(12, 12);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(595, 646);
            this.treeView1.TabIndex = 7;
            this.treeView1.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.Tree_Nodes_Checked);
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.Tree_Nodes_Selected);
            // 
            // listView1
            // 
            this.listView1.CheckBoxes = true;
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(613, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(871, 384);
            this.listView1.TabIndex = 8;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.Set_TreeNodes_from_ListView);
            this.listView1.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.Refresh_ListView);
            // 
            // Generate_Plan
            // 
            this.Generate_Plan.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.Generate_Plan.FlatAppearance.BorderSize = 5;
            this.Generate_Plan.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Generate_Plan.Location = new System.Drawing.Point(1260, 429);
            this.Generate_Plan.Name = "Generate_Plan";
            this.Generate_Plan.Size = new System.Drawing.Size(224, 58);
            this.Generate_Plan.TabIndex = 9;
            this.Generate_Plan.Text = "Test Plan Generate";
            this.Generate_Plan.UseVisualStyleBackColor = false;
            this.Generate_Plan.Click += new System.EventHandler(this.Generate_Plan_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(642, 410);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 16);
            this.label1.TabIndex = 10;
            this.label1.Text = "TEST ITEM";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(813, 410);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 16);
            this.label2.TabIndex = 11;
            this.label2.Text = "TEMP(C)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(957, 410);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 16);
            this.label3.TabIndex = 12;
            this.label3.Text = "Input PORT";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(1106, 410);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 16);
            this.label4.TabIndex = 13;
            this.label4.Text = "Output PORT";
            // 
            // Testsetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1498, 697);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Generate_Plan);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.Test_Port_Output);
            this.Controls.Add(this.Test_Field);
            this.Controls.Add(this.Test_Port_Input);
            this.Controls.Add(this.Extream_Condition);
            this.Name = "Testsetting";
            this.Text = "S-parameter Test Planner";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckedListBox Extream_Condition;
        private System.Windows.Forms.CheckedListBox Test_Port_Input;
        private System.Windows.Forms.CheckedListBox Test_Field;
        private System.Windows.Forms.CheckedListBox Test_Port_Output;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ListView listView1;
        private Button Generate_Plan;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
    }
}