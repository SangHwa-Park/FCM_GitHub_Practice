namespace FlexTestLib.MsgBox
{
    partial class MsgBoxForm
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
            this.PanMsg = new System.Windows.Forms.Panel();
            this.labMsg = new System.Windows.Forms.Label();
            this.panBtn = new System.Windows.Forms.Panel();
            this.btn2 = new System.Windows.Forms.Button();
            this.btn1 = new System.Windows.Forms.Button();
            this.btn0 = new System.Windows.Forms.Button();
            this.PanMsg.SuspendLayout();
            this.panBtn.SuspendLayout();
            this.SuspendLayout();
            // 
            // PanMsg
            // 
            this.PanMsg.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.PanMsg.AutoSize = true;
            this.PanMsg.Controls.Add(this.labMsg);
            this.PanMsg.Location = new System.Drawing.Point(16, 9);
            this.PanMsg.Margin = new System.Windows.Forms.Padding(4);
            this.PanMsg.Name = "PanMsg";
            this.PanMsg.Size = new System.Drawing.Size(200, 66);
            this.PanMsg.TabIndex = 0;
            // 
            // labMsg
            // 
            this.labMsg.AutoSize = true;
            this.labMsg.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labMsg.Location = new System.Drawing.Point(5, 6);
            this.labMsg.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labMsg.Name = "labMsg";
            this.labMsg.Size = new System.Drawing.Size(56, 23);
            this.labMsg.TabIndex = 0;
            this.labMsg.Text = "label1";
            // 
            // panBtn
            // 
            this.panBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.panBtn.Controls.Add(this.btn2);
            this.panBtn.Controls.Add(this.btn1);
            this.panBtn.Controls.Add(this.btn0);
            this.panBtn.Location = new System.Drawing.Point(16, 82);
            this.panBtn.Margin = new System.Windows.Forms.Padding(4);
            this.panBtn.Name = "panBtn";
            this.panBtn.Size = new System.Drawing.Size(293, 62);
            this.panBtn.TabIndex = 1;
            // 
            // btn2
            // 
            this.btn2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn2.Location = new System.Drawing.Point(197, 5);
            this.btn2.Margin = new System.Windows.Forms.Padding(4);
            this.btn2.Name = "btn2";
            this.btn2.Size = new System.Drawing.Size(87, 49);
            this.btn2.TabIndex = 2;
            this.btn2.Text = "Button 2";
            this.btn2.UseVisualStyleBackColor = true;
            this.btn2.Click += new System.EventHandler(this.btn_Click);
            // 
            // btn1
            // 
            this.btn1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn1.Location = new System.Drawing.Point(103, 5);
            this.btn1.Margin = new System.Windows.Forms.Padding(4);
            this.btn1.Name = "btn1";
            this.btn1.Size = new System.Drawing.Size(87, 49);
            this.btn1.TabIndex = 1;
            this.btn1.Text = "Button 1";
            this.btn1.UseVisualStyleBackColor = true;
            this.btn1.Click += new System.EventHandler(this.btn_Click);
            // 
            // btn0
            // 
            this.btn0.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn0.Location = new System.Drawing.Point(8, 5);
            this.btn0.Margin = new System.Windows.Forms.Padding(4);
            this.btn0.Name = "btn0";
            this.btn0.Size = new System.Drawing.Size(87, 49);
            this.btn0.TabIndex = 0;
            this.btn0.Text = "Button 0";
            this.btn0.UseVisualStyleBackColor = true;
            this.btn0.Click += new System.EventHandler(this.btn_Click);
            // 
            // MsgBoxForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(324, 153);
            this.Controls.Add(this.panBtn);
            this.Controls.Add(this.PanMsg);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MsgBoxForm";
            this.Text = "MsgBoxForm";
            this.TopMost = true;
            this.PanMsg.ResumeLayout(false);
            this.PanMsg.PerformLayout();
            this.panBtn.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel PanMsg;
        public System.Windows.Forms.Label labMsg;
        private System.Windows.Forms.Panel panBtn;
        private System.Windows.Forms.Button btn0;
        private System.Windows.Forms.Button btn1;
        private System.Windows.Forms.Button btn2;
    }
}