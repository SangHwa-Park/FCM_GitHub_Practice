using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using S_para_planner;
using Excel_Base;

namespace S_para_planner
{
    public partial class Popup_ProgressBar : Form
    {
        Excel_File this_excel;

        public Popup_ProgressBar(int count)
        {
            InitializeComponent();
            Init(count);
            this.button1.Enabled = false;
        }

        public void Init(int count)
        {
            this.progressBar1.Style = ProgressBarStyle.Continuous;
            this.progressBar1.Minimum = 0;
            this.progressBar1.Maximum = count;
            this.progressBar1.Step = (int)((progressBar1.Maximum - progressBar1.Minimum) / count);
            this.progressBar1.Value = 0;
            this.progressBar1.MarqueeAnimationSpeed = 1;
        }

        public bool Done(Excel_File excel_File)
        {
            this.button1.Enabled = true;
            this.this_excel = excel_File;
            return true;
        }

        public void execute_step(string text1, string text2)
        {
            this.label1.Text = text1 + " %\r\n";
            this.label2.Text = text2 + "\r\n";
            this.progressBar1.PerformStep();
        }
        public void execute_step_msgOnly(string text2)
        {
            //this.label1.Text = text1 + " %\r\n";
            this.label2.Text = text2 + "\r\n";
            this.progressBar1.PerformStep();
        }

        public void ShowMSG(string text)
        {
            this.label2.Text = text + "\r\n";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.this_excel.App.ScreenUpdating = true;
            this.this_excel.Worksheet.Activate();
            this.this_excel.App.Visible = true;
            Close();
        }
    }
}
