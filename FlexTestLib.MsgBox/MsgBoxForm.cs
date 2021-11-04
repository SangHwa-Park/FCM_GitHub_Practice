using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace FlexTestLib.MsgBox
{
    public partial class MsgBoxForm : Form
    {
        public static string MsgBoxRtn = "InitialState";

        public bool btn1_Enable;
        public bool btn2_Enable;

        public MsgBoxForm(string BtnText)
        {
            this.Setup(BtnText);
        }

        public MsgBoxForm()
        {
            this.Setup("OK");
        }

        private void Setup(string BtnText)
        {
            InitializeComponent();
            string[] BtnTextAry = BtnText.Split('|');
            this.btn0.Text = (BtnText == "" ? "OK" : BtnTextAry[0]);
            if (this.btn0.Text == "OK") this.btn0.Font = new Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.btn1_Enable = (BtnTextAry.Length >= 2);
            this.btn2_Enable = (BtnTextAry.Length >= 3);

            // Configure Btn1
            this.btn1.Visible = this.btn1_Enable;
            this.btn1.Text = (this.btn1_Enable ? BtnTextAry[1] : "--");

            // Configure Btn2
            this.btn2.Visible = (this.btn2_Enable);
            this.btn2.Text = (this.btn2_Enable ? BtnTextAry[2] : "--");
        }

        public string ShowForm(string Prompt, string Msg, System.Drawing.Font MsgFont)
        {
            this.ShowForm(Prompt, Msg, -1, -1, MsgFont);
            return MsgBoxRtn;
        }

        public string ShowForm(string Prompt, string Msg)
        {
            this.ShowForm(Prompt, Msg, -1, -1, null);
            return MsgBoxRtn;
        }

        public string ShowForm(string Prompt, string Msg, int PsnX, int PsnY, System.Drawing.Font MsgFont)
        {
            MsgBoxRtn = "";
            if (MsgFont != null) this.labMsg.Font = MsgFont;
            this.Text = Prompt;
            this.labMsg.Text = Msg + "\n ";
            //int panBtnWidMin = (this.btn2_Enable ? 228 : 156);
            int panBtnWidMin = (this.btn2_Enable ? 293 : (this.btn1_Enable ? 228 : 156));
            this.panBtn.Width = HelperMethods.ClipLo((int)(Prompt.Length * 10), panBtnWidMin);  // This is an attempt to widen the MsgBox to accommodate the Caption Text

            // Set Position, then Show Form
            if (PsnX >= 0 && PsnY >= 0)
            {
                this.StartPosition = FormStartPosition.Manual;
                this.Location = new System.Drawing.Point(PsnX, PsnY);
            }
            else
            {
                this.StartPosition = FormStartPosition.CenterParent;
            }

            this.ShowDialog();

            // Update Results of User Response
            return MsgBoxRtn;
        }

        private void FormAlignCenter()
        {
            //ClsDisplay Disp = (ClsCache.GetObj("Cache_Display", null) as ClsDisplay);
            //if (Disp == null) return;
            //int x = Disp.WorkingAreaX - this.Width;
            //int y = Disp.WorkingAreaY - this.Height;

            int WorkingAreaX = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            int WorkingAreaY = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height;

            int x = WorkingAreaX - this.Width;
            int y = WorkingAreaY - this.Height;
            this.Location = new Point(x / 2, y / 2);
        }

        private void btn_Click(object sender, EventArgs e)
        {
            Button MsgBoxBtn = (sender as Button);
            MsgBoxRtn = MsgBoxBtn.Text.ToUpper();
            this.Close();
        }

        private static class HelperMethods
        {
            public static float ClipLo(float Input, float LimLo)
            {
                float Output = Input;
                if (Input < LimLo) Output = LimLo;
                return Output;
            }

            public static double ClipLo(double Input, double LimLo)
            {
                double Output = Input;
                if (Input < LimLo) Output = LimLo;
                return Output;
            }

            public static int ClipLo(int Input, int LimLo)
            {
                int Output = Input;
                if (Input < LimLo) Output = LimLo;
                return Output;
            }

        }

    }
}
