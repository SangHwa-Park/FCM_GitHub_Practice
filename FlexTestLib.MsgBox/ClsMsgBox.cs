using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlexTestLib.MsgBox
{
    public class ClsMsgBox
    {
        public static string Show(string Msg)
        {
            return Show("MsgBox", Msg, "OK", null);
        }

        public static string Show(string Msg, System.Drawing.Font MsgFont)
        {
            return Show("MsgBox", Msg, "OK", MsgFont);
        }

        public static string Show(string Prompt, string Msg)
        {
            return Show(Prompt, Msg, "OK", null);
        }

        public static string Show(string Prompt, string Msg, System.Drawing.Font MsgFont)
        {
            return Show(Prompt, Msg, "OK", MsgFont);
        }

        public static string Show(string Prompt, string Msg, string BtnText)
        {
            string StrRtn = "";
            MsgBoxForm MsgBox = new MsgBoxForm(BtnText);
            StrRtn = MsgBox.ShowForm(Prompt, Msg, null);
            return StrRtn;
        }

        public static string Show(string Prompt, string Msg, string BtnText, System.Drawing.Font MsgFont)
        {
            string StrRtn = "";
            MsgBoxForm MsgBox = new MsgBoxForm(BtnText);
            StrRtn = MsgBox.ShowForm(Prompt, Msg, MsgFont);
            return StrRtn;
        }

        public static string Show(string Prompt, string Msg, int PsnX, int PsnY)
        {
            return Show(Prompt, Msg, "OK", PsnX, PsnY, null);
        }

        public static string Show(string Prompt, string Msg, int PsnX, int PsnY, System.Drawing.Font MsgFont)
        {
            return Show(Prompt, Msg, "OK", PsnX, PsnY, MsgFont);
        }

        public static string Show(string Prompt, string Msg, string BtnText, int PsnX, int PsnY, System.Drawing.Font MsgFont)
        {
            string StrRtn = "";
            MsgBoxForm MsgBox = new MsgBoxForm(BtnText);
            StrRtn = MsgBox.ShowForm(Prompt, Msg, PsnX, PsnY, MsgFont);
            return StrRtn;
        }
    }
}
