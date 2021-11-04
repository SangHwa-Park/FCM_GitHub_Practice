using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Base;
namespace Test_Planner
{
    public static class Globals
    {
        public static bool Project_Changed = false;
        public static bool Load_CM_Initialze = false;
        public static bool Load_INI_Complete = false;

        public static string Path_Default = "C:\\ProgramData\\FlexTest\\GENTLE_BREED\\";
        public static string CMsheet_INI_Dir = find_Path("DEFAULT_PATH_CM");
        public static string PortInfo_INI_Dir = find_Path("DEFAULT_PATH_PORT");
        public static string Spara_Info = find_Path("DEFAULT_PATH_PORT");
        
        public static Excel_Base.TestConfig IniFile = new Excel_Base.TestConfig(); //Load CM sheet setting config from INI
        public static TestConfig_Spara Spara_config_INFO = new TestConfig_Spara(); //Load Port & S-Param setting config from INI
        public static Dictionary<string, Excel_Base.Band_Condition> DUT_CM = new Dictionary<string, Excel_Base.Band_Condition>();

        public static Dictionary<string, List<string>> Spara_TestDic = new Dictionary<string, List<string>>(); //key = TestID "IL", value = combination "Band","Index"
        public static Dictionary<string, List<string>> TX_TestDic = new Dictionary<string, List<string>>();
        public static Dictionary<string, List<string>> RX_TestDic = new Dictionary<string, List<string>>();
        public static Dictionary<string, List<string>> DC_TestDic = new Dictionary<string, List<string>>();
        public static Dictionary<string, List<string>> NOISE_TestDic = new Dictionary<string, List<string>>();

        public static bool LoadCM_completed = false;
        public static bool Kill_CM_Sheet = false;
        public static bool SPara_Plan_Generate = false;

        public static string Default_ANT = Default_ANT;
        public static string Default_ISO_GainModes = IniFile.Default_ISO_gainMode;

        public static Dictionary<string, Dictionary<string, List<TestCon>>> Expaned_Spara_Seq = new Dictionary<string, Dictionary<string, List<TestCon>>>();  //Band, Item, Testcondition (with temp, port extended)
        public static SortedList<string, Dictionary<string, List<Spara_Trigger_Group>>> Spara_TestCon = new SortedList<string, Dictionary<string, List<Spara_Trigger_Group>>>();
        public static int Spara_TestTrigger_Count = 0;

        private static void ClearINI()
        {
            IniFile = new Excel_Base.TestConfig();
        }

        private static string find_Path(string default_key)
        {
            string PathDefault = @"C:\ProgramData\FlexTest\GENTLE_BREED\GentleBreed_default.ini";
            string Return_Path = "";
            
            try
            {
                using (StreamReader sr = new StreamReader(PathDefault))
                {
                    string Line;
                    string Key;
                    string Val;

                    while ((Line = sr.ReadLine()) != null)
                    {
                        bool Valid_Line = (Line.Contains('=') && !IsComment(Line));
                        if (!Valid_Line) continue;

                        string[] Substr = Line.Split('=');
                        Key = Substr[0].Trim();
                        Val = Substr[1].Trim();
                        RemoveComment(ref Val);

                        if (Key.ToUpper().Contains(default_key))
                        {
                            Return_Path = Val.Trim();
                            return Return_Path;
                        }
                    }
                }

            }
            catch (Exception)
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error: Initialize TestConfig()");
                ErrMsg.AppendFormat("\n ");
                ErrMsg.AppendFormat("\nCannot open Default INI file ");
                ErrMsg.AppendFormat("\n<C:\\ProgramData\\FlexTest\\GENTLE_BREED\\GentleBreed_default.ini>");
                MessageBox.Show("Error on file loading in initialization", ErrMsg.ToString());
                Environment.Exit(0);
                throw;
            }

            return Return_Path;

        }

        private static bool IsComment(string strings)
        {
            bool IsCom = false;
            string[] CommentIndicators = new string[] { "//", "#", "%" };

            foreach (string ComInd in CommentIndicators)
            {
                if (strings.Trim().StartsWith(ComInd))
                {
                    IsCom = true;
                    return IsCom;
                }
            }
            return IsCom;
        }

        private static void RemoveComment(ref string Val)
        {
            int SplitIndex;

            if (Val.Contains(" //")) // Remove Chars to the RIGHT of "//".
            {
                SplitIndex = Val.IndexOf(" //");
                Val = Val.Substring(0, SplitIndex).Trim();
            }

            if (Val.Contains("//")) // Remove Chars to the RIGHT of "//" (but NOT for "://" --> such as "https://").
            {
                SplitIndex = Val.IndexOf("//");
                int SplitIndexMinus1 = ClipLo(SplitIndex - 1, 0);
                if (Val.Substring(SplitIndexMinus1, 1) != ":") Val = Val.Substring(0, SplitIndex).Trim();
            }
        }
        private static int ClipLo(int Input, int LimLo)
        {
            int Output = Input;
            if (Input < LimLo) Output = LimLo;
            return Output;
        }

    }
    
}
