using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlexTestLib.MsgBox;

namespace Excel_Base
{
    public class TestConfig
    {
        public string PFN = "";
        public Dictionary<string, string> Dic = new Dictionary<string, string>();
        public Dictionary<string, string> Frequency_table = new Dictionary<string, string>();
        public Dictionary<string, List<string>> Band_CAs = new Dictionary<string, List<string>>();
        public Dictionary<string, string> Band_RXOUTs = new Dictionary<string, string>();
        public List<string> Band_Sheet_Name = new List<string>();
        public string Header_Start_ID = "";
        public List<string> Selected_Headers = new List<string>();
        public List<string> RX_GainModes = new List<string>();
        public string Default_ISO_gainMode = "";

        //CM Sheet parameter defined
        public List<string> Test_SpecID;
        public List<string> Band;
        public List<string> CA_Band2;
        public List<string> CA_Band3;
        public List<string> CA_Band4;
        public List<string> Direction;
        public List<string> Description;
        public List<string> Parameter;
        public List<string> Input_Port;
        public List<string> Output_Port;
        public List<string> LNA_Gain_Mode;
        public List<string> Vbatt;
        public List<string> Vdd_LNA;

        public List<string> TXIn_VSWR;
        public List<string> ANTout_VSWR;
        public List<string> ANTIn_VSWR;
        public List<string> RXOut_VSWR;
        public List<string> Temperature;

        public List<string> Start_Freq;
        public List<string> Stop_Freq;

        public List<string> IBW;
        public List<string> PA_MODE;
        public List<string> TXBand_In_RXtest;
        public List<string> Target_Pout;
        public List<string> Signal_Standard;
        public List<string> Waveform_Category;
        public List<string> MPR;
        public List<string> Test_Limit_L;
        public List<string> Test_Limit_Typ;
        public List<string> Test_Limit_U;
        public List<string> Unit;
        public List<string> Compliance;

        public TestConfig()
        {
            this.clear();
            //this.LoadINI();
        }
        public void clear()
        {
            this.Dic = new Dictionary<string, string>();
            this.Frequency_table = new Dictionary<string, string>();
            this.Band_CAs = new Dictionary<string, List<string>>();
            this.Band_RXOUTs = new Dictionary<string, string>();
            this.RX_GainModes = new List<string>();
            this.Default_ISO_gainMode = "";

            this.Test_SpecID = new List<string>();
            this.Band = new List<string>();
            this.CA_Band2 = new List<string>();
            this.CA_Band3 = new List<string>();
            this.CA_Band4 = new List<string>();
            this.Direction = new List<string>();
            this.Description = new List<string>();
            this.Parameter = new List<string>();
            this.Input_Port = new List<string>();
            this.Output_Port = new List<string>();
            this.LNA_Gain_Mode = new List<string>();
            this.Vbatt = new List<string>();
            this.Vdd_LNA = new List<string>();

            this.TXIn_VSWR = new List<string>();
            this.ANTout_VSWR = new List<string>();
            this.ANTIn_VSWR = new List<string>();
            this.RXOut_VSWR = new List<string>();
            this.Temperature = new List<string>();

            this.Start_Freq = new List<string>();
            this.Stop_Freq = new List<string>();

            this.IBW = new List<string>();
            this.PA_MODE = new List<string>();
            this.TXBand_In_RXtest = new List<string>();
            this.Target_Pout = new List<string>();
            this.Signal_Standard = new List<string>();
            this.Waveform_Category = new List<string>();
            this.MPR = new List<string>();
            this.Test_Limit_L = new List<string>();
            this.Test_Limit_Typ = new List<string>();
            this.Test_Limit_U = new List<string>();
            this.Unit = new List<string>();
            this.Compliance = new List<string>();
        }

        public void LoadINI(string RootDir)
        {
            try
            {
                if(RootDir == "") RootDir = "C:\\ProgramData\\FlexTest\\GENTLE_BREED\\INI\\CMsheet_config.ini";
                this.Dic = new Dictionary<string, string>();
                this.Frequency_table = new Dictionary<string, string>();
                this.Band_Sheet_Name = new List<string>();
                this.PFN = RootDir;

                using (StreamReader sr = new StreamReader(this.PFN))
                {
                    string Line;
                    string Key;
                    string Val;

                    while ((Line = sr.ReadLine()) != null)
                    {
                        bool ValidLine = (Line.Contains('=') && !IsComment(Line));
                        if (!ValidLine) continue;

                        string[] Substr = Line.Split('=');
                        Key = Substr[0].Trim();
                        Val = Substr[1].Trim();

                        this.RemoveComment(ref Val);
                        this.AddParam_FromINI(Key, Val);

                        if (Key.ToUpper().Contains("HEADER_INDEX"))
                        {
                            this.Header_Start_ID = Val.Trim().ToUpper();
                        }
                    }
                }
            }
            catch
            {
                this.Dic.Clear();
                ClsMsgBox.Show("Error during Loading General Ini File:\n" + PFN);
                Environment.Exit(0);
            }
        }

        private void AddParam_FromINI(string key, string val)
        {
            string keyUP = key.ToUpper();
            bool IsNewKey = (!this.Dic.ContainsKey(keyUP));
            bool CA_IsNewKey = (!this.Band_CAs.ContainsKey(keyUP));


            if (key.ToUpper().Contains("CA_TX_B"))
            {
                //Build All INI setting param list include value

                string[] RX_CAs = val.ToUpper().Trim().Split(',');
                List<string> RXCA_List = new List<string>();

                foreach (string each_RX in RX_CAs)
                {
                    string[] Band = each_RX.Split('_');
                    RXCA_List.Add(Band[1].Trim());
                }

                this.Band_CAs.Add(keyUP, RXCA_List);

            }

            if (key.ToUpper().Contains("RXOUT_") && !key.ToUpper().Contains("VSWR"))
            {
                //Build All INI setting param list include value

                string[] TEMP = key.ToUpper().Trim().Split('_');
                string OUT_Port_Key = TEMP[1];
                string OUT_BAND = val.ToUpper().Trim();
                this.Band_RXOUTs.Add(OUT_Port_Key, OUT_BAND);
            }

            if (key.ToUpper().Contains("DUT_RX_GAIN_MODE_DEFINE"))
            {
                string[] RX_Gains_modes = val.ToUpper().Trim().Split(',');

                foreach (string Modes in RX_Gains_modes)
                {
                    this.RX_GainModes.Add(Modes.ToUpper().Trim());
                    this.RX_GainModes = this.RX_GainModes.Distinct().ToList();
                }
            }

            if (key.ToUpper().Contains("DEFAULT_ISO_GAIN_MODE"))
            {
                this.Default_ISO_gainMode = val.ToUpper().Trim();
            }

            if (IsNewKey) //Build All INI setting param list include value
            {
                this.Dic.Add(keyUP, val);
                if (key.ToUpper().Contains("FREQ_TX") || key.ToUpper().Contains("FREQ_RX")) this.Frequency_table.Add(keyUP, val);
            }
            else if (!IsNewKey && !key.ToUpper().Contains("CA_TX_B"))
            {
                this.Dic.Remove(keyUP);
                this.Dic.Add(keyUP, val);

                this.Frequency_table.Remove(keyUP);
                this.Frequency_table.Add(keyUP, val);

                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error Dupplicated Key");
                ErrMsg.AppendFormat("\nThis item has duplicated : " + key + ", value=" + val);
                ClsMsgBox.Show("Error on INI file loading", ErrMsg.ToString());
            }

            if (key.ToUpper().Contains("CM_SHEET_READ")) //Add specific Band to build plan
            {
                string[] Target_Band = val.Split(',');
                for (int i = 0; i < Target_Band.Length; i++)
                {
                    string Band_ToInsert = Target_Band[i].Trim(); //sustain Lowercase and Uppercase character from INI, it means you should enter correct sheet name on INI file. 
                    if (!this.Band_Sheet_Name.Contains(Band_ToInsert)) this.Band_Sheet_Name.Add(Band_ToInsert);

                }
            }

            if (key.ToUpper().Contains("TEST_SPECID")) this.Test_SpecID = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("BAND") && !key.ToUpper().Contains("RX") && !key.ToUpper().Contains("CA")) this.Band = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("BAND") && key.ToUpper().Contains("CA") && key.ToUpper().Contains("2")) this.CA_Band2 = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("BAND") && key.ToUpper().Contains("CA") && key.ToUpper().Contains("3")) this.CA_Band3 = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("BAND") && key.ToUpper().Contains("CA") && key.ToUpper().Contains("4")) this.CA_Band4 = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("DESCRIPTION")) this.Description = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("PARAMETER")) this.Parameter = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("INPUT_PORT")) this.Input_Port = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("OUTPUT_PORT")) this.Output_Port = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("LNA_GAIN_MODE")) this.LNA_Gain_Mode = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("VBATT")) this.Vbatt = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("VDD_LNA")) this.Vdd_LNA = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TXIN_VSWR")) this.TXIn_VSWR = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("ANTOUT_VSWR")) this.ANTout_VSWR = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("ANTIN_VSWR")) this.ANTIn_VSWR = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("RXOUT_VSWR")) this.RXOut_VSWR = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TEMPERATURE")) this.Temperature = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("START_FREQ")) this.Start_Freq = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("STOP_FREQ")) this.Stop_Freq = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("IBW")) this.IBW = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("PA_MODE")) this.PA_MODE = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TXBAND_IN_RXTEST")) this.TXBand_In_RXtest = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TARGET_POUT")) this.Target_Pout = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("SIGNAL_STANDARD")) this.Signal_Standard = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("WAVEFORM_CATEGORY")) this.Waveform_Category = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("MPR")) this.MPR = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TEST_LIMIT_L")) this.Test_Limit_L = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TEST_LIMIT_TYP")) this.Test_Limit_Typ = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("TEST_LIMIT_U")) this.Test_Limit_U = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("UNIT")) this.Unit = this.Load_MatchingList(val);
            if (key.ToUpper().Contains("COMPLIANCE")) this.Compliance = this.Load_MatchingList(val);

        }

        private List<string> Load_MatchingList(string candidated_header_name)
        {
            List<string> Listed_Names = new List<string>();
            string[] HeaderName_splited = candidated_header_name.Split(',');

            for (int i = 0; i < HeaderName_splited.Length; i++)
            {
                if (HeaderName_splited[i].Trim().ToUpper() != "") Listed_Names.Add(HeaderName_splited[i].Trim().ToUpper());
            }

            return Listed_Names;
        }

        public bool IsComment(string strings)
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

        public void RemoveComment(ref string Val)
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

        public static int ClipLo(int Input, int LimLo)
        {
            int Output = Input;
            if (Input < LimLo) Output = LimLo;
            return Output;
        }

    }

}
