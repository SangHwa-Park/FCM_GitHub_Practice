using System;
using System.Collections.Generic;
using System.Linq;
using System.IO; //FILE IO as like as "stream reader"
using System.Data;
using System.Text;
using Excel_Base;
using Excel = Microsoft.Office.Interop.Excel; //for excel control
using System.Threading.Tasks;
using System.Globalization;
using Test_Planner;

namespace S_para_planner
{
    class Build_Spara
    {
        Excel_File Spara_Excel;
        Dictionary<string, int> Spara_Header_Dic;
        int Testnum = 4;

        string Spara_TestCon = "Condition_FBAR";
        string Spara_TestData = "TEST_DATA";
        string Spara_InfoPort = "PORT_INFO";
        string Spara_InfoMipi = "MIPI_INFO";

        public Build_Spara()
        {
            Excel_File Spara_plan = new Excel_File(@"C:\ProgramData\FlexTest\GENTLE_BREED\default_Spara_Plan");
            List<string> Spara_Sheet_List = new List<string>();
            Spara_Sheet_List.Add(Spara_TestCon);
            Spara_Sheet_List.Add(Spara_TestData);
            Spara_Sheet_List.Add(Spara_InfoPort);
            Spara_Sheet_List.Add(Spara_InfoMipi);

            Spara_plan.Create_SheetName(Spara_Sheet_List, Spara_TestCon);
            this.Spara_Excel = Spara_plan;
            this.Spara_Excel.Worksheet = Spara_Excel.getSheet(this.Spara_TestCon);

            this.Spara_Excel.App.Visible = false;
        }

        public void GeneratePlan(SortedList<string, Dictionary<string, List<Spara_Trigger_Group>>> Spara_TestCon)
        {
            //this.Init_ProgressBar(Globals.Spara_TestTrigger_Count);

            Popup_ProgressBar bar1 = new Popup_ProgressBar(Globals.Spara_TestTrigger_Count);
            bar1.Show();
            int Progress_bar_step = 0;

            Build_Spara_Plan_Header("VCC_V", 15); //Header ocuppied row 1~3         
            int TD_RowIndex = 4;
            int Before_Index = 0;

            this.Spara_Excel.App.ScreenUpdating = false;

            foreach (string Temperature in Spara_TestCon.Keys)
            {
                //Temp Header line with formatting
                StringBuilder Title_Temp = new StringBuilder();
                Title_Temp.AppendFormat("Test Temperature : {0}", Temperature);
                Spara_Excel.Cformat_LineColor(this.Spara_Excel.Worksheet, TD_RowIndex, this.Spara_Header_Dic.Keys.Count, Excel.XlThemeColor.xlThemeColorAccent1, 0);
                TD_RowIndex = Spara_Excel.Cell_WriteTitle(this.Spara_Excel.Worksheet, Title_Temp.ToString(), TD_RowIndex, this.Spara_Header_Dic["Test Parameter"], this.Spara_Header_Dic.Keys.Count);
                //End Temperature Title

                foreach (string Band in Spara_TestCon[Temperature].Keys)
                {
                    //Add Band Header Line here
                    StringBuilder Title_Band = new StringBuilder();
                    Title_Band.AppendFormat("Band : {0}", Band);
                    Spara_Excel.Cformat_LineColor(this.Spara_Excel.Worksheet, TD_RowIndex, this.Spara_Header_Dic.Keys.Count, Excel.XlThemeColor.xlThemeColorAccent5, -0.499984740745262);
                    TD_RowIndex = Spara_Excel.Cell_WriteTitle(this.Spara_Excel.Worksheet, Title_Band.ToString(), TD_RowIndex, this.Spara_Header_Dic["Test Parameter"], this.Spara_Header_Dic.Keys.Count);
                    //End Band Title

                    List<Spara_Trigger_Group> WriteGroup = new List<Spara_Trigger_Group>();
                    WriteGroup = Spara_TestCon[Temperature][Band];

                    foreach (Spara_Trigger_Group Each_Trigger in WriteGroup)
                    {
                        Before_Index = TD_RowIndex;

                        List<List<string>> Testcons = new List<List<string>>();

                        Testcons.Add(Build_Trigger_Header(Each_Trigger, Band));

                        foreach (TestCon Each_Testcon in Each_Trigger.TestCon_List)
                        {
                            Testcons.Add(Build_Test_condition(Each_Testcon, Band));
                            if(Each_Testcon.Parameter.Contains("RX_Gain_CA_G"))
                            {
                                TestCon New_RX_CA_min = new TestCon();
                                New_RX_CA_min = Each_Testcon.Clone();
                                New_RX_CA_min.Spara_Searchmethod = "MIN";
                                Testcons.Add(Build_Test_condition(New_RX_CA_min, Band));

                            }
                        }
                        
                        TD_RowIndex = Spara_Excel.Cell_WriteTrigger(this.Spara_Excel.Worksheet, TD_RowIndex, 1, Testcons);
                        Spara_Excel.Cformat_PartialColor(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Test Parameter"], this.Spara_Header_Dic["Convert_SIGN_FOR_ISO"], Excel.XlThemeColor.xlThemeColorAccent1, 0.799981688894314);
                        Spara_Excel.Cformat_PartialColor(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["TRX_ON"], this.Spara_Header_Dic["ASM_UAT"], Excel.XlThemeColor.xlThemeColorAccent6, 0.799981688894314);
                        Spara_Excel.Cformat_PartialColor(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["TXREG00"], this.Spara_Header_Dic["REGCUSTOM"], Excel.XlThemeColor.xlThemeColorAccent2, 0.799981688894314);
                        Spara_Excel.Cformat_Color(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Port_Define"], Excel.XlThemeColor.xlThemeColorAccent4, 0.599993896298105);
                        Spara_Excel.Cformat_Color(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Channel Number"], Excel.XlThemeColor.xlThemeColorAccent4, 0.599993896298105);
                        Spara_Excel.Cformat_Color(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Set_Temp"], Excel.XlThemeColor.xlThemeColorAccent4, 0.599993896298105);

                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Para.Spec"], 15);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Test Parameter"], 15);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Parameter Header"], 55);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Input Port"], 12);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Output Port"], 12);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Port_Define"], 14);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Start_Freq"], 10);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["Stop_Freq"], 10);
                        Spara_Excel.Cformat_Width(this.Spara_Excel.Worksheet, Before_Index, this.Spara_Header_Dic["MIPI_Description"], 20);

                        StringBuilder new_text = new StringBuilder();
                        new_text.AppendFormat("Build Spara : Temp={0},  Band={1},  Input = {2}, Output = {3},  Row={4}", Temperature, Testcons[0][4], Testcons[0][12], Testcons[0][13], Convert.ToString(TD_RowIndex));
                        //this.Progress_perform_Spara(new_text.ToString());

                        Progress_bar_step++;
                        double Progress_rate = ((double) Progress_bar_step / (double) Globals.Spara_TestTrigger_Count) * 100;

                        StringBuilder Rate = new StringBuilder();
                        Rate.AppendFormat("Progress : {0}", Progress_rate.ToString("0.0"));
                        bar1.execute_step(Rate.ToString(), new_text.ToString());

                    }
                }
            }

            bar1.Done(this.Spara_Excel);

            //this.Spara_Excel.App.ScreenUpdating = true;
            //this.Spara_Excel.Worksheet.Activate();
            //this.Spara_Excel.App.Visible = true;

        }

        private List<string> Build_Test_condition(TestCon Each_Testcon, string Band)
        {
            List<string> Trigger_Setting = new List<string>();
            int Col_length = this.Spara_Header_Dic.Count;

            foreach (string item in this.Spara_Header_Dic.Keys)
            {
                StringBuilder Description = new StringBuilder();
                //Description.AppendFormat("{0} to {1}, RX : {2}", Each_Testcon.TX_Input, Each_Testcon.ANT_Output, Each_Testcon.RX_Output);

                switch (item)
                {
                    case "Enable": Trigger_Setting.Add("x"); break;
                    case "Test_Num": Trigger_Setting.Add(Convert.ToString(this.Testnum)); this.Testnum++; break;
                    case "Test Mode": Trigger_Setting.Add("FBAR"); break;
                    case "Spec Sheet Name": Trigger_Setting.Add(Band); break;
                    case "Para.Spec": Trigger_Setting.Add(Each_Testcon.Test_SpecID); break;
                    case "PA_BAND": Trigger_Setting.Add(convert_Band_Name(Each_Testcon.Band)); break;                 //need modification
                    case "POWER_MODE": Trigger_Setting.Add(Each_Testcon.PA_MODE); break;
                    //case "TUNABLE_BAND": Trigger_Setting.Add(Each_Trigger.Tunable_Band); break;       //need modification
                    case "LNA_GAIN": Trigger_Setting.Add(Each_Testcon.LNA_Gain_Mode); break;
                    case "Test Parameter": 
                        if(Each_Testcon.Parameter.ToUpper().Contains("RIPPLE"))
                        {
                            Trigger_Setting.Add("RIPPLE_BETWEEN");
                        }
                        else if (Each_Testcon.Parameter.ToUpper().Contains("PHASE"))
                        {
                            Trigger_Setting.Add("PHASE_AT");
                        }
                        else if (Each_Testcon.Parameter.ToUpper().Contains("VSWR"))
                        {
                            Trigger_Setting.Add("INPUT_VSWR");
                        }
                        else if (Each_Testcon.Parameter.ToUpper().Contains("MU_FACTOR"))
                        {
                            Trigger_Setting.Add("MU-FACTOR");
                        }
                        else if (Each_Testcon.Parameter.ToUpper().Contains("K_FACTOR"))
                        {
                            Trigger_Setting.Add("K-FACTOR");
                        }
                        else if (Each_Testcon.Parameter.ToUpper().Contains("GROUP_DELAY"))
                        {
                            Trigger_Setting.Add("RIPPLE_BETWEEN");
                        }
                        else
                        {
                            Trigger_Setting.Add("MAG_BETWEEN");
                        }
                        break;
                    case "Parameter Header":
                        StringBuilder str = new StringBuilder();

                        if (Each_Testcon.Test_Name.Contains("\n"))
                        {
                            string temp = Each_Testcon.Test_Name;
                            Each_Testcon.Test_Name = temp.Replace("\n", " ");
                        }
                        str.AppendFormat("[{0}] {1}", Each_Testcon.Parameter, Each_Testcon.Test_Name);
                        Trigger_Setting.Add(str.ToString()); 
                        break;
                    case "Input Port": Trigger_Setting.Add(Each_Testcon.Input_Port); break;
                    case "Output Port": Trigger_Setting.Add(Each_Testcon.Output_Port); break;
                    case "S-Parameter": Trigger_Setting.Add(Each_Testcon.Spara_ID); break;
                    case "DM_S-Param": Trigger_Setting.Add(Each_Testcon.Spara_ID_DNM); break;
                    case "Channel Number": Trigger_Setting.Add("1"); break;
                    case "Start_Freq": Trigger_Setting.Add(reFormat_Frequency(Each_Testcon.Start_Freq)); break;
                    case "Stop_Freq": Trigger_Setting.Add(reFormat_Frequency(Each_Testcon.Stop_Freq)); break;
                    case "Search_Method": Trigger_Setting.Add(Each_Testcon.Spara_Searchmethod); break;
                    case "MIN_Limit": Trigger_Setting.Add(Each_Testcon.Test_Limit_L); break;
                    case "TYP_Limit": Trigger_Setting.Add(Each_Testcon.Test_Limit_Typ); break;
                    case "MAX_Limit": Trigger_Setting.Add(Each_Testcon.Test_Limit_U); break;
                    case "Use_Previous": Trigger_Setting.Add(" "); break;
                    case "Set_Temp": Trigger_Setting.Add(Each_Testcon.Temperature); break;
                    case "Convert_SIGN_FOR_ISO": Trigger_Setting.Add(Each_Testcon.Spara_ConvertSign); break;
                    default: Trigger_Setting.Add(""); break;

                }
            }

            return Trigger_Setting;
        }

        private string reFormat_Frequency(string frequency)
        {
            string reFormatted_Freq = "";
            try
            {
                float f_frequency = Convert.ToSingle(frequency);
                reFormatted_Freq = f_frequency.ToString("F1") + " M";
            }
            catch
            {
                reFormatted_Freq = frequency + " M";
            }

            return reFormatted_Freq;
        }

        public bool Get_IsTDD(string Band)
        {
            string TargetBand = "B" + Band;
            bool IsTDD = false;
            bool find_TX = false;
            bool find_RX = false;

            if (TargetBand == "B40") TargetBand = "B40F";

            List<float> TX_Freq = new List<float>();
            List<float> RX_Freq = new List<float>();

            foreach (string band_name in Globals.IniFile.Frequency_table.Keys)
            {
                string[] Table_string = band_name.Split('_');
                string Table_band = Table_string[2].Trim().ToUpper();
                string Direction = Table_string[1].Trim().ToUpper();

                if (Table_band.Trim().ToUpper() == TargetBand.Trim().ToUpper())
                {
                    string Freq_range = Globals.IniFile.Frequency_table[band_name];
                    string[] Frequencys = Freq_range.Split(',');

                    if (Direction == "TX")
                    {
                        foreach (string freq in Frequencys)
                        {
                            TX_Freq.Add(Convert.ToSingle(freq.Trim()));
                        }
                        TX_Freq.Sort();
                        find_TX = true;
                    }
                    else if(Direction == "RX")
                    {
                        foreach (string freq in Frequencys)
                        {
                            RX_Freq.Add(Convert.ToSingle(freq.Trim()));
                        }
                        RX_Freq.Sort();
                        find_RX = true;
                    }
                }

                if (find_TX && find_RX) break;

            }

            if (find_TX && find_RX)
            {
                if (RX_Freq[0] >= TX_Freq[0] && RX_Freq[1] <= TX_Freq[1])
                {
                    IsTDD = true;
                    return IsTDD;
                }
                else if (TargetBand.Contains("B53") || TargetBand.Contains("B1P6G"))
                {
                    IsTDD = true;
                    return IsTDD;
                }
            }

            return false;
        }

        private string convert_Band_Name(string band)
        {
            string Converted_band = band;
            if (band.Contains('.')) { Converted_band = band.Replace('.', 'P'); }
            return Converted_band;
        }

        private void Compare_PentaHexa(ref string TX_band_desc, ref string RX_band_desc)
        {
            List<string> PentaFlex = new List<string>();
            PentaFlex.Add("B1_");
            PentaFlex.Add("B3_");
            PentaFlex.Add("B32_");
            PentaFlex.Add("B40_");
            PentaFlex.Add("B40F_");
            PentaFlex.Add("B32_");

            List<string> HexaFlex = new List<string>();
            HexaFlex.Add("B25_");
            HexaFlex.Add("B2_");
            HexaFlex.Add("B30_");
            HexaFlex.Add("B66_");
            HexaFlex.Add("B4_");

            if (TX_band_desc.Contains("B7_") || RX_band_desc.Contains("B7_"))
            {
                foreach (var Penta_CA in PentaFlex)
                {
                    if (TX_band_desc.Contains(Penta_CA) || RX_band_desc.Contains(Penta_CA))
                    {
                        TX_band_desc = TX_band_desc.Replace("B7_", "B7PENTA_");
                        RX_band_desc = RX_band_desc.Replace("B7_", "B7PENTA_");
                        break;
                    }
                }

                foreach (var Hexa_CA in HexaFlex)
                {
                    if (TX_band_desc.Contains(Hexa_CA) || RX_band_desc.Contains(Hexa_CA))
                    {
                        TX_band_desc = TX_band_desc.Replace("B7_", "B7HEXA_");
                        RX_band_desc = RX_band_desc.Replace("B7_", "B7HEXA_");
                        break;
                    }
                }

                TX_band_desc = TX_band_desc.Replace("B7_", "B7PENTA_");
                RX_band_desc = RX_band_desc.Replace("B7_", "B7PENTA_");

            }
            else if (TX_band_desc.Contains("B41_") || RX_band_desc.Contains("B41_") ||
                     TX_band_desc.Contains("B41F_") || RX_band_desc.Contains("B41F_"))
            {
                foreach (var Penta_CA in PentaFlex)
                {
                    if (TX_band_desc.Contains(Penta_CA) || RX_band_desc.Contains(Penta_CA))
                    {
                        TX_band_desc = TX_band_desc.Replace("B41_", "B41FPENTA_");
                        RX_band_desc = RX_band_desc.Replace("B41_", "B41FPENTA_");
                        break;
                    }
                }

                foreach (var Hexa_CA in HexaFlex)
                {
                    if (TX_band_desc.Contains(Hexa_CA) || RX_band_desc.Contains(Hexa_CA))
                    {
                        TX_band_desc = TX_band_desc.Replace("B41_", "B41FHEXA_");
                        RX_band_desc = RX_band_desc.Replace("B41_", "B41FHEXA_");
                        break;
                    }
                }

                TX_band_desc = TX_band_desc.Replace("B41_", "B41FPENTA_");
                RX_band_desc = RX_band_desc.Replace("B41_", "B41FPENTA_");
            }
            else if (TX_band_desc.Contains("B40_") || RX_band_desc.Contains("B40_"))
            {
                TX_band_desc = TX_band_desc.Replace("B40_", "B40F_");
                RX_band_desc = RX_band_desc.Replace("B40_", "B40F_");
            }
        }

        private List<string> Build_Trigger_Header(Spara_Trigger_Group Each_Trigger, string Band)
        {
            List<string> Trigger_Setting = new List<string>();
            int Col_length = this.Spara_Header_Dic.Count;

            Spara_Path Path_info = GetPath_from_Trigger(Each_Trigger);

            foreach (string item in this.Spara_Header_Dic.Keys)
            {
                StringBuilder Description = new StringBuilder();
                Description.AppendFormat("{0} to {1}, OUT: {2}", Path_info.TX_BAND, Path_info.TX_ANT, Path_info.RX_CA_BAND);
                Compare_PentaHexa(ref Path_info.TX_BAND, ref Path_info.RX_CA_BAND);

                switch (item)
                {
                    case "Enable": Trigger_Setting.Add("x"); break;
                    case "Test_Num": Trigger_Setting.Add(Convert.ToString(this.Testnum)); this.Testnum++; break;
                    case "Test Mode": Trigger_Setting.Add("DC"); break;
                    case "Spec Sheet Name": Trigger_Setting.Add(Band); break;
                    case "Para.Spec": Trigger_Setting.Add("SETUP"); break;
                    case "PA_BAND": Trigger_Setting.Add(convert_Band_Name(Each_Trigger.TestCon_List[0].Band)); break;                 //need modification
                    case "POWER_MODE": Trigger_Setting.Add(Each_Trigger.PowerMode); break;
                    case "TUNABLE_BAND": Trigger_Setting.Add(Each_Trigger.CA_Case); break;       //need modification
                    case "LNA_GAIN": Trigger_Setting.Add(Path_info.LNA_GAIN); break;
                    case "Test Parameter": Trigger_Setting.Add("SETUP_TRIG"); break;
                    case "Parameter Header": Trigger_Setting.Add(Each_Trigger.Status_File); break;
                    case "Input Port": Trigger_Setting.Add(Each_Trigger.Test_Input); break;
                    case "Output Port": Trigger_Setting.Add(Each_Trigger.Test_Output); break;
                    case "Port_Define": 
                        Trigger_Setting.Add(Each_Trigger.Ports_Sequence);

                        foreach (TestCon each_TC in Each_Trigger.TestCon_List)
                        {
                            string prefix = "";
                            int first = 0;
                            int Last = 0;

                            if (each_TC.Spara_ID.ToUpper().Contains("S") && each_TC.Spara_ID.Trim().Length < 5)
                            {
                                first = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 1));
                                Last = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 1, 1));
                                prefix = "S";
                            }
                            else if (each_TC.Spara_ID.ToUpper().Contains("GDEL") && each_TC.Spara_ID.Trim().Length < 8)
                            {
                                first = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 1));
                                Last = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 1, 1));
                                prefix = "GDEL";
                            }
                            else
                            {
                                first = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 4, 2));
                                Last = Convert.ToInt32(each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 2));
                                if (each_TC.Spara_ID.ToUpper().Contains("S")) prefix = "S";
                                if (each_TC.Spara_ID.ToUpper().Contains("GDEL")) prefix = "GDEL";
                            }

                            int DNM_first = Each_Trigger.Ports_Assigned[first];
                            int DNM_Last = Each_Trigger.Ports_Assigned[Last];

                            each_TC.Spara_ID_DNM = prefix + DNM_first + DNM_Last;
                        }
                        break;               
                    case "Channel Number": Trigger_Setting.Add("1"); break;
                    case "Set_Temp": Trigger_Setting.Add(Each_Trigger.Tempearature); break;
                    case "MIPI_Description": Trigger_Setting.Add(Description.ToString()); break;
                    case "Convert_SIGN_FOR_ISO": Trigger_Setting.Add("OFF"); break;
                    case "TRX_ON": Trigger_Setting.Add(Path_info.TRX_ON_Direction); break;
                    case "TX_MODE": Trigger_Setting.Add(Path_info.TX_Tech); break;
                    case "TX_BAND": Trigger_Setting.Add(Path_info.TX_BAND); break;
                    case "TX_INPUT": Trigger_Setting.Add(Path_info.TX_INPUT); break;
                    case "TX_OUTPUT": Trigger_Setting.Add(Path_info.TX_ANT); break;
                    case "PA_MODE": Trigger_Setting.Add(Path_info.PA_MODE); break;
                    case "RX_BAND": Trigger_Setting.Add(Path_info.RX_CA_BAND); break;
                    case "RX_INPUT": Trigger_Setting.Add(Path_info.RX_ANT); break;
                    case "RX_OUTPUT": Trigger_Setting.Add(Path_info.RX_OUT); break;
                    case "LNA_MODE": Trigger_Setting.Add(Path_info.LNA_GAIN); break;
                    case "TDD_PRIORITY": Trigger_Setting.Add(Path_info.TDD_Priority); break;
                    case "ASM_ANT1":
                        if (Path_info.ASM.ContainsKey("ANT1"))
                        {
                            Trigger_Setting.Add(Path_info.ASM["ANT1"]);
                        }
                        else { Trigger_Setting.Add(""); }
                        break;
                    case "ASM_ANT2":
                        if (Path_info.ASM.ContainsKey("ANT2"))
                        {
                            Trigger_Setting.Add(Path_info.ASM["ANT2"]);
                        }
                        else { Trigger_Setting.Add(""); }
                        break;
                    case "ASM_UAT":
                        if (Path_info.ASM.ContainsKey("UAT"))
                        {
                            Trigger_Setting.Add(Path_info.ASM["UAT"]);
                        }
                        else { Trigger_Setting.Add(""); }
                        break;

                    default: Trigger_Setting.Add(""); break;
                }
            }

            

            return Trigger_Setting;
        }

        private void Build_Spara_Plan_Header(string DC_identifier_1st, int DC_slot_num)
        {
            int PXIe_DC_slot_index = DC_slot_num;
            int PXIe_DC_current_index = DC_slot_num;
            bool find_slot = false;

            this.Spara_Header_Dic = Create_spara_Header_List();
            int VCCSlot_Index = GetVCCslot(this.Spara_Header_Dic, DC_identifier_1st); // ex) DC_identifier_1st = "VCC_V"

            List<string> temp_Line2 = new List<string>();
            List<string> temp_Line3 = new List<string>();

            foreach (var item in Spara_Header_Dic.Keys)
            {
                if ((item == DC_identifier_1st || find_slot) && PXIe_DC_slot_index < DC_slot_num + 3)
                {
                    temp_Line2.Add(Convert.ToString(PXIe_DC_slot_index));
                    PXIe_DC_slot_index++;
                    find_slot = true;
                }
                else if (find_slot && PXIe_DC_slot_index > PXIe_DC_current_index)
                {
                    temp_Line2.Add(Convert.ToString(PXIe_DC_current_index));
                    PXIe_DC_current_index++;
                }
                else
                {
                    temp_Line2.Add("");
                }
                temp_Line3.Add(item);
            }

            this.Spara_Excel.Cell_WriteHeader(this.Spara_Excel.getSheet(this.Spara_TestCon), 2, 1, temp_Line2, true);
            this.Spara_Excel.Cell_WriteHeader(this.Spara_Excel.getSheet(this.Spara_TestCon), 3, 1, temp_Line3, true);
        }

        private Dictionary<string, int> Create_spara_Header_List()
        {
            Dictionary<string, int> Spara_Header = new Dictionary<string, int>();

            //---------------Basic information---------------//

            Spara_Header.Add("#Test_Range", 1);     //1 Header Start
            Spara_Header.Add("Enable", 2);          //2
            Spara_Header.Add("Test_Num", 3);        //3
            Spara_Header.Add("Test Mode", 4);       //4
            Spara_Header.Add("Spec Sheet Name", 5); //5
            Spara_Header.Add("Para.Spec", 6);       //6
            Spara_Header.Add("PA_BAND", 7);         //7
            Spara_Header.Add("POWER_MODE", 8);      //8
            Spara_Header.Add("TUNABLE_BAND", 9);    //9
            Spara_Header.Add("LNA_GAIN", 10);       //10
            Spara_Header.Add("Test Parameter", 11); //11
            Spara_Header.Add("Parameter Header", 12);//12
            Spara_Header.Add("Input Port", 13);     //13
            Spara_Header.Add("Output Port", 14);    //14
            Spara_Header.Add("Port_Define", 15);    //15
            Spara_Header.Add("S-Parameter", 16);    //16
            Spara_Header.Add("DM_S-Param", 17);    //17
            Spara_Header.Add("Channel Number", 18); //18
            Spara_Header.Add("Start_Freq", 19);     //19
            Spara_Header.Add("Stop_Freq", 20);      //20
            Spara_Header.Add("Search_Method", 21);  //21

            Spara_Header.Add("MIN_Limit", 22);  //22
            Spara_Header.Add("TYP_Limit", 23);  //23
            Spara_Header.Add("MAX_Limit", 24);  //24

            Spara_Header.Add("Use_Previous", 25);   //25
            Spara_Header.Add("Set_Temp", 26);       //26
            Spara_Header.Add("MIPI_Description", 27);       //27
            Spara_Header.Add("Convert_SIGN_FOR_ISO", 28);   //28

            //---------------RF config information---------------//

            Spara_Header.Add("TRX_ON", 29);                 //29
            Spara_Header.Add("TX_MODE", 30);                //30
            Spara_Header.Add("TX_BAND", 31);                //31
            Spara_Header.Add("TX_INPUT", 32);               //32
            Spara_Header.Add("TX_OUTPUT", 33);              //33
            Spara_Header.Add("PA_MODE", 34);                //34
            Spara_Header.Add("RX_BAND", 35);                //35
            Spara_Header.Add("RX_INPUT", 36);               //36
            Spara_Header.Add("RX_OUTPUT", 37);              //37
            Spara_Header.Add("LNA_MODE", 38);               //38
            Spara_Header.Add("TDD_PRIORITY", 39);           //39

            Spara_Header.Add("ASM_ANT1", 40);               //40
            Spara_Header.Add("ASM_ANT2", 41);               //41
            Spara_Header.Add("ASM_UAT", 42);                //42

            //---------------Actual Mipi Value---------------//

            int Dynamic_Col = Spara_Header.Count + 1; //initial index for dynamic range

            foreach (string MIPI_addr in GetMIPI_AddrVal())
            {
                Spara_Header.Add(MIPI_addr, Dynamic_Col);
                Dynamic_Col++;
            }

            Spara_Header.Add("VCC_V", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("VBAT_V", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("LNAVDD_V", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("VCC_I", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("VBAT_I", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("LNAVDD_I", Dynamic_Col); Dynamic_Col++;

            Spara_Header.Add("ParameterNote", Dynamic_Col); Dynamic_Col++;
            Spara_Header.Add("#End", Dynamic_Col);          //Header End

            return Spara_Header;
        }

        public Spara_Path GetPath_from_Trigger(Spara_Trigger_Group test_trigger)
        {
            Spara_Path TRIG_RF_Path = new Spara_Path();
            TestCon Sample_TC = test_trigger.TestCon_List[0];
            bool IsTX = false;
            bool IsASM = false;

            //ASM case has no TX setting
            //Set TRX_ON "Test Direction" TX or RX

            if ((Sample_TC.Direction.ToUpper().Contains("TX") ||
                 Sample_TC.Direction.ToUpper().Contains("TRX")) &&
                !Sample_TC.Band.Trim().ToUpper().Contains("ALL"))
            {
                TRIG_RF_Path.TRX_ON_Direction = "TXRX";
                TRIG_RF_Path.TX_Tech = "LTE";
                TRIG_RF_Path.PA_MODE = "ET_HPM";
                IsTX = true;
            }
            else if (Sample_TC.Direction.ToUpper().Contains("RX") &&
                    !Sample_TC.Direction.ToUpper().Contains("TRX"))
            {
                TRIG_RF_Path.TRX_ON_Direction = "RX";
                TRIG_RF_Path.TX_Tech = "LTE";
                TRIG_RF_Path.PA_MODE = "ET_HPM";
                IsTX = false;
            }
            else
            {
                IsASM = true;
            }
            
            if (IsTX && !IsASM) //TX 
            {
                bool find_TXIN = false;
                bool find_TXOUT = false;
                bool Is_LMB = false;
                string Is_HighBand = "HB";

                bool B11_B21_ANT_Exception = false;
                

                // Set Band Name
                if (!Sample_TC.Band.ToUpper().Contains("ALL") && Sample_TC.Band != "")
                {
                    string prefix_Band = "B";
                    if (Sample_TC.Band.ToUpper().Contains("N")) prefix_Band = "";  //for NR
                    if (Sample_TC.Band.Contains('.')) Sample_TC.Band = Sample_TC.Band.Replace('.', 'P');

                    StringBuilder Band_Name = new StringBuilder();
                    Band_Name.AppendFormat("{0}{1}_", prefix_Band, Sample_TC.Band.Trim().ToUpper());
                    TRIG_RF_Path.TX_BAND = Band_Name.ToString();
                }

                bool IsTDD = Get_IsTDD(Sample_TC.Band);
                if (IsTDD) TRIG_RF_Path.TDD_Priority = "TX";

                // Set TX Input, TX output
                foreach (TestCon item in test_trigger.TestCon_List)
                {
                    if (item.Parameter == "ISO:TX, ASM") IsTDD = IsTDD;

                    switch (item.Parameter)
                    {
                        case "Input_RL":
                        case "Gain_Ripple":
                        case "TX_OOB_Gain":
                            TRIG_RF_Path.TX_INPUT = item.Input_Port;
                            TRIG_RF_Path.TX_ANT = item.Output_Port;
                            find_TXIN = true; find_TXOUT = true;
                            break;
                        case "ISO:ANT, InAct_ANT":
                        case "ISO:ANT, ANT":
                        case "Output_RL":
                            if (Globals.Spara_config_INFO.Dic_PortDefinition[test_trigger.Test_Input] == "TX_INPUT")
                            {
                                TRIG_RF_Path.TX_INPUT = test_trigger.Test_Input;
                                find_TXIN = true;
                            }
                            else
                            {
                                foreach (var Band_Freq in Globals.IniFile.Frequency_table.Keys)
                                {
                                    if (Band_Freq.Contains("B" + Sample_TC.Band))
                                    {
                                        string[] temp = Globals.IniFile.Frequency_table[Band_Freq].Split(',');
                                        if (Convert.ToSingle(temp[1].Trim()) < 1700f) 
                                        {
                                            Is_HighBand = "LMB";
                                            break;
                                        }
                                        else if (Convert.ToSingle(temp[1].Trim()) < 2300f) //stop frequency is smaller than 2300MHz
                                        {
                                            Is_HighBand = "MB";
                                            break;
                                        }
                                        else if (Convert.ToSingle(temp[1].Trim()) > 2300f) //stop frequency is smaller than 2300MHz
                                        {
                                            Is_HighBand = "HB";
                                            break;
                                        }
                                    }
                                }

                                foreach (var DefinedPorts in Globals.Spara_config_INFO.Dic_PortDefinition)
                                {
                                    if (DefinedPorts.Key.ToUpper().Contains(Is_HighBand) && DefinedPorts.Value.Contains("TX_INPUT"))
                                    {
                                        TRIG_RF_Path.TX_INPUT = DefinedPorts.Key;
                                        find_TXIN = true;
                                        break;
                                    }
                                }
                            }

                            if (item.Parameter == "Output_RL")
                            {
                                TRIG_RF_Path.TX_ANT = item.Output_Port;
                                find_TXOUT = true;
                            }
                            else
                            {
                                TRIG_RF_Path.TX_ANT = item.Input_Port;
                                find_TXOUT = true;
                            }
                            break;
                        case "ISO:TX, RX":
                        case "ISO:RX, InAct_RX":
                        case "ISO:InAct_RX, InAct_RX":
                        case "ISO:TX, InAct_RX":
                        case "ISO:TX, ASM":

                            if (Globals.Spara_config_INFO.Dic_PortDefinition[item.Input_Port] == "TX_INPUT")
                            {
                                TRIG_RF_Path.TX_INPUT = item.Input_Port;
                                find_TXIN = true;
                            }
                            else
                            {
                                foreach (var Band_Freq in Globals.IniFile.Frequency_table.Keys)
                                {
                                    if (Band_Freq.Contains("B" + Sample_TC.Band))
                                    {
                                        string[] temp = Globals.IniFile.Frequency_table[Band_Freq].Split(',');
                                        if (Convert.ToSingle(temp[1].Trim()) < 1700f)
                                        {
                                            Is_HighBand = "LMB";
                                            break;
                                        }
                                        else if (Convert.ToSingle(temp[1].Trim()) < 2300f) //stop frequency is smaller than 2300MHz
                                        {
                                            Is_HighBand = "MB";
                                            break;
                                        }
                                        else if (Convert.ToSingle(temp[1].Trim()) > 2300f) //stop frequency is smaller than 2300MHz
                                        {
                                            Is_HighBand = "HB";
                                            break;
                                        }
                                    }
                                }

                                foreach (var DefinedPorts in Globals.Spara_config_INFO.Dic_PortDefinition)
                                {
                                    if (DefinedPorts.Key.ToUpper().Contains(Is_HighBand) && DefinedPorts.Value.Contains("TX_INPUT"))
                                    {
                                        TRIG_RF_Path.TX_INPUT = DefinedPorts.Key;
                                        find_TXIN = true;
                                        break;
                                    }
                                }
                            }
                            
                            if (Globals.Spara_config_INFO.Dic_PortDefinition[test_trigger.Test_Output] == "ANT_OUT")
                            {
                                TRIG_RF_Path.TX_ANT = test_trigger.Test_Output;
                                find_TXOUT = true;
                            }
                            else
                            {
                                if (TRIG_RF_Path.TX_BAND.Contains("B11_") || TRIG_RF_Path.TX_BAND.Contains("B21_"))
                                {
                                    if (Globals.Default_ANT == "ANT2")
                                    {
                                        TRIG_RF_Path.TX_ANT = "ANT1"; //Exceptional for B11,B21 on Cheddar, these bands not use "ANT2"
                                    }
                                    else
                                    {
                                        TRIG_RF_Path.TX_ANT = Globals.Default_ANT; //General condition 
                                    }
                                }
                                else
                                {
                                    TRIG_RF_Path.TX_ANT = Globals.Default_ANT; //General condition 
                                }
                                find_TXOUT = true;
                            }

                            if (Globals.Spara_config_INFO.Dic_PortDefinition[test_trigger.Test_Output] == "RX_OUT")  //Need verification (B11, B21 differenct RX Antenna case) 
                            {
                                if (TRIG_RF_Path.TX_BAND.Contains("B11_") || TRIG_RF_Path.TX_BAND.Contains("B21_"))
                                {
                                    foreach (string Antenna_name in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                                    {
                                        if (Globals.Spara_config_INFO.Dic_PortDefinition[Antenna_name] == "ANT_OUT")
                                        {
                                            if (test_trigger.CA_Case.Contains("OUT"))
                                            {
                                                if (test_trigger.CA_Case.Contains("B11")|| test_trigger.CA_Case.Contains("B21"))
                                                {
                                                    if (Antenna_name == TRIG_RF_Path.TX_ANT)
                                                    {
                                                        TRIG_RF_Path.RX_ANT = Antenna_name;
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    if (Antenna_name != TRIG_RF_Path.TX_ANT)
                                                    {
                                                        TRIG_RF_Path.RX_ANT = Antenna_name;
                                                        break;
                                                    }
                                                }
                                            }                                      
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }

                    if (find_TXIN && find_TXOUT) break;
                }

                //Find RX OUT, ANT port (if need)

                if(test_trigger.CA_Case.Contains("OUT"))
                {
                    string[] CA_RX_OUTS = test_trigger.CA_Case.Split('.');

                    StringBuilder RXBands = new StringBuilder();
                    StringBuilder RXOUT = new StringBuilder();

                    foreach (string Each_Ports in CA_RX_OUTS)
                    {
                        int index_split = Each_Ports.IndexOf('O'); //find text "OUT"
                        string Band = Each_Ports.Substring(0, (index_split - 0)).Trim();
                        string RX_OUT = Each_Ports.Substring(index_split, (Each_Ports.Length - index_split)).Trim();

                        if (RXBands.Length == 0){ RXBands.AppendFormat("{0}_", Band); }
                        else { RXBands.AppendFormat("+{0}_", Band); }

                        if (RXOUT.Length == 0){ RXOUT.AppendFormat("{0}", RX_OUT); }
                        else { RXOUT.AppendFormat("+{0}", RX_OUT); }
                    }

                    TRIG_RF_Path.RX_CA_BAND = RXBands.ToString();
                    TRIG_RF_Path.RX_OUT = RXOUT.ToString();
                    if (TRIG_RF_Path.RX_ANT == "") TRIG_RF_Path.RX_ANT = TRIG_RF_Path.TX_ANT;
                }

                //Find ASM & ANT out revise

                List<string> ASMs = new List<string>();
                ASMs.Add(test_trigger.ASM1); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM2); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM3); //Antenn1 connection with ASM

                List<string> Antennas = new List<string>();
                foreach (string Antenna_name in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                {
                    if(Globals.Spara_config_INFO.Dic_PortDefinition[Antenna_name]=="ANT_OUT")
                    {
                        Antennas.Add(Antenna_name);
                    }
                }

                int Current_ANT_index = 0;
                for (int i = 0; i < Antennas.Count; i++)
                {
                    if (Antennas[i] == TRIG_RF_Path.TX_ANT) Current_ANT_index = i;
                }

                if (ASMs[Current_ANT_index] == "TERM")
                {
                    ASMs[Current_ANT_index] = "";
                }
                else
                {
                    string temp = ASMs[Current_ANT_index];

                    for (int i = 0; i < ASMs.Count; i++)
                    {
                        if (i == Current_ANT_index) continue;
                        if (ASMs[i] == "TERM") ASMs[i] = temp;
                    }

                    ASMs[Current_ANT_index] = "";

                    if (test_trigger.TestCon_List[0].Parameter == "ISO:TX, ASM")
                    {
                        string temp_duplicate_remover = ASMs[0];
                        for (int i = 0; i < ASMs.Count; i++)
                        {
                            if (temp_duplicate_remover == ASMs[i] && temp_duplicate_remover != "TERM")
                            {
                                temp_duplicate_remover = ASMs[i];
                                ASMs[i] = "TERM";
                            }
                            else
                            {
                                temp_duplicate_remover = ASMs[i];
                            }

                            if (ASMs[Current_ANT_index] == "TERM") ASMs[Current_ANT_index] = "";
                        }
                    }
                }

                int Secondary_ANT_index = 0;

                if (test_trigger.TestCon_List[0].Parameter == "ISO:ANT, ANT")
                {
                    for (int i = 0; i < Antennas.Count; i++)
                    {
                        if (Antennas[i] == test_trigger.Test_Output) Secondary_ANT_index = i;
                    }

                    if (ASMs[Secondary_ANT_index] == "TERM")
                    {
                        ASMs[Secondary_ANT_index] = "MIMO"; //Set temporary ANT
                    }
                }

                for (int i = 0; i < Antennas.Count; i++)
                {
                    TRIG_RF_Path.ASM.Add(Antennas[i], ASMs[i]); 
                }

                if(TRIG_RF_Path.RX_OUT != "") TRIG_RF_Path.LNA_GAIN = Globals.IniFile.Default_ISO_gainMode; //TX case set all gain as default RX Gain as "G1"

            }
            else if(!IsTX && !IsASM)//RX 
            {
                //Find default TX input (not used but only described)
                string Is_HighBand = "HB";
                if (Sample_TC.Band.Contains('.')) Sample_TC.Band = Sample_TC.Band.Replace('.', 'P');

                foreach (var Band_Freq in Globals.IniFile.Frequency_table.Keys)
                {
                    if (Band_Freq.Contains("B" + Sample_TC.Band))
                    {
                        string[] temp = Globals.IniFile.Frequency_table[Band_Freq].Split(',');
                        if (Convert.ToSingle(temp[1].Trim()) < 1700f)
                        {
                            Is_HighBand = "LMB";
                            break;
                        }
                        else if (Convert.ToSingle(temp[1].Trim()) < 2300f) //stop frequency is smaller than 2300MHz
                        {
                            Is_HighBand = "MB";
                            break;
                        }
                        else if (Convert.ToSingle(temp[1].Trim()) > 2300f) //stop frequency is smaller than 2300MHz
                        {
                            Is_HighBand = "HB";
                            break;
                        }
                    }
                }

                foreach (var DefinedPorts in Globals.Spara_config_INFO.Dic_PortDefinition)
                {
                    if (DefinedPorts.Key.ToUpper().Contains(Is_HighBand) && DefinedPorts.Value.Contains("TX_INPUT"))
                    {
                        TRIG_RF_Path.TX_INPUT = DefinedPorts.Key;
                        break;
                    }
                }

                // Set Band Name
                if (!Sample_TC.Band.ToUpper().Contains("ALL") && Sample_TC.Band != "")
                {
                    string prefix_Band = "B";
                    if (Sample_TC.Band.ToUpper().Contains("N")) prefix_Band = "";  //for NR

                    StringBuilder Band_Name = new StringBuilder();
                    Band_Name.AppendFormat("{0}{1}_", prefix_Band, Sample_TC.Band.Trim().ToUpper());

                    string C_Band_Name = Band_Name.ToString();
                    if (C_Band_Name.Contains('.')) { C_Band_Name = C_Band_Name.Replace('.', 'P'); }
                    TRIG_RF_Path.TX_BAND = C_Band_Name;
                }

                bool IsTDD = Get_IsTDD(Sample_TC.Band);
                if (IsTDD) TRIG_RF_Path.TDD_Priority = "RX";

                TRIG_RF_Path.LNA_GAIN = test_trigger.RX_Mode;


                //Set RX Band name & RX OUT
                if (test_trigger.CA_Case.Contains("OUT"))
                {
                    string[] CA_RX_OUTS = test_trigger.CA_Case.Split('.');

                    StringBuilder RXBands = new StringBuilder();
                    StringBuilder RXOUT = new StringBuilder();

                    foreach (string Each_Ports in CA_RX_OUTS)
                    {
                        int index_split = Each_Ports.IndexOf('O'); //find text "OUT"
                        string Band = Each_Ports.Substring(0, (index_split - 0)).Trim();
                        string RX_OUT = Each_Ports.Substring(index_split, (Each_Ports.Length - index_split)).Trim();

                        if (RXBands.Length == 0) { RXBands.AppendFormat("{0}_", Band); }
                        else { RXBands.AppendFormat("+{0}_", Band); }

                        if (RXOUT.Length == 0) { RXOUT.AppendFormat("{0}", RX_OUT); }
                        else { RXOUT.AppendFormat("+{0}", RX_OUT); }
                    }

                    TRIG_RF_Path.RX_CA_BAND = RXBands.ToString();
                    TRIG_RF_Path.RX_OUT = RXOUT.ToString();
                }

                //TRIG_RF_Path.RX_ANT = TRIG_RF_Path.TX_ANT;
                bool find_ANT = false;

                foreach (TestCon item in test_trigger.TestCon_List)
                {
                    string Test_ID = item.Parameter;
                    if (Test_ID.Contains("RX_Gain_G")) Test_ID = "RX_Gain_STD";
                    if (Test_ID.Contains("RX_Gain_CA")) Test_ID = "RX_Gain_CA";

                    switch (Test_ID)
                    {
                        case "Gain_Ripple":
                        case "RX_OOB_Gain":
                        case "K_factor":
                        case "MU_factor":
                        case "Group_Delay":
                        case "Phase_Delta":
                        case "RX_Gain_STD":
                            TRIG_RF_Path.RX_ANT = item.Input_Port;
                            TRIG_RF_Path.TX_ANT = item.Input_Port;
                            find_ANT = true;
                            break;
                        case "REV_ISO:RX, ANT":
                            TRIG_RF_Path.RX_ANT = item.Output_Port;
                            TRIG_RF_Path.TX_ANT = item.Output_Port;
                            find_ANT = true;
                            break;
                        case "ISO:ANT, ANT":
                            TRIG_RF_Path.RX_ANT = test_trigger.Test_Input;
                            TRIG_RF_Path.TX_ANT = test_trigger.Test_Input;
                            find_ANT = true;
                            break;
                        case "RX_Gain_CA":
                            TRIG_RF_Path.TX_ANT = item.Input_Port;

                            if ((TRIG_RF_Path.RX_CA_BAND.Contains("B1P6G") || TRIG_RF_Path.RX_CA_BAND.Contains("B11") || TRIG_RF_Path.RX_CA_BAND.Contains("B21")))
                            {
                                foreach (string Antenna_name in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                                {
                                    if (Globals.Spara_config_INFO.Dic_PortDefinition[Antenna_name] == "ANT_OUT")
                                    {
                                        if (Antenna_name != TRIG_RF_Path.TX_ANT)
                                        {
                                            TRIG_RF_Path.RX_ANT = Antenna_name;
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                TRIG_RF_Path.RX_ANT = item.Input_Port;
                            }

                            find_ANT = true;
                            break;

                        default:
                            break;
                    }

                    if (find_ANT == true) break;
                }

                //Find ASM & ANT out revise

                List<string> ASMs = new List<string>();
                ASMs.Add(test_trigger.ASM1); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM2); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM3); //Antenn1 connection with ASM

                List<string> Antennas = new List<string>();
                foreach (string Antenna_name in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                {
                    if (Globals.Spara_config_INFO.Dic_PortDefinition[Antenna_name] == "ANT_OUT")
                    {
                        Antennas.Add(Antenna_name);
                    }
                }

                int Current_ANT_index = 0;
                for (int i = 0; i < Antennas.Count; i++)
                {
                    if (Antennas[i] == TRIG_RF_Path.TX_ANT) Current_ANT_index = i;
                }

                if (ASMs[Current_ANT_index] == "TERM")
                {
                    ASMs[Current_ANT_index] = "";
                }
                else
                {
                    string temp = ASMs[Current_ANT_index];

                    for (int i = 0; i < ASMs.Count; i++)
                    {
                        if (i == Current_ANT_index) continue;
                        if (ASMs[i] == "TERM") ASMs[i] = temp;
                    }

                    ASMs[Current_ANT_index] = "";
                }

                int Secondary_ANT_index = 0;

                if (test_trigger.TestCon_List[0].Parameter == "ISO:ANT, ANT")
                {
                    for (int i = 0; i < Antennas.Count; i++)
                    {
                        if (Antennas[i] == test_trigger.Test_Output) Secondary_ANT_index = i;
                    }

                    if (ASMs[Secondary_ANT_index] == "TERM")
                    {
                        ASMs[Secondary_ANT_index] = "MIMO"; //Set temporary ANT
                    }
                }

                for (int i = 0; i < Antennas.Count; i++)
                {
                    TRIG_RF_Path.ASM.Add(Antennas[i], ASMs[i]);
                }

            }
            else if(IsASM)
            {
                List<string> ASMs = new List<string>();
                ASMs.Add(test_trigger.ASM1); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM2); //Antenn1 connection with ASM
                ASMs.Add(test_trigger.ASM3); //Antenn1 connection with ASM

                List<string> Antennas = new List<string>();
                foreach (string Antenna_name in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                {
                    if (Globals.Spara_config_INFO.Dic_PortDefinition[Antenna_name] == "ANT_OUT")
                    {
                        Antennas.Add(Antenna_name);
                    }
                }

                for (int i = 0; i < Antennas.Count; i++)
                {
                    TRIG_RF_Path.ASM.Add(Antennas[i], ASMs[i]);
                }
            }

            return TRIG_RF_Path;
        }

        private List<string> GetMIPI_AddrVal()
        {
            //will revise fetch mipi address value from external source > return List<string>

            List<string> MipiAddress = new List<string>();

            MipiAddress.Add("TXREG00");
            MipiAddress.Add("TXREG03");
            MipiAddress.Add("TXREG04");
            MipiAddress.Add("TXREG05");
            MipiAddress.Add("TXREG06");
            MipiAddress.Add("TXREG07");
            MipiAddress.Add("TXREG08");
            MipiAddress.Add("TXREG0B");
            MipiAddress.Add("TXREG0C");
            MipiAddress.Add("TXREG0D");

            MipiAddress.Add("RXREG00");
            MipiAddress.Add("RXREG01");
            MipiAddress.Add("RXREG02");
            MipiAddress.Add("RXREG03");
            MipiAddress.Add("RXREG04");
            MipiAddress.Add("RXREG0B");
            MipiAddress.Add("RXREG0D");
            MipiAddress.Add("RXREG0F");
            MipiAddress.Add("RXREG11");
            MipiAddress.Add("RXREG13");
            MipiAddress.Add("REGCUSTOM");

            return MipiAddress;
        }
        private int GetVCCslot(Dictionary<string, int> Header, string target)
        {
            int find_index = -99;

            try
            {
                find_index = Header[target];
            }
            catch
            {
                foreach (string Header_name in Header.Keys)
                {
                    if (Header_name.ToUpper().Contains(target.ToUpper()))
                    {
                        find_index = Header[Header_name];
                    }
                }

            }

            return find_index;
        }

    }

    class Spara_Path
    {
        public string TRX_ON_Direction;

        public string TX_BAND;
        public string TX_INPUT;
        public string TX_ANT;

        public string TX_Tech;
        public string PA_MODE;

        public string RX_CA_BAND;

        public string RX_ANT;
        public string RX_OUT;

        public string LNA_GAIN;
        public string TDD_Priority;

        public Dictionary<string, string> ASM;

        public Spara_Path()
        {
            clear();
        }

        public void clear()
        {
            this.TRX_ON_Direction = "";
            this.TX_BAND = "";
            this.TX_INPUT = "";
            this.TX_ANT = "";
            this.RX_ANT = "";
            this.RX_OUT = "";
            this.LNA_GAIN = "";
            this.TDD_Priority = "";
            this.RX_CA_BAND = "";
            this.ASM = new Dictionary<string, string>();

            this.TX_Tech = "";
            this.PA_MODE = "";
        }

    }
}
