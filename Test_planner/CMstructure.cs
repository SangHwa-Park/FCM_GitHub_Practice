using System;
using System.Collections.Generic;
using System.IO; //FILE IO as like as "stream reader"
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Base;
using FlexTestLib.MsgBox;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;

namespace Test_Planner
{
    public class CMstructure
    {
        public void Revise_MemoryCM_Full(string current_sheet, ref Dictionary<string, Excel_Base.Band_Condition> CM_Sheet)
        {
            int row_range = CM_Sheet[current_sheet].Test_SpecID.Count;
            string Test_name_adjust = "";
            List<int> Row_Index = new List<int>();

            for (int i = 0; i < row_range; i++)
            {
                Row_Index.Add(i); //add row number as Key index per each band sheet condition
                //Due to merged Cell, need "Test Name param align" using last valid test name "Test_name_adjust"
                if (CM_Sheet[current_sheet].Test_SpecID[i].Trim() != "" && CM_Sheet[current_sheet].Test_Name[i].Trim() != "")
                {
                    Test_name_adjust = CM_Sheet[current_sheet].Test_Name[i];
                }
                if (CM_Sheet[current_sheet].Test_SpecID[i].Trim() != "" && CM_Sheet[current_sheet].Test_Name[i].Trim() == "") //Spec ID exist but there are no test name = means merged cell
                {
                    CM_Sheet[current_sheet].Test_Name[i] = Test_name_adjust;
                }

                //Allocated matched freuqency value to empty null "Frequency start"&"stop" cell
                if (IsEmpty(CM_Sheet[current_sheet].Start_Freq, i) &&
                    IsEmpty(CM_Sheet[current_sheet].Stop_Freq, i) &&
                    (CM_Sheet[current_sheet].Input_Port[i].Trim() != "" ||
                    CM_Sheet[current_sheet].Output_Port[i].Trim() != ""))
                {
                    string start_Frequency = CM_Sheet[current_sheet].Start_Freq[i];
                    string stop_Frequency = CM_Sheet[current_sheet].Stop_Freq[i];

                    Matching_TXRX_Frequency(CM_Sheet[current_sheet].Test_SpecID[i], ref start_Frequency, ref stop_Frequency);

                    CM_Sheet[current_sheet].Start_Freq[i] = start_Frequency;
                    CM_Sheet[current_sheet].Stop_Freq[i] = stop_Frequency;
                }
                
                if (CM_Sheet[current_sheet].Input_Port[i].Contains("PRX_OUT1,2,3,4"))  //port description exception
                {
                    CM_Sheet[current_sheet].Input_Port[i] = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
                }
                else if (CM_Sheet[current_sheet].Output_Port[i].Contains("PRX_OUT1,2,3,4"))
                {
                    CM_Sheet[current_sheet].Output_Port[i] = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
                }

                string param_ID = "";
                string Test_Direction = "";

                try
                {
                    if (i == 207) param_ID = "";

                    if (!(CM_Sheet[current_sheet].Input_Port[i].Trim() == "" &&
                            CM_Sheet[current_sheet].Output_Port[i].Trim() == "")) // true when both input and output port is not null
                    {
                        if (Test_Param_definition(CM_Sheet[current_sheet], i, ref param_ID, ref Test_Direction, current_sheet))
                        {
                            CM_Sheet[current_sheet].Parameter[i] = param_ID;
                            CM_Sheet[current_sheet].Direction[i] = Test_Direction;
                        }
                    }
                    else if(CM_Sheet[current_sheet].Test_Name[i].ToUpper().Contains("CURRENT")) //exception when IDD test parameter has no port information
                    {
                        if (Test_Param_definition(CM_Sheet[current_sheet], i, ref param_ID, ref Test_Direction, current_sheet))
                        {
                            CM_Sheet[current_sheet].Parameter[i] = param_ID;
                            CM_Sheet[current_sheet].Direction[i] = Test_Direction;
                        }
                    }
                }
                catch (Exception)
                {
                    StringBuilder ErrMsg = new StringBuilder();
                    ErrMsg.AppendFormat("Error: during Test define process at Band \"{0}\"\n Test param = {1}, index = {2}", CM_Sheet[current_sheet].Band[i], CM_Sheet[current_sheet].Test_Name[i], i);
                    ErrMsg.AppendFormat("\n Test SpecID = {0}", CM_Sheet[current_sheet].Test_SpecID[i]);
                    ErrMsg.AppendFormat("\n");
                    ErrMsg.AppendFormat("\nNeed to debugging \"Test_Param_definition\" at Revise_MemoryCM_Full function");
                    ClsMsgBox.Show("Error on revise test definition after CM loading", ErrMsg.ToString());
                    Environment.Exit(0);
                    throw;
                }

                

            }

            CM_Sheet[current_sheet].Key_Index = Row_Index;
        }

        private bool IsEmpty(List<string> Array, int index)
        {
            bool Is_Empty = false;

            if (Array[index].Trim().ToUpper() == "") Is_Empty = true;
            if (Array[index].Trim().ToUpper() == "-") Is_Empty = true;
            if (Array[index].Trim().ToUpper() == " ") Is_Empty = true;

            return Is_Empty;
        }

        private bool IsEmpty(string Array)
        {
            bool Is_Empty = false;

            if (Array.Trim().ToUpper() == "") Is_Empty = true;
            if (Array.Trim().ToUpper() == "-") Is_Empty = true;
            if (Array.Trim().ToUpper() == " ") Is_Empty = true;

            return Is_Empty;
        }



        public void Matching_TXRX_Frequency(string Spec_ID, ref string start_F, ref string stop_F)
        {
            string[] SpecID_split = Spec_ID.Trim().ToUpper().Split('_');
            string Band_ID = "";
            string Signal_Direction = "";

            if (IsEmpty(start_F)|| IsEmpty(stop_F))
            {
                if (SpecID_split[0].Contains(".")) SpecID_split[0] = SpecID_split[0].Replace('.', 'P');
                if (!SpecID_split[0].ToUpper().Contains("B")&& !SpecID_split[0].ToUpper().Contains("N")) SpecID_split[0] = "B" + SpecID_split[0];

                foreach (string Item in SpecID_split)
                {
                    if (Item.Trim().ToUpper().Contains("B")|| Item.Trim().ToUpper().Contains("N")) Band_ID = Item.Trim().ToUpper();
                    if (Item.Trim().ToUpper().Contains("TX")) Signal_Direction = "TX";
                    if (Item.Trim().ToUpper().Contains("RX") && !Item.Trim().ToUpper().Contains("TRX")) Signal_Direction = "RX";
                    if (Item.Trim().ToUpper().Contains("TRX")) Signal_Direction = "TRX";
                }

                if(Band_ID.Contains("B40") && !Band_ID.Contains("B40A"))
                {
                    Band_ID = "B40F";
                }

                if (Band_ID != "" || Signal_Direction != "") //this condition means it has spec ID, but no frequency 
                {
                    foreach (string key in Globals.IniFile.Frequency_table.Keys)
                    {
                        string[] Compare_key = key.Split('_');
                        if (Compare_key[2].Trim().ToUpper() == Band_ID.Trim().ToUpper() && Compare_key[1].Trim().ToUpper() == Signal_Direction.Trim().ToUpper())
                        {
                            string[] start_stop_freq = Globals.IniFile.Frequency_table[key].Trim().Split(',');
                            start_F = start_stop_freq[0].Trim();
                            stop_F = start_stop_freq[1].Trim();
                            break;
                        }
                    }

                }
            }
        }

        public string Find_TestDirection(string Spec_ID)
        {
            string default_dirction = "";

            if (Spec_ID.Trim().ToUpper().Contains("TRX"))
            {
                default_dirction = "TRX";
            }
            else if (Spec_ID.Trim().ToUpper().Contains("TX"))
            {
                default_dirction = "TX";
            }
            else if (Spec_ID.Trim().ToUpper().Contains("RX"))
            {
                default_dirction = "RX";
            }

            return default_dirction;
        }

        public string Find_Mode(string Test_mode)
        {
            string default_mode = "";

            if (Test_mode.Trim().ToUpper().Contains("ET") && Test_mode.Trim().ToUpper().Contains("APT"))
            {
                default_mode = "ETnAPT";
            }
            else if (Test_mode.Trim().ToUpper().Contains("ET"))
            {
                default_mode = "ET";
            }
            else if (Test_mode.Trim().ToUpper().Contains("APT"))
            {
                default_mode = "APT";
            }
            else if (Test_mode.Trim().ToUpper().Contains("G"))
            {
                Test_mode = Test_mode.Replace('[', ' ');
                Test_mode = Test_mode.Replace(']', ' ');
                Test_mode = Test_mode.Replace('{', ' ');
                Test_mode = Test_mode.Replace('}', ' ');
                Test_mode = Test_mode.Trim().ToUpper();
                default_mode = Test_mode;
            }

            return default_mode;
        }

        public string Find_Harmonic_Step(string Test_name)
        {
            string default_step = "HAR";

            if (Test_name.Trim().ToUpper().Contains("2ND"))
            {
                default_step = "HAR_2";
            }
            else if (Test_name.Trim().ToUpper().Contains("3RD"))
            {
                default_step = "HAR_3";
            }
            else if (Test_name.Trim().ToUpper().Contains("4TH"))
            {
                default_step = "HAR_4";
            }
            else if (Test_name.Trim().ToUpper().Contains("5TH"))
            {
                default_step = "HAR_5";
            }
            return default_step;
        }

        public bool ExceptionCASE(string testname)
        {
            bool IsExceptionCase = false;
            string Exception_case = testname.Trim();

            switch (Exception_case)
            {
                case "Noise Transfer from VCC to RF Output":
                    IsExceptionCase = true;
                    break;

                default:
                    break;
            }
            
            return IsExceptionCase;
        }

        public bool InBand_Frequency(string Start_Freq, string stop_Freq, ref string Band_found)
        {
            //Globals.IniFile.Frequency_table

            foreach (string band_name in Globals.IniFile.Frequency_table.Keys)
            {
                string[] Table_string = band_name.Split('_');
                string Table_band = Table_string[2].Trim().ToUpper();
                if (Table_band == Band_found)
                {
                    string Freq_range = Globals.IniFile.Frequency_table[band_name];
                    string[] Frequencys = Freq_range.Split(',');
                    List<float> float_freq = new List<float>();
                    foreach (string freq in Frequencys)
                    {
                        float_freq.Add(Convert.ToSingle(freq.Trim()));
                    }
                    float_freq.Sort();

                    //if current frequency is in range of defined band frequency
                    if(Convert.ToSingle(Start_Freq.Trim())>=float_freq[0] && Convert.ToSingle(stop_Freq.Trim())<=float_freq[1])
                    {
                        Band_found = band_name;
                        return true;
                    }

                }
            }

            return false;
        }

        public string search_ISO_Port(string port_description, List<string> port_list)
        {
            string converted_port_description = "";

            foreach (string Port in port_list)
            {
                if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "TX_INPUT"))
                {
                    if (port_description.Contains("INACTIVE"))
                    {
                        converted_port_description = "InAct_TX";
                    }
                    else //Means Active or none description
                    {
                        converted_port_description = "TX";
                    }
                }
                else if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ANT_OUT"))
                {
                    if (port_description.Contains("INACTIVE"))
                    {
                        converted_port_description = "InAct_ANT";
                    }
                    else //Means Active or none description
                    {
                        converted_port_description = "ANT";
                    }
                }
                else if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "RX_OUT"))
                {
                    if (port_description.Contains("INACTIVE"))
                    {
                        converted_port_description = "InAct_RX";
                    }
                    else //Means Active or none description
                    {
                        converted_port_description = "RX";
                    }
                }
                else if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ASM"))
                {
                    if (port_description.Contains("INACTIVE"))
                    {
                        converted_port_description = "InAct_ASM";
                    }
                    else //Means Active or none description
                    {
                        converted_port_description = "ASM";
                    }
                }
                else if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "B11B21_IN"))
                {
                    if (port_description.Contains("INACTIVE"))
                    {
                        converted_port_description = "InAct_B11B21_IN";
                    }
                    else //Means Active or none description
                    {
                        converted_port_description = "B11B21_IN";
                    }
                }
            }

            return converted_port_description;
        }

        public void Spara_defined(string Param_ID, string Band, int index)
        {
            string Band_Index = Band.Trim() + "," + Convert.ToString(index);

            if(Globals.Spara_TestDic.ContainsKey(Param_ID))
            {
                Globals.Spara_TestDic[Param_ID].Add(Band_Index);
            }
            else
            {
                List<string> new_param_list = new List<string>();
                new_param_list.Add(Band_Index);
                Globals.Spara_TestDic.Add(Param_ID, new_param_list);
            }
        }

        public void TX_defined(string Param_ID, string Band, int index)
        {
            string Band_Index = Band.Trim() + "," + Convert.ToString(index);

            if (Globals.TX_TestDic.ContainsKey(Param_ID))
            {
                Globals.TX_TestDic[Param_ID].Add(Band_Index);
            }
            else
            {
                List<string> new_param_list = new List<string>();
                new_param_list.Add(Band_Index);
                Globals.TX_TestDic.Add(Param_ID, new_param_list);
            }
        }

        public void RX_defined(string Param_ID, string Band, int index)
        {
            string Band_Index = Band.Trim() + "," + Convert.ToString(index);

            if (Globals.RX_TestDic.ContainsKey(Param_ID))
            {
                Globals.RX_TestDic[Param_ID].Add(Band_Index);
            }
            else
            {
                List<string> new_param_list = new List<string>();
                new_param_list.Add(Band_Index);
                Globals.RX_TestDic.Add(Param_ID, new_param_list);
            }
        }
        public void DC_defined(string Param_ID, string Band, int index)
        {
            string Band_Index = Band.Trim() + "," + Convert.ToString(index);

            if (Globals.DC_TestDic.ContainsKey(Param_ID))
            {
                Globals.DC_TestDic[Param_ID].Add(Band_Index);
            }
            else
            {
                List<string> new_param_list = new List<string>();
                new_param_list.Add(Band_Index);
                Globals.DC_TestDic.Add(Param_ID, new_param_list);
            }
        }
        public void Noise_defined(string Param_ID, string Band, int index)
        {
            string Band_Index = Band.Trim() + "," + Convert.ToString(index);

            if (Globals.NOISE_TestDic.ContainsKey(Param_ID))
            {
                Globals.NOISE_TestDic[Param_ID].Add(Band_Index);
            }
            else
            {
                List<string> new_param_list = new List<string>();
                new_param_list.Add(Band_Index);
                Globals.NOISE_TestDic.Add(Param_ID, new_param_list);
            }
        }

        public bool Test_Param_definition(Excel_Base.Band_Condition CM_Sheet, int index, ref string Param_identifier, ref string Test_Direction, string Sheet_Name)
        {
            string TestName = CM_Sheet.Test_Name[index];
            Test_Direction = "";
            bool IsConverted = false;

            string TestName_Up = TestName.Trim().ToUpper();
            Param_identifier = CM_Sheet.Parameter[index];
            string Param_ID_return = Param_identifier; //in case of no category found
            if (ExceptionCASE(TestName)) return IsConverted;
            //Exception or selection cases

            if (TestName_Up.Contains("INSERTION") && TestName_Up.Contains("LOSS"))
            {
                IsConverted = true;
                Param_identifier = "IL"; //"Insertion Loss" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);

                return IsConverted;
            }

            if (TestName_Up.Contains("RETURN") && TestName_Up.Contains("LOSS"))
            {
                
                if (TestName_Up.Contains("INPUT"))
                {
                    IsConverted = true;
                    Param_identifier = "Input_RL"; //"Input return loss" on Spara
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Spara_defined(Param_identifier, Sheet_Name, index);
                }
                else if (TestName_Up.Contains("OUTPUT"))
                {
                    IsConverted = true;
                    Param_identifier = "Output_RL"; //"Output return loss" on Spara
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Spara_defined(Param_identifier, Sheet_Name, index);
                }
                else
                {
                    List<string> InputPorts = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Input_Port[index]);
                    List<string> OutputPorts = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);

                    string direction = (CM_Sheet.Test_SpecID[index].ToUpper().Contains("TX") ? "TX" : "RX");
                    string Input_Port_Define = "";
                    string Output_Port_Define = "";

                    if (direction == "RX")
                    {
                        Input_Port_Define = "ANT_OUT";
                        Output_Port_Define = "RX_OUT";
                    }
                    else
                    {
                        Input_Port_Define = "TX_INPUT";
                        Output_Port_Define = "ANT_OUT";
                    }

                    foreach (var item in InputPorts)
                    {
                        if(Globals.Spara_config_INFO.Dic_PortDefinition[item].Contains(Input_Port_Define))
                        {
                            IsConverted = true;
                            Param_identifier = "Input_RL"; //"Input return loss" on Spara
                            Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                            Spara_defined(Param_identifier, Sheet_Name, index);
                            return IsConverted;
                        }
                        else if(Globals.Spara_config_INFO.Dic_PortDefinition[item].Contains(Output_Port_Define))
                        {
                            IsConverted = true;
                            Param_identifier = "Output_RL"; //"Input return loss" on Spara
                            Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                            Spara_defined(Param_identifier, Sheet_Name, index);
                            return IsConverted;
                        }
                    }

                }
                return IsConverted;
            }

            if (TestName_Up.Contains("INPUT") && TestName_Up.Contains("VSWR"))
            {
                IsConverted = true;
                Param_identifier = "Input_VSWR"; //"Input return loss" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("PHASE") && TestName_Up.Contains("DELTA"))
            {
                IsConverted = true;
                Param_identifier = "Phase_Delta"; //"Phase Delta variation" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("GROUP") && TestName_Up.Contains("DELAY"))
            {
                IsConverted = true;
                Param_identifier = "Group_Delay"; //"Phase Delta variation" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("FACTOR") && TestName_Up.Contains("K"))
            {
                IsConverted = true;
                Param_identifier = "K_factor"; //"RX K factor" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("FACTOR") && TestName_Up.Contains("MU"))
            {
                IsConverted = true;
                Param_identifier = "MU_factor"; //"RX MU factor" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("RIPPLE"))
            {
                IsConverted = true;
                Param_identifier = "Gain_Ripple"; //"RX Gain ripple" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);

                float check_freq = 0f;
                bool IsString_Start = float.TryParse(CM_Sheet.Start_Freq[index], out check_freq);
                bool IsString_Stop = float.TryParse(CM_Sheet.Stop_Freq[index], out check_freq);

                if (!IsString_Start || !IsString_Stop)
                {
                    string Band_info = Convert.ToString(CM_Sheet.Band[index]);
                    if (Band_info == "" && !Band_info.Contains("B"))
                    {
                        string[] Band = CM_Sheet.Test_SpecID[index].Split('_');
                        Band_info = Band[0].Trim().ToUpper();
                    }
                    else
                    {
                        Band_info = "B" + Band_info;
                    }

                    StringBuilder Frequency_Key = new StringBuilder();
                    Frequency_Key.AppendFormat("FREQ_{0}_{1}", Test_Direction.Trim(), Band_info);
                    string[] Frequency_range = Globals.IniFile.Frequency_table[Frequency_Key.ToString()].Split(',');

                    CM_Sheet.Start_Freq[index] = Frequency_range[0].Trim();
                    CM_Sheet.Stop_Freq[index] = Frequency_range[1].Trim();

                }

                return IsConverted;
            }

            if (TestName_Up.Contains("GAIN") && TestName_Up.Contains("SLOPE")) //"Gain slope" on APT
            {
                IsConverted = true;
                Param_identifier = "Gain_Slope";
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if ((TestName_Up.Contains("GAIN") && TestName_Up.Contains("MODE"))||
                     (TestName_Up.Contains("GAIN") && TestName_Up.Contains("POWER"))) //"APT Mode gain" on APT
            {
                IsConverted = true;
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Param_identifier = "Gain_" + Find_Mode(CM_Sheet.PA_MODE[index]);
                
                if (Test_Direction.Contains("TX"))
                {
                    try
                    {
                        if (Convert.ToSingle(CM_Sheet.Target_Pout[index]) < 10f)
                        {
                            Param_identifier = "Low" + Param_identifier;
                        }
                    }
                    catch
                    {
                        string error_convert_float = "";
                    }
                }

                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if (TestName_Up.Contains("GAIN") && TestName_Up.Contains("RMS")) //"Max power RMS gain" on ET
            {
                IsConverted = true;
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);

                Param_identifier = "Gain_" + Find_Mode(CM_Sheet.PA_MODE[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if (TestName_Up.Contains("GAIN") && TestName_Up.Contains("RIPPLE"))
            {
                IsConverted = true;
                Param_identifier = "Gain_Ripple"; //"Gain rippe" on Spara
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if ((TestName_Up.Contains("GAIN") && TestName_Up.Contains("IN-BAND") && !TestName_Up.Contains("VSWR")) ||
                     (TestName_Up.Contains("GAIN") && 
                      Is_InBand(CM_Sheet.Test_SpecID[index],CM_Sheet.Band[index], CM_Sheet.Start_Freq[index], CM_Sheet.Stop_Freq[index]) &&
                      !TestName_Up.Contains("OOB") && !TestName_Up.Contains("WIFI") && !TestName_Up.Contains("VSWR"))
                    )
            {
                IsConverted = true;
                string CA_Bands = "";
                if (CM_Sheet.CA_Band2.Count != 0)
                {
                    if (CM_Sheet.CA_Band2[index] != "") CA_Bands = "_CA" + CM_Sheet.CA_Band2[index].Trim();
                }
                if (CM_Sheet.CA_Band3.Count != 0)
                {
                    if (CM_Sheet.CA_Band3[index] != "") CA_Bands = CA_Bands + "_" + CM_Sheet.CA_Band3[index].Trim();
                }
                if (CM_Sheet.CA_Band4.Count != 0)
                {
                    if (CM_Sheet.CA_Band4[index] != "") CA_Bands = CA_Bands + "_" + CM_Sheet.CA_Band4[index].Trim();
                }

                if (CA_Bands != "")
                {
                    //Param_identifier = "RX_Gain_B" + CM_Sheet.Band[index] + CA_Bands + "_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX CA Gain" on Spara
                    Param_identifier = "RX_Gain_CA_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX CA Gain" on Spara
                }
                else
                {
                    Param_identifier = "RX_Gain_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX Inband Gain" on Spara
                }
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                RX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if (TestName_Up.Contains("GAIN") && !TestName_Up.Contains("PHASE"))
            {
                if(Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX"))
                {
                    Param_identifier = "RX_OOB_Gain"; //"OOB Gain" on Spara
                }
                else
                {
                    Param_identifier = "TX_OOB_Gain"; //"OOB Gain" on Spara
                }
                IsConverted = true;
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                Spara_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if ((TestName_Up.Contains("CURRENT") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("TX")) ||
                (TestName_Up.Contains("CURRENT") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("TRX"))
                )
            {
                IsConverted = true;
                Param_identifier = "Current_" + Find_Mode(CM_Sheet.PA_MODE[index]);
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);

                if (Test_Direction.Contains("TX"))
                {
                    try
                    {
                        if (Convert.ToSingle(CM_Sheet.Target_Pout[index]) < 10f)
                        {
                            Param_identifier = "Low" + Param_identifier;
                        }
                    }
                    catch
                    {
                        string error_convert_float = "";
                    }
                }


                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if ((TestName_Up.Contains("CURRENT") && TestName_Up.Contains("LNA")) ||
                (TestName_Up.Contains("CURRENT") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX"))
                )
            {
                IsConverted = true;
                Param_identifier = "Current_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]);
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                //Globals.DUT_CM
                RX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }


            if (TestName_Up.Contains("ACP") || TestName_Up.Contains("ACLR"))
            {
                IsConverted = true;
                Param_identifier = CM_Sheet.Signal_Standard[index].Trim().ToUpper() + "_ACP";
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("EVM"))
            {
                IsConverted = true;
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);

                if (Test_Direction == "RX")
                {
                    Param_identifier = "RX_EVM";
                }
                else
                {
                    Param_identifier = CM_Sheet.Signal_Standard[index].Trim().ToUpper() + "_EVM";
                }

                if (Test_Direction == "TX" || Test_Direction == "TRX")
                {
                    TX_defined(Param_identifier, Sheet_Name, index);
                }
                else if (Test_Direction == "RX")
                {
                    RX_defined(Param_identifier, Sheet_Name, index);
                }
                return IsConverted;
            }

            if (TestName_Up.Contains("HARMONIC") && !TestName_Up.Contains("LEAKAGE")) //except Leakage to VCC harmonic from general harmonic test. 
            {
                IsConverted = true;
                Param_identifier = Find_Harmonic_Step(TestName_Up);
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);

                if (Test_Direction == "TX" || Test_Direction == "TRX")
                {
                    TX_defined(Param_identifier, Sheet_Name, index);
                }
                else if (Test_Direction == "RX")
                {
                    RX_defined(Param_identifier, Sheet_Name, index);
                }
                return IsConverted;
            }

            if (TestName_Up.Contains("NOISE") && !TestName_Up.Contains("TRANSFER") && !Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) //except noise transfer function gain test
            {
                //string[] current_port = CM_Sheet.Output_Port[index].Split['']

                List<string> result = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);

                foreach (string port_abrnormal_check in result)
                {
                    if (port_abrnormal_check == "_Null") return false;
                }

                foreach (string Port in result)
                {
                    if (Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ANT_OUT")
                    {
                        IsConverted = true;
                        Param_identifier = "ANT_NOISE";
                        Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                        Noise_defined(Param_identifier, Sheet_Name, index);
                        return IsConverted;
                    }

                    if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "RX_OUT"))
                    {
                        IsConverted = true;
                        Param_identifier = "RXBN";
                        Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                        Noise_defined(Param_identifier, Sheet_Name, index);
                        return IsConverted;
                    }

                    if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ASM"))
                    {
                        IsConverted = true;
                        Param_identifier = "RXBN_ASM";
                        if (Port.Contains("MIMO") || Port.Contains("mimo")) Param_identifier = Param_identifier + "_MIMO";
                        if (Port.Contains("DRX") || Port.Contains("drx")) Param_identifier = Param_identifier + "_DRX";
                        if (Port.Contains("LMB") || Port.Contains("lmb")) Param_identifier = Param_identifier + "_LMB";

                        Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                        Noise_defined(Param_identifier, Sheet_Name, index);
                        return IsConverted;
                    }
                }
            }

            if ((TestName_Up.Contains("NOISE") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) ||
                (TestName_Up.Contains("NF") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX"))
                ) //noise figure (RX) and noise figure rise
            {
                if (TestName_Up.Contains("RISE"))
                {
                    IsConverted = true;
                    if (TestName_Up.Contains("RFFE")|| TestName_Up.Contains("MIPI"))
                    {
                        Param_identifier = "NFR_MIPI";
                    }
                    else
                    {
                        Param_identifier = "NFR";
                    }
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Noise_defined(Param_identifier, Sheet_Name, index);
                    return IsConverted;
                }
                else if (TestName_Up.Contains("FIGURE")|| TestName_Up.Contains("NF"))
                {
                    IsConverted = true;
                    string CA_Bands = "";
                    if (CM_Sheet.CA_Band2.Count != 0)
                    {
                        if (CM_Sheet.CA_Band2[index] != "") CA_Bands = "_CA" + CM_Sheet.CA_Band2[index].Trim();
                    }
                    if (CM_Sheet.CA_Band3.Count != 0)
                    {
                        if (CM_Sheet.CA_Band3[index] != "") CA_Bands = CA_Bands + "_" + CM_Sheet.CA_Band3[index].Trim();
                    }
                    if (CM_Sheet.CA_Band4.Count != 0)
                    {
                        if (CM_Sheet.CA_Band4[index] != "") CA_Bands = CA_Bands + "_" + CM_Sheet.CA_Band4[index].Trim();
                    }

                    if (CA_Bands != "")
                    {
                        //Param_identifier = "NF_B" + CM_Sheet.Band[index] + CA_Bands + "_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX CA Gain" on Spara
                        Param_identifier = "NF_CA_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX CA Gain" on Spara
                    }
                    else
                    {
                        Param_identifier = "NF_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]); //"RX Inband Gain" on Spara
                    }
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    RX_defined(Param_identifier, Sheet_Name, index);
                    return IsConverted;
                }
            }

            if (TestName_Up.Contains("SPECTRUM") && TestName_Up.Contains("EMISSION") && TestName_Up.Contains("MASK"))
            {
                IsConverted = true;
                Param_identifier = CM_Sheet.Signal_Standard[index].Trim().ToUpper() + "_SEM";
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if ((TestName_Up.Contains("NS_") && TestName_Up.Contains("EMISSION")) ||
                (TestName_Up.Contains("NS0") && TestName_Up.Contains("EMISSION")) ||
                (TestName_Up.Contains("NS2") && TestName_Up.Contains("EMISSION")) ||
                (TestName_Up.Contains("FCC") && TestName_Up.Contains("EMISSION"))) //older should be important : between spur > NS > supplementary emissions
            {
                IsConverted = true;
                if (TestName_Up.Contains("5")) Param_identifier = "NS_05";
                if (TestName_Up.Contains("3")) Param_identifier = "NS_03";
                if (TestName_Up.Contains("4")) Param_identifier = "NS_04";
                if (TestName_Up.Contains("21")) Param_identifier = "NS_21";
                if (!TestName_Up.Contains("21") && TestName_Up.Contains("FCC")) Param_identifier = "FCC_Emission";

                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if ((TestName_Up.Contains("SPURIOUS") && TestName_Up.Contains("EMISSION")) ||
                TestName_Up.Contains("SUPPLEMENTARY") || TestName_Up.Contains("CANADA") || TestName_Up.Contains("KOREA") || TestName_Up.Contains("VERIZON"))
            {
                List<string> result = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);

                foreach (string Port in result)
                {
                    if (Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ANT_OUT")
                    {
                        IsConverted = true;
                        Param_identifier = "SPE_" + CM_Sheet.Signal_Standard[index].Trim().ToUpper();

                        if (TestName_Up.Contains("KOREA")) Param_identifier = "KOR_" + Param_identifier;
                        if (TestName_Up.Contains("VERIZON")) Param_identifier = "VER_" + Param_identifier;

                        Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                        Noise_defined(Param_identifier, Sheet_Name, index);
                        return IsConverted;
                    }
                }

            }

            if (TestName_Up.Contains("TX") && TestName_Up.Contains("LEAKAGE")) //need to debug here : 2020.04.09
            {
                string Band_info = Convert.ToString(CM_Sheet.Band[index]);
                if (Band_info == "" && !Band_info.Contains("B"))
                {
                    string[] Band = CM_Sheet.Test_SpecID[index].Split('_');
                    Band_info = Band[0].Trim().ToUpper();
                }
                else
                {
                    Band_info = "B" + Band_info;
                }
                
                if (TestName_Up.ToUpper().Contains("HARMONIC"))
                {
                    string Frequency_Band_key = "FREQ_TX_" + Band_info;
                    string[] Frequency_range = Globals.IniFile.Frequency_table[Frequency_Band_key].Split(',');

                    string Harmonic_Step = Find_Harmonic_Step(TestName_Up);
                    int multiplier = Convert.ToInt32(Regex.Match(Harmonic_Step, @"\d+").Value);

                    double start_Freq = Convert.ToSingle(Frequency_range[0].Trim()) * multiplier;
                    double stop_Freq = Convert.ToSingle(Frequency_range[1].Trim()) * multiplier;

                    CM_Sheet.Start_Freq[index] = start_Freq.ToString();
                    CM_Sheet.Stop_Freq[index] = stop_Freq.ToString();

                    IsConverted = true;
                    Param_identifier = "TXL_" + Harmonic_Step;
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    
                    if (Test_Direction == "RX") Test_Direction = "TX"; //TXL_HAR test located in RX spec .. sometimes .. dont know why.

                    TX_defined(Param_identifier, Sheet_Name, index);
                    return IsConverted;
                }


                if(InBand_Frequency(CM_Sheet.Start_Freq[index], CM_Sheet.Stop_Freq[index], ref Band_info))
                {
                    List<string> result = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);

                    foreach (string Port in result)
                    {
                        if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "RX_OUT"))
                        {
                            IsConverted = true;
                            Param_identifier = "TXL";
                            Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                            TX_defined(Param_identifier, Sheet_Name, index);
                            return IsConverted;
                        }

                        if ((Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ASM"))
                        {
                            IsConverted = true;
                            Param_identifier = "TXL_ASM";
                            Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                            TX_defined(Param_identifier, Sheet_Name, index);
                            return IsConverted;
                        }

                        if (Globals.Spara_config_INFO.Dic_PortDefinition[Port] == "ANT_OUT")
                        {
                            IsConverted = true;
                            Param_identifier = "TXL_CPL";
                            Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                            TX_defined(Param_identifier, Sheet_Name, index);
                            return IsConverted;
                        }
                    }
                }
                else
                {
                    IsConverted = true;
                    Param_identifier = "TXL_DPX"; //not included in band range = duplex spacing TX leakage
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Noise_defined(Param_identifier, Sheet_Name, index);
                    return IsConverted;
                }
                
            }
            else if (TestName_Up.Contains("LEAKAGE")) //leakage to VCC someting
            {
                IsConverted = true;
                Param_identifier = "LEAK_RF_DC"; //not included in band range = duplex spacing TX leakage
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                DC_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

                if (TestName_Up.Contains("ISOLATION"))
            {
                if (!TestName_Up.Contains("VCC")) //except case "VCC to other" isolation case it need other test configuration.
                {
                    int index_pos = TestName_Up.IndexOf("TO");

                    string Port_input = "";
                    string Port_output = "";
                    List<string> Testcon_InputPort = new List<string>();
                    List<string> Testcon_outputPort = new List<string>();

                    if (index_pos != -1) //case there are no "TO" discription in test parameter name
                    {
                        Port_input = TestName_Up.Substring(0, index_pos).Trim();
                        Port_output = TestName_Up.Substring(index_pos + 2).Trim();

                        Testcon_InputPort = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Input_Port[index]);
                        Testcon_outputPort = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);
                    }
                    else
                    {
                        Testcon_InputPort = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Input_Port[index]);
                        Testcon_outputPort = Get_PortName(Globals.Spara_config_INFO, CM_Sheet.Output_Port[index]);

                        Port_input = Globals.Spara_config_INFO.Dic_PortDefinition[Testcon_InputPort[0]];
                        Port_output = Globals.Spara_config_INFO.Dic_PortDefinition[Testcon_outputPort[0]];

                    }

                    foreach (string port_abnormal_check in Testcon_InputPort)
                    {
                        if(port_abnormal_check == "_Null")
                        {
                            Param_identifier = "Not_Available";
                            Test_Direction = "Null";
                            return IsConverted;
                        }
                    }

                    foreach (string port_abnormal_check in Testcon_outputPort)
                    {
                        if (port_abnormal_check == "_Null")
                        {
                            Param_identifier = "Not_Available";
                            Test_Direction = "Null";
                            return IsConverted;
                        }
                    }



                    string Prefix_input = search_ISO_Port(Port_input, Testcon_InputPort);
                    string Postfix_output = search_ISO_Port(Port_output, Testcon_outputPort);
                    string Is_reverse_ISO = "";

                    if (TestName_Up.Contains("REVERSE"))
                    {
                        Is_reverse_ISO = "REV_";
                    }

                    IsConverted = true;
                    Param_identifier = Is_reverse_ISO + "ISO:" + Prefix_input + ", " + Postfix_output;
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Spara_defined(Param_identifier, Sheet_Name, index);
                    return IsConverted;
                }

            }

            if (TestName_Up.Contains("CW") && TestName_Up.Contains("P2DB") && !Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) //TX P2dB
            {
                IsConverted = true;
                Param_identifier = "CW_P2dB"; 
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if (TestName_Up.Contains("CW") && TestName_Up.Contains("P3DB") && !Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) //TX P3dB
            {
                IsConverted = true;
                Param_identifier = "CW_P3dB"; 
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }
            else if (TestName_Up.Contains("CW") && TestName_Up.Contains("P1DB") && !Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) //TX P1dB
            {
                IsConverted = true;
                Param_identifier = "CW_P1dB"; 
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                TX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("P1DB") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX")) //RX p1dB
            {
                IsConverted = true;
                Param_identifier = "RX_P1dB_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]);
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                RX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("IIP3") && Find_TestDirection(CM_Sheet.Test_SpecID[index]).Contains("RX"))
            {
                IsConverted = true;
                Param_identifier = "RX_IIP3_" + Find_Mode(CM_Sheet.LNA_Gain_Mode[index]);
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                RX_defined(Param_identifier, Sheet_Name, index);
                return IsConverted;
            }

            if (TestName_Up.Contains("MAX") && TestName_Up.Contains("OUTPUT") && TestName_Up.Contains("POWER"))
            {
                IsConverted = true;
                Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);

                if (Test_Direction == "RX")
                {
                    IsConverted = true;
                    Test_Direction = Find_TestDirection(CM_Sheet.Test_SpecID[index]);
                    Param_identifier = Test_Direction + "_Psat"; //"Insertion Loss" on Spara
                    RX_defined(Param_identifier, Sheet_Name, index);

                    return IsConverted;
                }

                if (Find_Mode(CM_Sheet.PA_MODE[index]) != "")
                {
                    Param_identifier = "MAX_Power_" + Find_Mode(CM_Sheet.PA_MODE[index]);
                    TX_defined(Param_identifier, Sheet_Name, index);
                }
                else
                {
                    Param_identifier = "MAX_Power";
                    TX_defined(Param_identifier, Sheet_Name, index);
                }
                
                return IsConverted;
            }

            Param_identifier = Param_ID_return;
            return IsConverted;
        }

        private bool Is_InBand(string spec_ID, string Band, string start_freq, string stop_freq)
        {
            bool IsInband = false;
            if (IsEmpty(start_freq) || IsEmpty(stop_freq)) return false;

            string direction = (spec_ID.ToUpper().Contains("TX") || spec_ID.ToUpper().Contains("TRX") ? "TX" : "RX");
            string[] Splited_ID = spec_ID.Split('_');
            string band_from_ID = Splited_ID[0].Trim().ToUpper();

            bool find_Range = false;

            string Frequency_Key = "FREQ_" + direction + "_" + band_from_ID.ToUpper().Trim();

            if(Globals.IniFile.Frequency_table.ContainsKey(Frequency_Key))
            {
                string[] frequencies = Globals.IniFile.Frequency_table[Frequency_Key].Split(',');
                float DEF_Start_Freq = Convert.ToSingle(frequencies[0].Trim());
                float DEF_Stop_Freq = Convert.ToSingle(frequencies[1].Trim());

                float check_start_Freq = Convert.ToSingle(start_freq);
                float check_stop_Freq = Convert.ToSingle(stop_freq);

                if (check_start_Freq >= DEF_Start_Freq && check_stop_Freq <= DEF_Stop_Freq)
                {
                    IsInband = true;
                    return IsInband;
                }
            }

            /*
            foreach (string Frequency_key in Globals.IniFile.Frequency_table.Keys)
            {
                if (Frequency_key.Contains(band_from_ID) && Frequency_key.Contains(direction) ||
                    Frequency_key.Contains(band_from_ID.ToUpper().Trim()) && Frequency_key.Contains(direction))
                {
                    
                }
            }
            */
            return IsInband;
        }
        
        public TestCon Getcondition_by_Index(string sheet_name, int key_index)
        {
            TestCon Current_row_condition = new TestCon();

            if (Globals.DUT_CM[sheet_name].Test_Name.Count != 0) Current_row_condition.Test_Name = Globals.DUT_CM[sheet_name].Test_Name[key_index];
            if (Globals.DUT_CM[sheet_name].Test_SpecID.Count != 0) Current_row_condition.Test_SpecID = Globals.DUT_CM[sheet_name].Test_SpecID[key_index];
            if (Globals.DUT_CM[sheet_name].Direction.Count != 0) Current_row_condition.Direction = Globals.DUT_CM[sheet_name].Direction[key_index];
            if (Globals.DUT_CM[sheet_name].Band.Count != 0) Current_row_condition.Band = Globals.DUT_CM[sheet_name].Band[key_index];
            if (Globals.DUT_CM[sheet_name].CA_Band2.Count != 0) Current_row_condition.CA_Band2 = Globals.DUT_CM[sheet_name].CA_Band2[key_index];
            if (Globals.DUT_CM[sheet_name].CA_Band3.Count != 0) Current_row_condition.CA_Band3 = Globals.DUT_CM[sheet_name].CA_Band3[key_index];
            if (Globals.DUT_CM[sheet_name].CA_Band4.Count != 0) Current_row_condition.CA_Band4 = Globals.DUT_CM[sheet_name].CA_Band4[key_index];
            if (Globals.DUT_CM[sheet_name].Parameter.Count != 0) Current_row_condition.Parameter = Globals.DUT_CM[sheet_name].Parameter[key_index];
            if (Globals.DUT_CM[sheet_name].Input_Port.Count != 0) Current_row_condition.Input_Port = Globals.DUT_CM[sheet_name].Input_Port[key_index];
            if (Globals.DUT_CM[sheet_name].Output_Port.Count != 0) Current_row_condition.Output_Port = Globals.DUT_CM[sheet_name].Output_Port[key_index];
            if (Globals.DUT_CM[sheet_name].LNA_Gain_Mode.Count != 0) Current_row_condition.LNA_Gain_Mode = Globals.DUT_CM[sheet_name].LNA_Gain_Mode[key_index];
            if (Globals.DUT_CM[sheet_name].Vbatt.Count != 0) Current_row_condition.Vbatt = Globals.DUT_CM[sheet_name].Vbatt[key_index];
            if (Globals.DUT_CM[sheet_name].Vdd_LNA.Count != 0) Current_row_condition.Vdd_LNA = Globals.DUT_CM[sheet_name].Vdd_LNA[key_index];

            if (Globals.DUT_CM[sheet_name].TXIn_VSWR.Count != 0) Current_row_condition.TXIn_VSWR = Globals.DUT_CM[sheet_name].TXIn_VSWR[key_index];
            if (Globals.DUT_CM[sheet_name].ANTout_VSWR.Count != 0) Current_row_condition.ANTout_VSWR = Globals.DUT_CM[sheet_name].ANTout_VSWR[key_index];
            if (Globals.DUT_CM[sheet_name].ANTIn_VSWR.Count != 0) Current_row_condition.ANTIn_VSWR = Globals.DUT_CM[sheet_name].ANTIn_VSWR[key_index];
            if (Globals.DUT_CM[sheet_name].RXOut_VSWR.Count != 0) Current_row_condition.RXOut_VSWR = Globals.DUT_CM[sheet_name].RXOut_VSWR[key_index];
            if (Globals.DUT_CM[sheet_name].Temperature.Count != 0) Current_row_condition.Temperature = Globals.DUT_CM[sheet_name].Temperature[key_index];

            if (Globals.DUT_CM[sheet_name].Start_Freq.Count != 0) Current_row_condition.Start_Freq = Globals.DUT_CM[sheet_name].Start_Freq[key_index];
            if (Globals.DUT_CM[sheet_name].Stop_Freq.Count != 0) Current_row_condition.Stop_Freq = Globals.DUT_CM[sheet_name].Stop_Freq[key_index];

            if (Globals.DUT_CM[sheet_name].IBW.Count != 0) Current_row_condition.IBW = Globals.DUT_CM[sheet_name].IBW[key_index];
            if (Globals.DUT_CM[sheet_name].PA_MODE.Count != 0) Current_row_condition.PA_MODE = Globals.DUT_CM[sheet_name].PA_MODE[key_index];
            if (Globals.DUT_CM[sheet_name].TXBand_In_RXtest.Count != 0) Current_row_condition.TXBand_In_RXtest = Globals.DUT_CM[sheet_name].TXBand_In_RXtest[key_index];
            if (Globals.DUT_CM[sheet_name].Target_Pout.Count != 0) Current_row_condition.Target_Pout = Globals.DUT_CM[sheet_name].Target_Pout[key_index];
            if (Globals.DUT_CM[sheet_name].Signal_Standard.Count != 0) Current_row_condition.Signal_Standard = Globals.DUT_CM[sheet_name].Signal_Standard[key_index];
            if (Globals.DUT_CM[sheet_name].Waveform_Category.Count != 0) Current_row_condition.Waveform_Category = Globals.DUT_CM[sheet_name].Waveform_Category[key_index];
            if (Globals.DUT_CM[sheet_name].MPR.Count != 0) Current_row_condition.MPR = Globals.DUT_CM[sheet_name].MPR[key_index];
            if (Globals.DUT_CM[sheet_name].Test_Limit_L.Count != 0) Current_row_condition.Test_Limit_L = Globals.DUT_CM[sheet_name].Test_Limit_L[key_index];
            if (Globals.DUT_CM[sheet_name].Test_Limit_Typ.Count != 0) Current_row_condition.Test_Limit_Typ = Globals.DUT_CM[sheet_name].Test_Limit_Typ[key_index];
            if (Globals.DUT_CM[sheet_name].Test_Limit_U.Count != 0) Current_row_condition.Test_Limit_U = Globals.DUT_CM[sheet_name].Test_Limit_U[key_index];
            if (Globals.DUT_CM[sheet_name].Unit.Count != 0) Current_row_condition.Unit = Globals.DUT_CM[sheet_name].Unit[key_index];
            if (Globals.DUT_CM[sheet_name].Compliance.Count != 0) Current_row_condition.Compliance = Globals.DUT_CM[sheet_name].Compliance[key_index];

            if (Globals.DUT_CM[sheet_name].Sample_1_min.Count != 0) Current_row_condition.Sample_1_min = Globals.DUT_CM[sheet_name].Sample_1_min[key_index];
            if (Globals.DUT_CM[sheet_name].Sample_1_max.Count != 0) Current_row_condition.Sample_1_max = Globals.DUT_CM[sheet_name].Sample_1_max[key_index];
            if (Globals.DUT_CM[sheet_name].Sample_2_min.Count != 0) Current_row_condition.Sample_2_min = Globals.DUT_CM[sheet_name].Sample_2_min[key_index];
            if (Globals.DUT_CM[sheet_name].Sample_2_max.Count != 0) Current_row_condition.Sample_2_max = Globals.DUT_CM[sheet_name].Sample_2_max[key_index];
            if (Globals.DUT_CM[sheet_name].Sample_3_min.Count != 0) Current_row_condition.Sample_3_min = Globals.DUT_CM[sheet_name].Sample_3_min[key_index];
            if (Globals.DUT_CM[sheet_name].Sample_3_max.Count != 0) Current_row_condition.Sample_3_max = Globals.DUT_CM[sheet_name].Sample_3_max[key_index];

            if (Globals.DUT_CM[sheet_name].Worst_Condition_text.Count != 0) Current_row_condition.Worst_Condition_text = Globals.DUT_CM[sheet_name].Worst_Condition_text[key_index];

            return Current_row_condition;
        }

        public string define_external_condition(string case_ID)
        {
            string Return_case = case_ID.Trim().ToUpper();

            if (Return_case.Contains("25") || Return_case.Contains("ROOM") || Return_case.Contains("RT"))
            {
                Return_case = "RT";
            }
            else if (Return_case.Contains("-30") || Return_case.Contains("85") || Return_case.Contains("OT"))
            {
                Return_case = "OT";
            }
            else if (Return_case.Contains("3:1") || Return_case.Contains("6:1") || Return_case.Contains("1:3") || Return_case.Contains("1:6") || Return_case.Contains("VSWR"))
            {
                Return_case = "VSWR";
            }
            else if (Return_case.Contains("50") || Return_case.Contains("1:1"))
            {
                Return_case = "50ohm";
            }
            else if (Return_case.Contains("ET") && Return_case.Contains("APT"))
            {
                Return_case = "All";
            }
            else if (Return_case.Contains("ET"))
            {
                Return_case = "ET";
            }
            else if (Return_case.Contains("APT"))
            {
                Return_case = "APT";
            }
            else if (Return_case.Contains("ALL_GAINMODE"))
            {
                Return_case = "All_GainMode";
            }
            else if (Return_case.Contains("G0")&& Return_case.Contains("G1")&& Return_case.Contains("G2") && Return_case.Contains("G3") && Return_case.Contains("G4") && Return_case.Contains("G5") && Return_case.Contains("G6"))
            {
                Return_case = "All_GainMode";
            }
            else if (Return_case.Contains("G0-G5"))
            {
                Return_case = "G0-G5";
            }
            else if (Return_case.Contains("G0") && Return_case.Contains("G1") && Return_case.Contains("G2") && Return_case.Contains("G3") && Return_case.Contains("G4") && Return_case.Contains("G5"))
            {
                Return_case = "G0-G5";
            }
            else if (Return_case.Contains("GX") || Return_case.Contains("GY"))
            {
                Return_case = "All_GainMode";
            }
            else if(Return_case.Contains("G0"))
            {
                Return_case = "G0";
            }
            else if (Return_case.Contains("G1"))
            {
                Return_case = "G1";
            }
            else if (Return_case.Contains("G2"))
            {
                Return_case = "G2";
            }
            else if (Return_case.Contains("G3"))
            {
                Return_case = "G3";
            }
            else if (Return_case.Contains("G4"))
            {
                Return_case = "G4";
            }
            else if (Return_case.Contains("G5"))
            {
                Return_case = "G5";
            }
            else if (Return_case.Contains("G6"))
            {
                Return_case = "G6";
            }

            else if (Return_case.Contains("ALL"))
            {
                Return_case = "All";
            }
            else
            {
                Return_case = "TBD";
            }

            return Return_case;
        }

        public List<string> Header_Define(string direcion)
        {
            List<string> Header_define = new List<string>();

            if(direcion.Trim().ToUpper().Contains("TX"))
            {
                Header_define.Add("Spec ID");
                Header_define.Add("Test Name");
                Header_define.Add("Band");
                Header_define.Add("Param_ID");
                Header_define.Add("Temp");
                Header_define.Add("VSWR(ANT)");
                Header_define.Add("PA MODE");
                Header_define.Add("Pout_dBm");
                Header_define.Add("SIGNAL");
                Header_define.Add("WAVEFORM");
                Header_define.Add("Start_Freq");
                Header_define.Add("Stop_Freq");
                Header_define.Add("Limit_L");
                Header_define.Add("Typical");
                Header_define.Add("Limit_U");
                Header_define.Add("Unit");
                Header_define.Add("S1_Min");
                Header_define.Add("S1_Max");
                Header_define.Add("S2_Min");
                Header_define.Add("S2_Max");
                Header_define.Add("S3_Min");
                Header_define.Add("S3_Max");
                Header_define.Add("Worst Case Description");
            }
            else if (direcion.Trim().ToUpper().Contains("TRX"))
            {
                Header_define.Add("Spec ID");
                Header_define.Add("Test Name");
                Header_define.Add("Band");
                Header_define.Add("Param_ID");
                Header_define.Add("Temp");
                Header_define.Add("VSWR(ANT)");
                Header_define.Add("PA MODE");
                Header_define.Add("Pout_dBm");
                Header_define.Add("SIGNAL");
                Header_define.Add("WAVEFORM");
                Header_define.Add("Start_Freq");
                Header_define.Add("Stop_Freq");
                Header_define.Add("Limit_L");
                Header_define.Add("Typical");
                Header_define.Add("Limit_U");
                Header_define.Add("Unit");
                Header_define.Add("S1_Min");
                Header_define.Add("S1_Max");
                Header_define.Add("S2_Min");
                Header_define.Add("S2_Max");
                Header_define.Add("S3_Min");
                Header_define.Add("S3_Max");
                Header_define.Add("Worst Case Description");
            }
            else if (direcion.Trim().ToUpper().Contains("RX"))
            {
                Header_define.Add("Spec ID");
                Header_define.Add("Test Name");
                Header_define.Add("Band");
                Header_define.Add("CA1");
                Header_define.Add("CA2");
                Header_define.Add("CA3");
                Header_define.Add("Param_ID");
                Header_define.Add("Temp");
                Header_define.Add("VSWR(ANT)");
                Header_define.Add("LNA GAIN");
                Header_define.Add("SIGNAL");
                Header_define.Add("Start_Freq");
                Header_define.Add("Stop_Freq");
                Header_define.Add("Limit_L");
                Header_define.Add("Typical");
                Header_define.Add("Limit_U");
                Header_define.Add("Unit");
                Header_define.Add("S1_Min");
                Header_define.Add("S1_Max");
                Header_define.Add("S2_Min");
                Header_define.Add("S2_Max");
                Header_define.Add("S3_Min");
                Header_define.Add("S3_Max");
                Header_define.Add("Worst Case Description");
            }

            return Header_define;
            
        }

        public string Summary_Table(Excel_File Tablefile,string sheet_name, ref int Row_index, List<string> Index_Param, string direction, string TestMode, string temp, string VSWR, List<string> exception)
        {
            CMstructure sortMemory = new CMstructure();
            List<List<string>> Sub_table = new List<List<string>>();
            List<string> BandParam_Index = Index_Param;

            Tablefile.Select_Sheet(sheet_name);

            string temp_spec = define_external_condition(temp);
            string VSWR_spec = define_external_condition(VSWR);
            string TestMode_Spec = define_external_condition(TestMode);
            Sub_table.Add(Header_Define(direction));

            string Log_text = sheet_name + ", " + BandParam_Index[0] + " : " + direction + "_" + TestMode + "_" + temp + "_" + VSWR;

            //            foreach (string exception_condition in exception)

            foreach (string band_key in BandParam_Index)
            {
                bool Is_exception = false;

                foreach (var exception_condition in exception)
                {
                    if (band_key.Contains(exception_condition))
                    {
                        Is_exception = true;
                    }
                }

                if (!Is_exception)
                {
                    if (!band_key.Trim().ToUpper().Contains(direction.Trim().ToUpper())) continue;

                    string[] SHEETnINDEX = band_key.Split(',');
                    string CMsheet_name = SHEETnINDEX[0].Trim();
                    int CMsheet_index = Convert.ToInt32(SHEETnINDEX[1].Trim());

                    TestCon TestCon = new TestCon();
                    TestCon = sortMemory.Getcondition_by_Index(CMsheet_name, CMsheet_index);

                    string temp_con = define_external_condition(TestCon.Temperature);
                    string VSWR_con = "TBD";
                    string Test_mode = "TBD";
                    List<string> Row_data = new List<string>();

                    if (band_key.ToUpper().Contains("TX") || band_key.ToUpper().Contains("TRX"))
                    {
                        VSWR_con = define_external_condition(TestCon.ANTout_VSWR);
                        Test_mode = define_external_condition(TestCon.PA_MODE);

                        if (temp_spec == temp_con && VSWR_spec == VSWR_con && TestMode_Spec == Test_mode)
                        {
                            Row_data.Add(TestCon.Test_SpecID);
                            Row_data.Add(TestCon.Test_Name.Replace('\n', ' '));
                            Row_data.Add(TestCon.Band);
                            Row_data.Add(TestCon.Parameter);
                            Row_data.Add(temp_con);
                            Row_data.Add(VSWR_con);
                            Row_data.Add(Test_mode);
                            Row_data.Add(TestCon.Target_Pout);
                            Row_data.Add(TestCon.Signal_Standard);
                            Row_data.Add(TestCon.Waveform_Category);
                            Row_data.Add(TestCon.Start_Freq);
                            Row_data.Add(TestCon.Stop_Freq);
                            Row_data.Add(TestCon.Test_Limit_L);
                            Row_data.Add(TestCon.Test_Limit_Typ);
                            Row_data.Add(TestCon.Test_Limit_U);
                            Row_data.Add(TestCon.Unit);
                            Row_data.Add(TestCon.Sample_1_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_1_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_2_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_2_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_3_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_3_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Worst_Condition_text);
                            Sub_table.Add(Row_data);
                        }
                    }
                    else if (band_key.ToUpper().Contains("RX"))
                    {
                        VSWR_con = define_external_condition(TestCon.ANTIn_VSWR);
                        Test_mode = define_external_condition(TestCon.LNA_Gain_Mode);

                        if (temp_spec == temp_con && VSWR_spec == VSWR_con && TestMode_Spec == Test_mode)
                        {
                            Row_data.Add(TestCon.Test_SpecID);
                            Row_data.Add(TestCon.Test_Name.Replace('\n',' '));
                            Row_data.Add(TestCon.Band);
                            Row_data.Add(TestCon.CA_Band2);
                            Row_data.Add(TestCon.CA_Band3);
                            Row_data.Add(TestCon.CA_Band4);
                            Row_data.Add(TestCon.Parameter);
                            Row_data.Add(temp_con);
                            Row_data.Add(VSWR_con);
                            Row_data.Add(Test_mode);
                            Row_data.Add(TestCon.Signal_Standard);
                            Row_data.Add(TestCon.Start_Freq);
                            Row_data.Add(TestCon.Stop_Freq);
                            Row_data.Add(TestCon.Test_Limit_L);
                            Row_data.Add(TestCon.Test_Limit_Typ);
                            Row_data.Add(TestCon.Test_Limit_U);
                            Row_data.Add(TestCon.Unit);
                            Row_data.Add(TestCon.Sample_1_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_1_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_2_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_2_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_3_min.Replace('\n', ' '));
                            Row_data.Add(TestCon.Sample_3_max.Replace('\n', ' '));
                            Row_data.Add(TestCon.Worst_Condition_text);
                            Sub_table.Add(Row_data);
                        }
                    }

                    
                }
            }

            //need to implement compare and Sort method for "Sub_table" here

            int Header_index = 0;
            List<string> Copy_header = new List<string>();

            foreach (List<String> Band_data in Sub_table)
            {
                if(Header_index==0)
                {
                    Tablefile.ExcelWrite_Header(sheet_name, Row_index, 1, Band_data);
                    Copy_header = Band_data;
                }
                else
                {
                    //Tablefile.ExcelWrite(sheet_name, Row_index, 1, Band_data);
                    Tablefile.ExcelWrite_Data(sheet_name, Row_index, 1, Band_data, Copy_header);
                }
                
                Row_index++;
                Header_index++;
            }

            Row_index++;

            return Log_text;
        }

        public string GetBandInfo_fromSheetName(string sheet_name)
        {
            string Band_string = "";

            string Band_name = sheet_name.Trim().ToUpper();
            string Direction = "";

            if(Band_name.Contains("TX")||Band_name.Contains("TRX"))
            {
                Direction = "TX";
            }
            else if (Band_name.Contains("RX"))
            {
                Direction = "RX";
            }
            else
            {
                Direction = "";
            }

            if(Band_name.Contains("BAND"))
            {
                Band_string = Regex.Replace(Band_name, @"\D", "");
                Band_string = (Band_name.Contains("41H") ? Band_string = "41H" : Band_string);
                Band_string = (Band_name.Contains("40A") ? Band_string = "40A" : Band_string);

                if (Direction != "") { Band_string = "B" + Band_string + "_" + Direction; }
                if (Direction == "") { Band_string = sheet_name; }
            }
            else
            {
                Band_string = sheet_name;
            }
            
            return Band_string;
        }

        public void GetInfo_FromKey(string key_text, ref int index, ref string temp, ref string param_ID, ref string input_path, ref string output_path)
        {
            string[] split_KEY_text = key_text.Split('|');
            index = Convert.ToInt16(split_KEY_text[0].Trim());
            temp = split_KEY_text[2].Trim();
            param_ID = split_KEY_text[3].Trim();
            input_path = split_KEY_text[4].Trim();
            output_path = split_KEY_text[5].Trim();
        }

        private List<string> Get_PortName(TestConfig_Spara SparaConfig, string path_strings)
        {
            List<string> Representive_Port_Name = new List<string>();

            if (path_strings.Contains("PRX_OUT1,2,3,4"))
            {
                //Port_name_exceiptions
                path_strings = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
            }

            string temp_string = path_strings.Replace('[',' ');
            temp_string = temp_string.Replace(']', ' ');
            string[] Subset_ports = temp_string.Trim().Split(',');

            for (int i = 0; i < Subset_ports.Count(); i++)
            {
                foreach (string candidate_port in SparaConfig.Dic_AvailablePort.Keys)
                {
                    if(candidate_port.Trim().ToUpper() == Subset_ports[i].Trim().ToUpper())
                    {
                        Representive_Port_Name.Add(SparaConfig.Dic_AvailablePort[candidate_port]);
                        Representive_Port_Name = Representive_Port_Name.Distinct().ToList(); //not allow duplication
                    }
                }
            }

            if (Representive_Port_Name.Count == 0) Representive_Port_Name.Add("_Null"); //To avoid null exception

            return Representive_Port_Name;
        }

        
    }

    public class TestCon
    {
        public string Test_Name;
        public string Direction;
        public string Test_SpecID;
        public string Band;
        public string CA_Band2;
        public string CA_Band3;
        public string CA_Band4;
        public string Parameter;
        //public List<string> Status_file = new List<string>();
        public string Status_file;
        public string Input_Port;
        public string Output_Port;
        public string CA_OutputPort_List;
        public string LNA_Gain_Mode;
        public string Vbatt;
        public string Vdd_LNA;

        public string TXIn_VSWR;
        public string ANTout_VSWR;
        public string ANTIn_VSWR;
        public string RXOut_VSWR;
        public string Temperature;

        public string Spara_ID;
        public string Spara_ID_DNM;
        public string Spara_Searchmethod;
        public string Spara_ConvertSign;

        public string Start_Freq;
        public string Stop_Freq;

        public string IBW;
        public string PA_MODE;
        public string TXBand_In_RXtest;
        public string Target_Pout;
        public string Signal_Standard;
        public string Waveform_Category;
        public string MPR;
        public string Test_Limit_L;
        public string Test_Limit_Typ;
        public string Test_Limit_U;
        public string Unit;
        public string Compliance;

        public string Sample_1_min;
        public string Sample_1_max;
        public string Sample_2_min;
        public string Sample_2_max;
        public string Sample_3_min;
        public string Sample_3_max;

        public string Worst_Condition_text;

        public TestCon()
        {
            Clear();
        }
        public void Clear()
        {
            this.Test_Name = "";
            this.Direction = "";
            this.Test_SpecID = "";
            this.Band = "";
            this.CA_Band2 = "";
            this.CA_Band3 = "";
            this.CA_Band4 = "";
            this.Parameter = "";
            //this.Status_file = new List<string>();
            this.Status_file = "";
            this.Input_Port = "";
            this.Output_Port = "";
            this.CA_OutputPort_List = "";
            this.LNA_Gain_Mode = "";
            this.Vbatt = "";
            this.Vdd_LNA = "";

            this.TXIn_VSWR = "";
            this.ANTout_VSWR = "";
            this.ANTIn_VSWR = "";
            this.RXOut_VSWR = "";
            this.Temperature = "";

            this.Spara_ID = "";
            this.Spara_ID_DNM = "";
            this.Spara_Searchmethod = "";
            this.Spara_ConvertSign = "";

            this.Start_Freq = "";
            this.Stop_Freq = "";

            this.IBW = "";
            this.PA_MODE = "";
            this.TXBand_In_RXtest = "";
            this.Target_Pout = "";
            this.Signal_Standard = "";
            this.Waveform_Category = "";
            this.MPR = "";
            this.Test_Limit_L = "";
            this.Test_Limit_Typ = "";
            this.Test_Limit_U = "";
            this.Unit = "";
            this.Compliance = "";
            
            this.Sample_1_min = "";
            this.Sample_1_max = "";
            this.Sample_2_min = "";
            this.Sample_2_max = "";
            this.Sample_3_min = "";
            this.Sample_3_max = "";

            this.Worst_Condition_text = "";

        }

        public TestCon Clone()
        {
            TestCon CloneTestCon = new TestCon();

            CloneTestCon.Test_Name = this.Test_Name;
            CloneTestCon.Direction = this.Direction;
            CloneTestCon.Test_SpecID = this.Test_SpecID;
            CloneTestCon.Band = this.Band;
            CloneTestCon.CA_Band2 = this.CA_Band2;
            CloneTestCon.CA_Band3 = this.CA_Band3;
            CloneTestCon.CA_Band4 = this.CA_Band4;
            CloneTestCon.CA_OutputPort_List = this.CA_OutputPort_List;
            CloneTestCon.Parameter = this.Parameter;
            CloneTestCon.Status_file = this.Status_file;
            CloneTestCon.Input_Port = this.Input_Port;
            CloneTestCon.Output_Port = this.Output_Port;
            CloneTestCon.LNA_Gain_Mode = this.LNA_Gain_Mode;
            CloneTestCon.Vbatt = this.Vbatt;
            CloneTestCon.Vdd_LNA = this.Vdd_LNA;

            CloneTestCon.TXIn_VSWR = this.TXIn_VSWR;
            CloneTestCon.ANTout_VSWR = this.ANTout_VSWR;
            CloneTestCon.ANTIn_VSWR = this.ANTIn_VSWR;
            CloneTestCon.RXOut_VSWR = this.RXOut_VSWR;
            CloneTestCon.Temperature = this.Temperature;

            CloneTestCon.Spara_ID = this.Spara_ID;
            CloneTestCon.Spara_ID_DNM = this.Spara_ID_DNM;
            CloneTestCon.Spara_Searchmethod = this.Spara_Searchmethod;
            CloneTestCon.Spara_ConvertSign = this.Spara_ConvertSign;

            CloneTestCon.Start_Freq = this.Start_Freq;
            CloneTestCon.Stop_Freq = this.Stop_Freq;

            CloneTestCon.IBW = this.IBW;
            CloneTestCon.PA_MODE = this.PA_MODE;
            CloneTestCon.TXBand_In_RXtest = this.TXBand_In_RXtest;
            CloneTestCon.Target_Pout = this.Target_Pout;
            CloneTestCon.Signal_Standard = this.Signal_Standard;
            CloneTestCon.Waveform_Category = this.Waveform_Category;
            CloneTestCon.MPR = this.MPR;
            CloneTestCon.Test_Limit_L = this.Test_Limit_L;
            CloneTestCon.Test_Limit_Typ = this.Test_Limit_Typ;
            CloneTestCon.Test_Limit_U = this.Test_Limit_U;
            CloneTestCon.Unit = this.Unit;
            CloneTestCon.Compliance = this.Compliance;

            CloneTestCon.Sample_1_min = this.Sample_1_min;
            CloneTestCon.Sample_1_max = this.Sample_1_max;
            CloneTestCon.Sample_2_min = this.Sample_2_min;
            CloneTestCon.Sample_2_max = this.Sample_2_max;
            CloneTestCon.Sample_3_min = this.Sample_3_min;
            CloneTestCon.Sample_3_max = this.Sample_3_max;

            CloneTestCon.Worst_Condition_text = this.Worst_Condition_text;

            return CloneTestCon;
        }
    }


    public class Table_data
    {
        public string Test_Name;
        public string Test_SpecID;
        public string Band;
        public string CA_Band2;
        public string CA_Band3;
        public string CA_Band4;
        public string Parameter;
        public string Status_file;
        public string Input_Port;
        public string Output_Port;
        public string LNA_Gain_Mode;
        public string Vbatt;
        public string Vdd_LNA;

        public string TXIn_VSWR;
        public string ANTout_VSWR;
        public string ANTIn_VSWR;
        public string RXOut_VSWR;
        public string Temperature;

        public string Start_Freq;
        public string Stop_Freq;

        public string IBW;
        public string PA_MODE;
        public string TXBand_In_RXtest;
        public string Target_Pout;
        public string Signal_Standard;
        public string Waveform_Category;
        public string MPR;
        public string Test_Limit_L;
        public string Test_Limit_Typ;
        public string Test_Limit_U;
        public string Unit;
        public string Compliance;

        public string Sample1_min;
        public string Sample1_max;
        public string Sample2_min;
        public string Sample2_max;
        public string Sample3_min;
        public string Sample3_max;

        public string Worst_condition;

        public Table_data()
        {
            Clear();
        }
        public void Clear()
        {
            this.Test_Name = "";
            this.Test_SpecID = "";
            this.Band = "";
            this.CA_Band2 = "";
            this.CA_Band3 = "";
            this.CA_Band4 = "";
            this.Parameter = "";
            this.Status_file = "";
            this.Input_Port = "";
            this.Output_Port = "";
            this.LNA_Gain_Mode = "";
            this.Vbatt = "";
            this.Vdd_LNA = "";

            this.TXIn_VSWR = "";
            this.ANTout_VSWR = "";
            this.ANTIn_VSWR = "";
            this.RXOut_VSWR = "";
            this.Temperature = "";

            this.Start_Freq = "";
            this.Stop_Freq = "";

            this.IBW = "";
            this.PA_MODE = "";
            this.TXBand_In_RXtest = "";
            this.Target_Pout = "";
            this.Signal_Standard = "";
            this.Waveform_Category = "";
            this.MPR = "";
            this.Test_Limit_L = "";
            this.Test_Limit_Typ = "";
            this.Test_Limit_U = "";
            this.Unit = "";
            this.Compliance = "";

            this.Sample1_min = "";
            this.Sample1_max = "";
            this.Sample2_min = "";
            this.Sample2_max = "";
            this.Sample3_min = "";
            this.Sample3_max = "";

            this.Worst_condition = "";
        }
    }

    public class Spara_Trigger_Group
    {
        public string Group_TYP;
        public string Status_File;
        public string Tempearature;

        public string Ports_Sequence;
        public SortedList<int, int> Ports_Assigned = new SortedList<int, int>(); //Portnum = key, order = value

        public string Direction;
        public string Tech;
        public string Test_Input;
        public string Test_Output;

        public string TX_Input;
        public string ANT_Output;
        public string PowerMode;

        public string CA_Case;
        public string RX_Output;
        public string RX_Mode;

        public string TDD_Priority;

        public string ASM1;
        public string ASM2;
        public string ASM3;

        SortedList<int, int> Port_num = new SortedList<int, int>();
        public List<TestCon> TestCon_List = new List<TestCon>();

        public Spara_Trigger_Group()
        {
            clear();
        }

        public void clear()
        {
            this.Group_TYP = "";
            this.Status_File = "";
            this.Tempearature = "";

            this.Ports_Sequence = "";
            this.Ports_Assigned = new SortedList<int, int>();

            this.Direction = "";
            this.Tech = "";
            this.Test_Input = "";
            this.Test_Output = "";

            this.TX_Input = "";
            this.ANT_Output = "";
            this.PowerMode = "";

            this.CA_Case = "";
            this.RX_Output = "";
            this.RX_Mode = "";

            this.TDD_Priority = "";

            this.ASM1 = "TERM";
            this.ASM2 = "TERM";
            this.ASM3 = "TERM";
            this.TestCon_List = new List<TestCon>();
            this.Port_num = new SortedList<int, int>();
        }

        public bool Port_matching(string target_input, string target_output, string Input_match_type, string Output_match_type)
        {
            bool Input_matched = false;
            bool Output_matched = false;
            bool perfect_match = false;

            if (Input_match_type.Contains("TRUE"))
            {
                if (this.Test_Input == target_input) Input_matched = true;
            }
            else if(Input_match_type.Contains("FALSE"))
            {
                if (this.Test_Input != target_input) Input_matched = true;
            }
            else if (Input_match_type.Contains("REV"))
            {
                if (this.Test_Output == target_input) Input_matched = true;
            }
            else if (Input_match_type.Contains("X"))
            {
                Input_matched = true;
            }

            if (Output_match_type.Contains("TRUE"))
            {
                if (this.Test_Output == target_output) Output_matched = true;
            }
            else if (Output_match_type.Contains("FALSE"))
            {
                if (this.Test_Output != target_output) Output_matched = true;
            }
            else if (Input_match_type.Contains("REV"))
            {
                if (this.Test_Input == target_output) Output_matched = true;
            }
            else if (Output_match_type.Contains("X"))
            {
                Output_matched = true;
            }

            if (Input_matched && Output_matched)
            {
                perfect_match = true;
            }

            return perfect_match;
        }

        public void Port_num_sorting()
        {
            int Dynamic_Port_num = 1;

            foreach (int Port_Physical_num in this.Port_num.Keys)
            {
                this.Port_num[Port_Physical_num] = Dynamic_Port_num;
                Dynamic_Port_num++;
            }
        }
    }

    public class TestConfig_Spara
    {
        public float NA_FREQ_start;
        public float NA_FREQ_stop;
        
        public Dictionary<string, int> Dic_PortNum; //return "port num" of defined port name
        public Dictionary<string, string> Dic_AvailablePort; //porting CM in/out port to "defined port name"
        public Dictionary<string, string> Dic_PortDefinition; //check "type" of defined port name

        public Dictionary<string, List<string>> Status_config; //string "port-relation", "status file name"
        public Dictionary<string, string> StatusFreqrange; //string "staus file name", "frequency range"

        public Dictionary<string, string> Group_Bands;
        public Dictionary<String, SortedList<string, SortedList<string, string>>> Group_def_dic;
        
        public TestConfig_Spara()
        {
            clear();
            Load_SparaConfig();
        }
        public void clear()
        {
            this.NA_FREQ_start = 0f;
            this.NA_FREQ_stop = 0f;
            this.Dic_PortNum = new Dictionary<string, int>();
            this.Dic_PortDefinition = new Dictionary<string, string>();
            this.Dic_AvailablePort = new Dictionary<string, string>();
            
            this.Status_config = new Dictionary<string, List<string>>();
            this.StatusFreqrange = new Dictionary<string, string>();

            this.Group_Bands = new Dictionary<string, string>();
            this.Group_def_dic = new Dictionary<string, SortedList<string, SortedList<string, string>>>(); //Group band / Sub_group <test ID, in-out definition>
        }

        public string Get_Group(string band_name)
        {
            string BAND_key = band_name.Trim().ToUpper();
            string find_group = this.Group_Bands[BAND_key];
            return find_group;
        }

        public string Get_SubGroup(string band_name, string test_id)
        {
            string BAND_key = band_name.Trim().ToUpper();
            string TestID_key = test_id.Trim().ToUpper();
            string sub_Group = "";

            foreach (var item in this.Group_def_dic[this.Group_Bands[BAND_key]])
            {
                if (item.Value.ContainsKey(TestID_key))
                {
                    sub_Group = item.Key;
                    break;
                }
            }

            if(sub_Group == "")
            {
                sub_Group = "Not_Found";
            }

            return sub_Group;
        }

        public string Get_SubGroup(string band_name, string test_id, string gain_Mode, bool Is_CA)
        {
            string BAND_key = band_name.Trim().ToUpper();
            string TestID_key = test_id.Trim().ToUpper();
            List<string> sub_Group = new List<string>();
            string return_group = "";

            foreach (var item in this.Group_def_dic[this.Group_Bands[BAND_key]])
            {
                if (item.Value.ContainsKey(TestID_key))
                {
                    sub_Group.Add(item.Key);
                }
            }

            if (sub_Group.Count == 0)
            {
                return_group = "Not_Found";
            }
            else
            {
                foreach (var item in sub_Group)
                {
                    if(!Is_CA)
                    {
                        if (item.Contains(gain_Mode))
                        {
                            return_group = item;
                            return return_group;
                        }
                        else
                        {
                            return_group = item;
                        }
                    }
                    else
                    {
                        if (item.Contains(gain_Mode) && item.Contains("_CA_"))
                        {
                            return_group = item;
                            return return_group;
                        }
                        else
                        {
                            return_group = item;
                        }
                    }
                    
                }
            }

            return return_group;
        }

        public string Get_SubGroup(string band_name, string test_id, bool Is_CA)
        {
            string BAND_key = band_name.Trim().ToUpper();
            string TestID_key = test_id.Trim().ToUpper();
            List<string> sub_Group = new List<string>();
            string return_group = "";

            foreach (var item in this.Group_def_dic[this.Group_Bands[BAND_key]])
            {
                if (item.Value.ContainsKey(TestID_key))
                {
                    sub_Group.Add(item.Key);
                }
            }

            if (sub_Group.Count == 0)
            {
                return_group = "Not_Found";
            }
            else
            {
                foreach (var item in sub_Group)
                {
                    if(Is_CA && item.Contains("CA"))
                    {
                        return_group = item;
                    }
                    else if(!Is_CA && !item.Contains("CA"))
                    {
                        return_group = item;
                    }
                }
            }

            return return_group;
        }

        public SortedList<string, string> Get_port_matching(string band_name, string test_id)
        {
            string BAND_key = band_name.Trim().ToUpper();
            string TestID_key = test_id.Trim().ToUpper();
            SortedList<string, string> port_matching = new SortedList<string, string>();

            foreach (var item in this.Group_def_dic[this.Group_Bands[BAND_key]])
            {
                if (item.Value.ContainsKey(TestID_key))
                {
                    string matching_definition = item.Value[TestID_key];
                    string[] matching_inout = matching_definition.Split(',');
                    int i = 0;
                    string port_ch = "";

                    foreach (string Port_matching in matching_inout)
                    {
                        if (i == 0) port_ch = "INPUT";
                        if (i == 1) port_ch = "OUTPUT";

                        if (Port_matching.Trim().ToUpper().Contains("TRUE"))
                        {
                            port_matching.Add(port_ch, Port_matching.Trim().ToUpper());
                        }
                        else
                        {
                            port_matching.Add(port_ch, Port_matching.Trim().ToUpper());
                        }
                        i++;
                    }
                    break;
                }
            }

            if(port_matching.Count<2)
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error: can't find matching port at Band \"{0}\"\n Test param = {1}", band_name, test_id);
                ErrMsg.AppendFormat("\nNeed to debugging \"TEST_GROUP_MAPPING \" at Port_config.ini");
                ClsMsgBox.Show("not found port during test plan generation", ErrMsg.ToString());
            }

            return port_matching;
        }

        public string Input_Match(string band_name, string test_id)
        {
            SortedList<string, string> port_matching = new SortedList<string, string>();
            port_matching = Get_port_matching(band_name, test_id);

            string matching = port_matching["INPUT"];
            return matching;
        }

        public string Output_Match(string band_name, string test_id)
        {
            SortedList<string, string> port_matching = new SortedList<string, string>();
            port_matching = Get_port_matching(band_name, test_id);

            string matching = port_matching["OUTPUT"];
            return matching;
        }


        public void Load_SparaConfig()
        {
            try
            {
                string RootDir = Globals.Spara_Info.Trim();
                bool Load_Portinfo = false;
                bool Load_Statusinfo = false;
                bool Load_TestGroupinfo = false;

                List<string> status_relation_list = new List<string>();
                List<string> status_name_list = new List<string>();
                List<string> status_freq_list = new List<string>();

                string current_Group = "";
                string current_SubGroup = "";
                string last_Group = "";
                string last_SubGroup = "";

                using (StreamReader sr = new StreamReader(RootDir))
                {
                    string Line;
 
                    while ((Line = sr.ReadLine()) != null)
                    {
                        int Port_Num;
                        string Port_Name;
                        string Port_define;
                        List<string> Available_Port = new List<string>();


                        bool ValidLine = ((Line.Contains('=')||Line.Contains('[') || Line.Contains(']') || Line.Contains(':')) && !IsComment(Line));
                        if (!ValidLine) continue;

                        string current_line = Line.Trim().ToUpper();

                        if(current_line.Contains("VNA_FREQ_RANGE"))
                        {
                            try
                            {
                                string[] Substr = Line.Split('=');
                                this.RemoveComment(ref Substr[1]);
                                string[] freq_range = Substr[1].Trim().Split('-');

                                this.NA_FREQ_start = Convert.ToSingle(freq_range[0].Trim());
                                this.NA_FREQ_stop = Convert.ToSingle(freq_range[1].Trim());
                                continue;
                            }
                            catch
                            {
                                this.NA_FREQ_start = 10f;
                                this.NA_FREQ_stop = 8500f;
                                ClsMsgBox.Show("There are no frequency range info in Spara_config File:\n");
                            }
                        }

                        if (current_line.Contains("DEFAULT_ANT"))
                        {
                            string[] Get_definition = Line.Split('=');
                            Globals.Default_ANT = Get_definition[1].Trim();
                        }

                        if (current_line.Contains("PORT_MAPPING") && current_line.Contains("[START]")) Load_Portinfo = true; 
                        if (current_line.Contains("PORT_MAPPING") && current_line.Contains("[END]")) Load_Portinfo = false;
                        if (Load_Portinfo && current_line.Contains('='))
                        {
                            try
                            {
                                string[] Get_definition = Line.Split(':');
                                Port_define = Get_definition[0].Trim().ToUpper();

                                string[] Substr = Get_definition[1].Trim().Split('=');
                                Port_Num = Convert.ToInt16(Substr[0].Trim()); //Port number index

                                string Port_value = Substr[1].Trim(); //representive Port name + available port name
                                this.RemoveComment(ref Port_value);

                                string[] available_port_name = Port_value.Split('<');
                                Port_Name = available_port_name[0].Trim().ToUpper(); //get representive port name
                                if (Port_Name.Contains("NOT USED")) continue;

                                string[] Available_Ports = available_port_name[1].Trim().Split(','); //get available port list
                                for (int i = 0; i < Available_Ports.Length; i++)
                                {
                                    Available_Port.Add(Available_Ports[i].Trim().ToUpper());
                                }
                                Set_PortInfo(Port_Name, Port_Num, Available_Port, Port_define); 
                                continue;
                            }
                            catch
                            {
                                ClsMsgBox.Show("Error during Loading Port information from Spara_config File:\n");
                            }
                        }
                        if (current_line.Contains("STATUS_FILE_MAPPING") && current_line.Contains("[START]")) Load_Statusinfo = true;
                        if (current_line.Contains("STATUS_FILE_MAPPING") && current_line.Contains("[END]")) Load_Statusinfo = false;
                        if (Load_Statusinfo && current_line.Contains('='))
                        {
                            try
                            {
                                string line_status = current_line;
                                this.RemoveComment(ref line_status);

                                string[] Get_definition = line_status.Split('=');
                                string Port_relation = Get_definition[0].Trim().ToUpper();
                                string[] assemble_string = Port_relation.Split('-');
                                Port_relation = assemble_string[0].Trim() + "-" + assemble_string[1].Trim(); //remove space between port relation string

                                string status_file_name = Get_definition[1].Trim().ToUpper();
                                string[] assemble_string_file = status_file_name.Split('.');
                                status_file_name = assemble_string_file[0].Trim() + ".znx"; //revise extension string to lower 

                                string frequency_range = Get_definition[2].Trim().ToUpper();
                                string[] assemble_string_freq = frequency_range.Split('-');
                                frequency_range = assemble_string_freq[0].Trim() + "-" + assemble_string_freq[1].Trim(); //remove space between port relation string

                                status_relation_list.Add(Port_relation);
                                status_name_list.Add(status_file_name);
                                status_freq_list.Add(frequency_range);
                                continue;
                            }
                            catch
                            {
                                ClsMsgBox.Show("Error during Loading Status mapping setting from Spara_config File:\n");
                            }
                        }
                        if (current_line.Contains("TEST_GROUP_MAP") && current_line.Contains("[START]")) Load_TestGroupinfo = true;
                        if (current_line.Contains("TEST_GROUP_MAP") && current_line.Contains("[END]")) Load_TestGroupinfo = false;
                        if (Load_TestGroupinfo && current_line.Contains('='))
                        {
                            string line_group = current_line;
                            List<string> Group_lines = new List<string>();
                            List<string> Group_list = new List<string>();

                            SortedList<string, string> temp_Subgroup = new SortedList<string, string>();
                            SortedList<string, SortedList<string, string>> Dic_Subgroup = new SortedList<string, SortedList<string, string>>();

                            bool flag_subgroup_end = false;
                            Group_lines.Add(line_group);

                            do
                            {
                                string temp_line = sr.ReadLine();
                                if (temp_line!=null)
                                {
                                    this.RemoveComment(ref temp_line);
                                    if (temp_line!="")
                                    {
                                        line_group = temp_line.ToUpper().Trim();
                                        Group_lines.Add(line_group); 
                                    }
                                }
                            } while (!line_group.Contains("TEST_GROUP_MAPPING") && !line_group.Contains("END"));

                            Load_TestGroupinfo = false;
                            bool open_subGroup = false;
                            bool close_subGroup = false;
                            bool lastline = false;

                            foreach (string each_line in Group_lines)
                            {
                                if (each_line.Contains("TEST_GROUP_MAPPING") && each_line.Contains("END")) lastline = true;

                                if (each_line.Contains("DEFINE_BAND"))
                                {
                                    string[] temp_line = each_line.Split(':');
                                    string[] Group_desc = temp_line[1].Trim().Split('=');
                                    string[] Band_desc = Group_desc[1].Split(',');

                                    string Group_key = Group_desc[0].Trim();
                                    Group_list.Add(Group_key);

                                    for (int i = 0; i < Band_desc.Length; i++)
                                    {
                                        this.Group_Bands.Add(Band_desc[i].Trim(), Group_key);
                                    }
                                    continue;
                                }

                                if (each_line.Contains("=") && !each_line.Trim().StartsWith("="))
                                {
                                    string[] SubGroup_desc = each_line.Split('=');
                                    current_Group = SubGroup_desc[0].Trim();
                                    current_SubGroup = SubGroup_desc[1].Trim();                                   
                                    if (open_subGroup) close_subGroup = true;
                                }

                                if(each_line.Trim().StartsWith("=")&& each_line.Trim().Contains(">"))
                                {
                                    string[] TestID_desc = each_line.Replace("=", " ").Trim().Split('>');
                                    temp_Subgroup.Add(TestID_desc[0].Trim(), TestID_desc[1].Trim());

                                    last_Group = current_Group;
                                    last_SubGroup = current_SubGroup;
                                    open_subGroup = true;
                                }

                                if (close_subGroup || lastline)
                                {
                                    SortedList<string, string> TestID_list = new SortedList<string, string>();
                                    foreach (var item in temp_Subgroup)
                                    {
                                        TestID_list.Add(item.Key, item.Value);
                                    }
                                    Dic_Subgroup.Add(last_SubGroup, TestID_list);
                                    temp_Subgroup.Clear();

                                    open_subGroup = false;
                                    close_subGroup = false;
                                }

                                if ((current_Group != last_Group && current_Group != "" && last_Group != "")|| lastline)
                                {
                                    SortedList<string, SortedList<string, string>> Dic_Sub = new SortedList<string, SortedList<string, string>>();
                                    foreach (var item_key in Dic_Subgroup.Keys)
                                    {
                                        Dic_Sub.Add(item_key, Dic_Subgroup[item_key]);
                                    }
                                    
                                    this.Group_def_dic.Add(last_Group, Dic_Sub);
                                    Dic_Subgroup.Clear();
                                }
                            }
                        }
                    }
                    //Post Process after closing Spara config file
                    Set_StatusInfo(status_relation_list, status_name_list, status_freq_list);

                }
            }
            catch
            {
                this.Dic_PortNum.Clear();
                ClsMsgBox.Show("Error during Loading Spara Config Ini File:\n" + Globals.Spara_Info);
                Environment.Exit(0);
            }
        }

        public string Get_GroupType(string param, string input_port, string Output_port )
        {
            string Type_group = "";

            foreach (var item in this.Group_Bands.Keys)
            {
                
            } 
            
            return Type_group;
        }


        private void Set_StatusInfo(List<string> port_relation, List<string> status_name, List<string> freq_range)
        {
            for (int i = 0; i < status_name.Count; i++)
            {
                string Port_relation_Key = "";
                List<string> status_file_list = new List<string>();

                for (int j = 0; j < port_relation.Count; j++)
                {
                    if (port_relation[i] == port_relation[j])
                    {
                        status_file_list.Add(status_name[j]);
                        Port_relation_Key = port_relation[i];
                    }
                }

                if(Port_relation_Key!="" && !this.Status_config.ContainsKey(Port_relation_Key))
                {
                    this.Status_config.Add(Port_relation_Key, status_file_list);
                }

                this.StatusFreqrange.Add(status_name[i], freq_range[i]); //this line should be placed here for exclusive matching. 
            }
        }

        public string Get_StatusInfo(string port1, string port2, string start_F, string stop_F, TestConfig_Spara SparaConfig)
        {
            List<string> candidated_status = new List<string>();

            string Revised_port1 = SparaConfig.Dic_PortDefinition[port1.Trim()];
            string Revised_port2 = SparaConfig.Dic_PortDefinition[port2.Trim()];

            string relation_key = Revised_port1 + "-" + Revised_port2;
            string relation_reverse_key = Revised_port2 + "-" + Revised_port1;


            foreach (string port_relation in this.Status_config.Keys)
            {
                if (port_relation.Contains(relation_key))
                {
                    candidated_status = this.Status_config[relation_key];
                    break;
                }
                else if (port_relation.Contains(relation_reverse_key))
                {
                    candidated_status = this.Status_config[relation_reverse_key];
                    break;
                }
            }

            string result = "";
            float Last_status_start = 0f;
            float Last_status_stop = 50000f;

            foreach (string keys in candidated_status)
            {

                string status_freq = this.StatusFreqrange[keys.Trim()];
                string[] frequencys = status_freq.Split('-');

                float status_start = Convert.ToSingle(frequencys[0].Trim());
                float status_stop = Convert.ToSingle(frequencys[1].Trim());
                float test_start = Convert.ToSingle(start_F.Trim());
                float test_stop = Convert.ToSingle(stop_F.Trim());

                if (status_start <= test_start && status_stop >= test_stop) 
                {
                    if (status_start >= Last_status_start && status_stop <= Last_status_stop)
                    {
                        result = keys;
                        Last_status_start = status_start;
                        Last_status_stop = status_stop;
                    }
                }

                if (result == "")
                {
                    ClsMsgBox.Show("Get_StatusInfo error in Sparacon sorting: No available status file\n");
                }
            }

            return result;
        }

        public string Get_PropStatusInfo(string port1, string port2, string portnum1, string portnum2, TestCon condition)
        {
            List<string> result = new List<string>();
            string final_Status = "";

            bool IsStatusFile_Exist = false;
            bool IsMatched_Frequency = false;
            bool IsContain_PortNum = false;

            foreach (string port_relation in this.Status_config.Keys)
            {
                string relation_key = port1.Trim() + "-" + port2.Trim();
                string relation_reverse_key = port2.Trim() + "-" + port1.Trim();

                if (port_relation.Contains(relation_key))
                {
                    result = this.Status_config[relation_key];
                    IsStatusFile_Exist = true;
                    break;
                }
                else if (port_relation.Contains(relation_reverse_key))
                {
                    result = this.Status_config[relation_reverse_key];
                    IsStatusFile_Exist = true;
                    break;
                }
            }

            foreach (string status_name in result)
            {
                string[] frequency = this.StatusFreqrange[status_name].Split('-');
                
                if(Convert.ToSingle(condition.Start_Freq) >= Convert.ToSingle(frequency[0].Trim())&&
                    Convert.ToSingle(condition.Stop_Freq) <= Convert.ToSingle(frequency[1].Trim()))
                {
                    IsMatched_Frequency = true;
                    final_Status = status_name;
                    break;
                }
            }

            return final_Status;
        }

        private void Set_PortInfo(string key, int port_index, List<string> availablePorts, string Port_define)
        {
            string keyUP = key.ToUpper();
            bool IsNewKey = (!this.Dic_PortNum.ContainsKey(keyUP) && !this.Dic_AvailablePort.ContainsKey(keyUP));
            
            if (IsNewKey) //Build All INI setting param list include value
            {
                this.Dic_PortNum.Add(keyUP, port_index);
                this.Dic_PortDefinition.Add(keyUP, Port_define);
            }
            else
            {
                this.Dic_PortNum.Remove(keyUP);
                this.Dic_PortDefinition.Remove(keyUP);

                this.Dic_PortNum.Add(keyUP, port_index);
                this.Dic_PortDefinition.Add(keyUP, Port_define);
            }

            bool IsNewKey_for_available_port = false;

            foreach (string each_port in availablePorts)
            {
                string Key_candidated_port = each_port.Trim().ToUpper();
                if(!this.Dic_AvailablePort.ContainsKey(Key_candidated_port))
                {
                    IsNewKey_for_available_port = true;
                }

                if(IsNewKey_for_available_port)
                {
                    this.Dic_AvailablePort.Add(Key_candidated_port, keyUP);
                    IsNewKey_for_available_port = false;
                }
                else
                {
                    this.Dic_AvailablePort.Remove(Key_candidated_port);
                    this.Dic_AvailablePort.Add(Key_candidated_port, keyUP);
                }

                
            }
        }

        private bool IsComment(string strings)
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

        private void RemoveComment(ref string Val)
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
