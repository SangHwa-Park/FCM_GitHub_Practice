using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Test_Planner
{
    public partial class Testsetting : Form
    {

        public List<string> Band_key = new List<string>();
        public List<string> Param_key = new List<string>();
        public TreeNode Root_Treeview = new TreeNode();

        public Testsetting()
        {
            InitializeComponent();
        }

        public void ShowForm()
        {
            this.Interface_Setting();
            this.Listview_Header_initialize();
            this.Set_treeview_with_Expanded_Seq(Globals.Expaned_Spara_Seq);
          
            this.ShowDialog();
        }

        public void Interface_Setting()
        {
            CMstructure sortMemory = new CMstructure();
            //TestConfig_Spara Spara_config = Globals.Spara_config_INFO;
            List<string> Defined_Bands = Globals.IniFile.Band_Sheet_Name;
            Dictionary<string, List<string>> Spara_List = Globals.Spara_TestDic;

            List<string> menu_band = new List<string>();
            List<string> menu_temp = new List<string>();
            List<string> menu_input_port = new List<string>();
            List<string> menu_output_port = new List<string>();
            List<string> menu_not_included = new List<string>();


            Dictionary<string, Dictionary<string, List<TestCon>>> FullSEQ_Bands = new Dictionary<string, Dictionary<string, List<TestCon>>>();

            foreach (string TXRX_Bands in Defined_Bands)
            {
                Dictionary<string, List<TestCon>> FullSEQ_TestID = new Dictionary<string, List<TestCon>>();

                foreach (string Test_ID in Spara_List.Keys)
                {  
                    List<TestCon> Current_Seq = new List<TestCon>();

                    foreach (string band_key in Spara_List[Test_ID])
                    {
                        string[] SHEETnINDEX = band_key.Split(',');
                        string CMsheet_name = SHEETnINDEX[0].Trim();
                        int CMsheet_index = Convert.ToInt32(SHEETnINDEX[1].Trim());

                        if (CMsheet_name.Trim().ToUpper() == TXRX_Bands.Trim().ToUpper())
                        {
                            menu_band = menu_Add_WO_Duplication(menu_band, band_key);

                            TestCon TestCon = new TestCon();
                            TestCon = sortMemory.Getcondition_by_Index(CMsheet_name, CMsheet_index);
                            TestCon.CA_OutputPort_List = TestCon.Output_Port;

                            if (Test_ID == "ISO:TX, RX") 
                            {
                                Get_BandCA(ref TestCon);
                            }

                            if (IsEmpty(TestCon.Input_Port)) TestCon.Input_Port = "";
                            if (IsEmpty(TestCon.Output_Port)) TestCon.Output_Port = "";

                            //Set status file based on target frequency and port, direction
                            TestCon.Status_file = setup_statusfile(CMsheet_name, TestCon);


                            TestCon CloneTestCon = TestCon.Clone();
                            if (CloneTestCon.Temperature.Contains("25"))
                            {
                                CloneTestCon.Temperature = "25";
                                menu_temp = menu_Add_WO_Duplication(menu_temp, CloneTestCon.Temperature);
                                Current_Seq = Testcon_expansion_Temp_Port(CloneTestCon, Current_Seq, ref menu_input_port, ref menu_output_port);

                            }
                            else if (CloneTestCon.Temperature.Contains("-30") || CloneTestCon.Temperature.Contains("85"))
                            {
                                CloneTestCon.Temperature = "25";
                                menu_temp = menu_Add_WO_Duplication(menu_temp, CloneTestCon.Temperature);
                                Current_Seq = Testcon_expansion_Temp_Port(CloneTestCon, Current_Seq, ref menu_input_port, ref menu_output_port);

                                TestCon CloneTestCon_HT = TestCon.Clone();
                                CloneTestCon_HT.Temperature = "85";
                                menu_temp = menu_Add_WO_Duplication(menu_temp, CloneTestCon_HT.Temperature);
                                Current_Seq = Testcon_expansion_Temp_Port(CloneTestCon_HT, Current_Seq, ref menu_input_port, ref menu_output_port);

                                TestCon CloneTestCon_LT = TestCon.Clone();
                                CloneTestCon_LT.Temperature = "-30";
                                menu_temp = menu_Add_WO_Duplication(menu_temp, CloneTestCon_LT.Temperature);
                                Current_Seq = Testcon_expansion_Temp_Port(CloneTestCon_LT, Current_Seq, ref menu_input_port, ref menu_output_port);
                            }
                            else
                            {
                                menu_temp = menu_Add_WO_Duplication(menu_temp, CloneTestCon.Temperature);
                                Current_Seq = Testcon_expansion_Temp_Port(CloneTestCon, Current_Seq, ref menu_input_port, ref menu_output_port);
                            }
                        }
                    }

                    if(Current_Seq.Count!=0)
                    {
                        FullSEQ_TestID.Add(Test_ID, Current_Seq);
                    }                   
                }
                FullSEQ_Bands.Add(TXRX_Bands, FullSEQ_TestID);
            }
            
            Globals.Expaned_Spara_Seq = FullSEQ_Bands;
        }

        public void Get_BandCA(ref TestCon TestCon_CA)
        {
            string testnameU = TestCon_CA.Test_Name.ToUpper().Trim();
            string CA_band = "";

            int index_pos = testnameU.IndexOf("TO");
            if (index_pos != -1)
            {
                string Pre_string = testnameU.Substring(0, index_pos).Trim();
                string Post_string = testnameU.Substring(index_pos + 2).Trim();

                if(Post_string.Contains("BAND") && !Post_string.Contains("/"))
                {
                    CA_band = System.Text.RegularExpressions.Regex.Match(Post_string, @"\d+").Value;

                    if (TestCon_CA.CA_Band2 == "") TestCon_CA.CA_Band2 = CA_band;
                    else if (TestCon_CA.CA_Band3 == "") TestCon_CA.CA_Band3 = CA_band;
                    else if (TestCon_CA.CA_Band4 == "") TestCon_CA.CA_Band4 = CA_band;

                }
                else if(Post_string.Contains("1/66"))
                {
                    CA_band = "1";
                    if (TestCon_CA.CA_Band2 == "") TestCon_CA.CA_Band2 = CA_band;
                    else if (TestCon_CA.CA_Band3 == "") TestCon_CA.CA_Band3 = CA_band;
                    else if (TestCon_CA.CA_Band4 == "") TestCon_CA.CA_Band4 = CA_band;
                }
                else if (Post_string.Contains("66"))
                {
                    CA_band = "66";
                    if (TestCon_CA.CA_Band2 == "") TestCon_CA.CA_Band2 = CA_band;
                    else if (TestCon_CA.CA_Band3 == "") TestCon_CA.CA_Band3 = CA_band;
                    else if (TestCon_CA.CA_Band4 == "") TestCon_CA.CA_Band4 = CA_band;
                }
                /*
                else
                {
                    string exception = "here";
                }
                */
            }
        }

        public string setup_statusfile(string TRX_direction, TestCon Testcon)
        {
            string Direction = TRX_direction.Trim().ToUpper();

            List<string> Input_Ports_Ref = new List<string>();
            List<string> Output_Ports_Ref = new List<string>();
            List<int> Port_numbers = new List<int>();

            if (Direction.Contains("TX")|| Direction.Contains("TRX")){ Direction = "TX"; }
            else if(Direction.Contains("RX")){ Direction = "RX"; }
            else{ Direction = "TX"; }

            foreach (string InputPort in Split_Testcon_Port(Testcon.Input_Port))
            {
                foreach (string OutputPort in Split_Testcon_Port(Testcon.Output_Port))
                {
                    Input_Ports_Ref.Add(InputPort);
                    Output_Ports_Ref.Add(OutputPort);
                }
            }

            Input_Ports_Ref = Input_Ports_Ref.Distinct().ToList();
            Output_Ports_Ref = Output_Ports_Ref.Distinct().ToList();

            foreach (string port in Input_Ports_Ref)
            {
                if (Globals.Spara_config_INFO.Dic_PortNum.ContainsKey(Find_Available_port(port)))
                {
                    Port_numbers.Add(Globals.Spara_config_INFO.Dic_PortNum[Find_Available_port(port)]);
                }
            }
            foreach (string port in Output_Ports_Ref)
            {
                if (Globals.Spara_config_INFO.Dic_PortNum.ContainsKey(Find_Available_port(port)))
                {
                    Port_numbers.Add(Globals.Spara_config_INFO.Dic_PortNum[Find_Available_port(port)]);
                }
            }

            Port_numbers.Sort();

            string Matched_Input = "Not_found";
            string Matched_Output = "Not_found";

            bool found_status_file = false;
            string final_status = "";
            List<string> found_status_files = new List<string>();

            if (Globals.Spara_config_INFO.Dic_PortNum.ContainsKey(Find_Available_port(Input_Ports_Ref[0])))
            {
                Matched_Input = Globals.Spara_config_INFO.Dic_PortDefinition[Find_Available_port(Input_Ports_Ref[0])];
            }

            if (Globals.Spara_config_INFO.Dic_PortNum.ContainsKey(Find_Available_port(Output_Ports_Ref[0])))
            {
                Matched_Output = Globals.Spara_config_INFO.Dic_PortDefinition[Find_Available_port(Output_Ports_Ref[0])];
            }

            if (Matched_Input != "Not_found" && Matched_Output != "Not_found")
            {
                foreach (string Status_key in Globals.Spara_config_INFO.Status_config.Keys)
                {
                    if (Status_key.Contains(Matched_Input) && Status_key.Contains(Matched_Output))
                    {
                        found_status_file = true;
                        found_status_files = Globals.Spara_config_INFO.Status_config[Status_key];
                        final_status = pick_status(found_status_files, Port_numbers, Testcon, Direction);

                        if (final_status == "T.B.D")
                        {
                            List<string> candidate_status = new List<string>();

                            foreach (string Status_key2 in Globals.Spara_config_INFO.Status_config.Keys)
                            {
                                if (Status_key2.Contains(Matched_Input) || Status_key2.Contains(Matched_Output))
                                {
                                    foreach (var item in Globals.Spara_config_INFO.Status_config[Status_key2])
                                    {
                                        candidate_status.Add(item);
                                    }
                                }
                            }
                            found_status_file = true;
                            final_status = pick_status(candidate_status, Port_numbers, Testcon, Direction);
                        }

                        break;
                    }
                }
            }
            else
            {
                List<string> candidate_status = new List<string>();

                foreach (string Status_key in Globals.Spara_config_INFO.Status_config.Keys)
                {
                    if (Status_key.Contains(Matched_Input) || Status_key.Contains(Matched_Output))
                    {
                        foreach (var item in Globals.Spara_config_INFO.Status_config[Status_key])
                        {
                            candidate_status.Add(item);
                        }
                    }
                }

                found_status_file = true;
                final_status = pick_status(candidate_status, Port_numbers, Testcon, Direction);
                
            }

            return final_status;
        }

        private string pick_status(List<string> candidated_status, List<int> port_to_use, TestCon Testcondition, string direction)
        {
            string final_status = "T.B.D";
            Dictionary<string, float> Nominated_status = new Dictionary<string, float>();

            foreach (string each_status_f in candidated_status)
            {
                bool IsContain_All_input_port = false;
                bool IsContain_All_Freq = false;

                int start = each_status_f.IndexOf('(');
                int end = each_status_f.IndexOf(')');
                string temp = each_status_f.Substring(start, (end - start + 1));

                string temp_string = temp.Replace('(', ' ');
                temp_string = temp_string.Replace(')', ' ');
                string[] Port_set = temp_string.Trim().ToUpper().Split(',');

                int match_count = 0;
                float Status_frequency_range = 0f;

                foreach (int Port_num in port_to_use)
                {
                    foreach (string Port_compare in Port_set)
                    {
                        if (Port_num == Convert.ToInt32(Port_compare.Trim()))
                        {
                            match_count++;
                            break;
                        }
                    }
                }

                if (match_count == port_to_use.Count) { IsContain_All_input_port = true; }
                string[] Status_frequency = Globals.Spara_config_INFO.StatusFreqrange[each_status_f].Trim().Split('-');

                if (IsEmpty(Testcondition.Start_Freq))
                {
                    string Band_U = Testcondition.Band.Trim().ToUpper();
                    string Dir_U = Testcondition.Direction.Trim().ToUpper();
                    foreach (var item in Globals.IniFile.Frequency_table.Keys)
                    {
                        if (item.Contains(Band_U) && item.Contains(Dir_U))
                        {
                            string[] frequencies = Globals.IniFile.Frequency_table[item].Split(',');
                            Testcondition.Start_Freq = frequencies[0].Trim();
                            Testcondition.Stop_Freq = frequencies[1].Trim();
                            break;
                        }
                    }
                }


                if (Convert.ToSingle(Status_frequency[0]) <= Convert.ToSingle(Testcondition.Start_Freq) &&
                    Convert.ToSingle(Status_frequency[1]) >= Convert.ToSingle(Testcondition.Stop_Freq))
                {
                    IsContain_All_Freq = true;
                    Status_frequency_range = Convert.ToSingle(Status_frequency[1]) - Convert.ToSingle(Status_frequency[0]);
                }

                if (IsContain_All_input_port && IsContain_All_Freq)
                {
                    Nominated_status.Add(each_status_f, Status_frequency_range);
                }
            }

            float min_range = 99999f;

            string Range_ID = Get_BandRange(Testcondition);

            foreach (string status_name in Nominated_status.Keys)
            {
                if (min_range > Nominated_status[status_name] && status_name.Contains(direction) && status_name.Contains(Range_ID))
                {
                    min_range = Nominated_status[status_name];
                    final_status = status_name;
                }

                //if(Testcondition.Parameter == "ISO:ANT, InAct_ANT")
                //{
                //    if (min_range > Nominated_status[status_name] && status_name.Contains(direction) && status_name.Contains(Range_ID))
                //    {
                //        min_range = Nominated_status[status_name];
                //        final_status = status_name;
                //    }
                //}
                //else
                //{
                //    if (min_range > Nominated_status[status_name] && status_name.Contains(direction))
                //    {
                //        min_range = Nominated_status[status_name];
                //        final_status = status_name;
                //    }
                //}
            }

            if (Nominated_status.Keys.Count != 0 && final_status == "T.B.D")  //found candidate but need determin
            {
                final_status = Nominated_status.Keys.Last();
            }

            return final_status;
        }
       
        private string Get_BandRange(TestCon Testcon)
        {
            string[] Category = Testcon.Test_SpecID.Trim().Split('_');
            string BAND = Category[0].Trim();
            string DIRECTION = Testcon.Direction;
            string return_Range = "";

            if(!(BAND.ToUpper().Contains("B")||BAND.ToUpper().Contains("N"))) { BAND = "B" + BAND; }
            if (BAND.Contains('.')) BAND = BAND.Replace('.', 'P');

            foreach (var Band_Freq in Globals.IniFile.Frequency_table.Keys)
            {
                string[] compare_Band_Freq = Band_Freq.Split('_');
                if (compare_Band_Freq[2].Trim().ToUpper() == BAND && Band_Freq.Contains(DIRECTION))
                {
                    string[] temp = Globals.IniFile.Frequency_table[Band_Freq].Split(',');
                    if (Convert.ToSingle(temp[1].Trim()) < 1700f)
                    {
                        return_Range = "LMB";
                        break;
                    }
                    else if (Convert.ToSingle(temp[1].Trim()) < 2300f) //stop frequency is smaller than 2300MHz
                    {
                        return_Range = "MB";
                        break;
                    }
                    else if (Convert.ToSingle(temp[1].Trim()) > 2300f) //stop frequency is smaller than 2300MHz
                    {
                        return_Range = "HB";
                        break;
                    }
                }
            }

            return return_Range;

        }


        private bool IsEmpty(string Array)
        {
            bool Is_Empty = false;

            if (Array.Trim().ToUpper() == "") Is_Empty = true;
            if (Array.Trim().ToUpper() == "-") Is_Empty = true;
            if (Array.Trim().ToUpper() == " ") Is_Empty = true;

            return Is_Empty;
        }
        //Set Treeview
        public void Set_treeview_with_Expanded_Seq(Dictionary<string, Dictionary<string,List<TestCon>>> Global_ExpandedSpara_Seq)
        {
            TreeNode Root = new TreeNode("Spara_test");
            Root.Name = "Spara_test";
            Root.Checked = true;
            this.Root_Treeview = Root;

            List<string> filter_Param = new List<string>();
            List<string> filter_Temp = new List<string>();
            List<string> filter_InPort = new List<string>();
            List<string> filter_OutPort = new List<string>();

            int nodes_index = 0;

            foreach (var Band in Globals.Expaned_Spara_Seq.Keys)
            {
                TreeNode Band_Nodes = new TreeNode(Band);
                Band_Nodes.Name = Band;


                foreach (var Parameter in Globals.Expaned_Spara_Seq[Band].Keys)
                {
                    TreeNode Param_ID = new TreeNode(Parameter);
                    Param_ID.Checked = true;
                    Param_ID.Name = Parameter;

                    TreeNode Item_Node_RT = new TreeNode("25_Room");
                    TreeNode Item_Node_LT = new TreeNode("-30_Low");
                    TreeNode Item_Node_HT = new TreeNode("85_Hot");

                    Item_Node_RT.Name = "25_Room";
                    Item_Node_LT.Name = "-30_Low";
                    Item_Node_HT.Name = "85_Hot";

                    filter_Temp = menu_Add_WO_Duplication(filter_Temp, "25");
                    filter_Temp = menu_Add_WO_Duplication(filter_Temp, "-30");
                    filter_Temp = menu_Add_WO_Duplication(filter_Temp, "85");
                    filter_Param = menu_Add_WO_Duplication(filter_Param, Parameter);

                    foreach (TestCon TestSeq_TestCon in Globals.Expaned_Spara_Seq[Band][Parameter])
                    {
                        string Base_node_name = TestSeq_TestCon.Test_SpecID + "_" + TestSeq_TestCon.Test_Name;
                        TreeNode Testconditon = new TreeNode(TestSeq_TestCon.Test_SpecID + "_" + TestSeq_TestCon.Test_Name);
                        Testconditon.Checked = true;
                        Testconditon.Tag = TestSeq_TestCon;
                        Testconditon.Name = Base_node_name;

                        filter_InPort = menu_Add_WO_Duplication(filter_InPort, TestSeq_TestCon.Input_Port);
                        filter_OutPort = menu_Add_WO_Duplication(filter_OutPort, TestSeq_TestCon.Output_Port);

                        if (TestSeq_TestCon.Temperature.Contains("25"))
                        {   
                            Item_Node_RT.Nodes.Add(Testconditon);
                            Item_Node_RT.Checked = true;
                        }
                        else if(TestSeq_TestCon.Temperature.Contains("-30"))
                        {
                            Item_Node_LT.Nodes.Add(Testconditon);
                            Item_Node_LT.Checked = true;
                        }
                        else if (TestSeq_TestCon.Temperature.Contains("85"))
                        {
                            Item_Node_HT.Nodes.Add(Testconditon);
                            Item_Node_HT.Checked = true;
                        }
                    }
                    Param_ID.Nodes.Add(Item_Node_RT);
                    Param_ID.Nodes.Add(Item_Node_LT);
                    Param_ID.Nodes.Add(Item_Node_HT);

                    Band_Nodes.Nodes.Add(Param_ID);
                }

                Band_Nodes.Checked = true;
                Root.Nodes.Add(Band_Nodes);
                Root.Name = "Spara_test";

            }

            Root.Expand();
            this.treeView1.Nodes.Add(Root);
            Item_Adding_to_Checkedlistbox(this.Test_Field, filter_Param);
            Item_Adding_to_Checkedlistbox(this.Extream_Condition, filter_Temp);
            Item_Adding_to_Checkedlistbox(this.Test_Port_Input, filter_InPort);
            Item_Adding_to_Checkedlistbox(this.Test_Port_Output, filter_OutPort);

        }

        private void Item_Adding_to_Checkedlistbox(CheckedListBox ListBox_E, List<string> list)
        {
            foreach (var item in list)
            {
                ListBox_E.Items.Add(item, true);
            }

        }

        private void Listview_Header_initialize()
        {
            this.listView1.Columns.Add("Spec_ID", 100, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Param_ID", 100, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Temp", 60, HorizontalAlignment.Center);
            this.listView1.Columns.Add("In Port", 80, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Out Port", 80, HorizontalAlignment.Center);
            this.listView1.Columns.Add("f_Start", 80, HorizontalAlignment.Center);
            this.listView1.Columns.Add("f_Stop", 80, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Status_File", 200, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Limit_L", 60, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Limit_T", 60, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Limit_U", 60, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Node_ID", 30, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Node_Index", 30, HorizontalAlignment.Center);
            this.listView1.Columns.Add("Node_Fullpath", 30, HorizontalAlignment.Center);
        }

        public List<string> Split_Testcon_Port(string port_names_text)
        {
            string temp_string = port_names_text.Replace('[', ' ');
            temp_string = temp_string.Replace(']', ' ');
            string[] Subset_ports = temp_string.Trim().ToUpper().Split(',');

            List<string> Port_list = new List<string>();

            for (int i = 0; i < Subset_ports.Length; i++)
            {
                Port_list.Add(Subset_ports[i].Trim().ToUpper());
            }

            return Port_list;
        }
        public List<TestCon> Testcon_expansion_Temp_Port(TestCon clone_condition, List<TestCon> Seq_List, ref List<string> input_matched, ref List<string> output_matched)
        {
            string Input_ports = clone_condition.Input_Port.Trim().ToUpper();
            string Output_ports = clone_condition.Output_Port.Trim().ToUpper();

            bool EmptyInput_port = IsEmpty(Input_ports);
            bool EmptyOutput_port = IsEmpty(Output_ports);

            if (Input_ports.Contains("PRX_OUT1,2,3,4"))
            {
                Input_ports = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
            }
            else if(Output_ports.Contains("PRX_OUT1,2,3,4"))
            {
                Output_ports = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
            }

            if (EmptyInput_port || EmptyOutput_port)
            {
                string target_port = "";
                if (EmptyInput_port) target_port = Output_ports;
                if (EmptyOutput_port) target_port = Input_ports;

                foreach (string EachPort in Split_Testcon_Port(target_port))
                {
                    string Matched_Port = Find_Available_port(EachPort);
                    if (Matched_Port != "Not_Found")
                    {
                        if (EmptyOutput_port) input_matched = menu_Add_WO_Duplication(input_matched, Matched_Port);
                        if (EmptyInput_port) output_matched = menu_Add_WO_Duplication(output_matched, Matched_Port);

                        TestCon Testcon_Clone = clone_condition.Clone();
                        if (EmptyOutput_port) Testcon_Clone.Input_Port = Matched_Port;
                        if (EmptyInput_port) Testcon_Clone.Output_Port = Matched_Port;
                        Seq_List.Add(Testcon_Clone);
                    }
                }

            }
            else
            {
                foreach (string InputPort in Split_Testcon_Port(Input_ports))
                {
                    foreach (string OutputPort in Split_Testcon_Port(Output_ports))
                    {
                        string Matched_Input = Find_Available_port(InputPort);
                        string Matched_Output = Find_Available_port(OutputPort);

                        if (Matched_Input != "Not_Found" && Matched_Output != "Not_Found")
                        {
                            input_matched = menu_Add_WO_Duplication(input_matched, Matched_Input);
                            output_matched = menu_Add_WO_Duplication(output_matched, Matched_Output);

                            TestCon Testcon_Clone = clone_condition.Clone();
                            Testcon_Clone.Input_Port = Matched_Input;
                            Testcon_Clone.Output_Port = Matched_Output;
                            Seq_List.Add(Testcon_Clone);

                        }
                    }
                }
            }

            return Seq_List;

        }
        public string Find_Available_port(string target_port)
        {
            string target_PORT_UP = target_port.Trim().ToUpper();
            string represent_port = "";

            foreach (string Candiates in Globals.Spara_config_INFO.Dic_AvailablePort.Keys)
            {
                if (target_PORT_UP.Contains(Candiates))
                {
                    represent_port = Globals.Spara_config_INFO.Dic_AvailablePort[Candiates];
                }
            }

            if (represent_port == "")
            {
                represent_port = "Not_Found";
            }

            return represent_port;
        }

        public List<string> menu_Add_WO_Duplication(List<string> menu_list, string item)
        {
            menu_list.Add(item);
            menu_list = menu_list.Distinct().ToList();
            return menu_list;
        }

        bool Check_event_trigger = true;

        private void Tree_Nodes_Checked(object sender, TreeViewEventArgs e)
        {
            if(Check_event_trigger)
            {
                Check_event_trigger = false; //not allow critical conflict

                if (e.Node == this.Root_Treeview)
                {
                    Child_TreeNodes_Check(e.Node, e.Node.Checked);

                    for (int i = 0; i < this.Test_Field.Items.Count; i++)
                    {
                        this.Test_Field.SetItemChecked(i, e.Node.Checked);
                    }

                    for (int i = 0; i < this.Extream_Condition.Items.Count; i++)
                    {
                        this.Extream_Condition.SetItemChecked(i, e.Node.Checked);
                    }

                }
                else
                {
                    Child_TreeNodes_Check(e.Node, e.Node.Checked);
                    Parent_TreeNodes_Check(e.Node, e.Node.Checked);
                    this.Tree_Nodes_Selected(sender, e);
                }

                Check_event_trigger = true;
            }
        }

        private void Child_TreeNodes_Check(TreeNode node, bool Checked)
        {
            foreach (TreeNode Child_Node in node.Nodes)
            {
                Child_TreeNodes_Check(Child_Node, Checked);
            }
            node.Checked = Checked;
        }

        private void Parent_TreeNodes_Check(TreeNode node, bool Checked)
        {
            bool Claim_true = false;

            if(node.Parent != this.Root_Treeview)
            {
                foreach (TreeNode brothers in node.Parent.Nodes)
                {
                    if(brothers.Checked == true)
                    {
                        Claim_true = true;
                    }
                }

                if (Claim_true)
                {
                    node.Parent.Checked = true;
                }
                else
                {
                    node.Parent.Checked = Checked;
                }

                Parent_TreeNodes_Check(node.Parent, Checked);
            }
        }


        bool Listview_updated = false;
        bool checked_from_Listview = false;

        private void Set_TreeNodes_from_ListView(object sender, ItemCheckEventArgs e)
        {
            if (Listview_updated)
            {
                int index = e.Index;

                string TreeNode_Path = this.listView1.Items[index].SubItems[11].Text; //Treenode full path in listview1 control
                int TreeNode_index = Convert.ToInt32(this.listView1.Items[index].SubItems[12].Text); //Treenode full path in listview1 control
                string TreeNode_fullpath = this.listView1.Items[index].SubItems[13].Text; //Treenode full path in listview1 control

                TreeNode[] target_Trees = this.treeView1.Nodes.Find(TreeNode_Path, true);
                TreeNode found_targetND = new TreeNode();

                foreach (TreeNode Target_Node in target_Trees)
                {
                    if(Target_Node.Index == TreeNode_index && Target_Node.FullPath == TreeNode_fullpath)
                    {
                        //Check_event_trigger = false;
                        this.treeView1.SelectedNode = Target_Node;
                        if (e.NewValue == CheckState.Unchecked)
                        {
                            Target_Node.Checked = false;
                            found_targetND = Target_Node;
                        }
                        if (e.NewValue == CheckState.Checked)
                        {
                            Target_Node.Checked = true;
                            found_targetND = Target_Node;
                        }
                        //Check_event_trigger = true;
                    }
                }
            }
            checked_from_Listview = true;
        }

        private void Refresh_ListView(object sender, ItemCheckedEventArgs e)
        {
            if(checked_from_Listview)
            {
                checked_from_Listview = false;
            }
        }

        bool Select_event_trigger = true;
        public void Tree_Nodes_Selected(object sender, TreeViewEventArgs e)
        {
            
            if (e.Node.Parent == null) return;
            if (Select_event_trigger)
            {
                Select_event_trigger = false;

                TreeNode current_node = this.treeView1.SelectedNode;

                if(e.Node.Tag!=null)
                {
                    this.listView1.Items.Clear();
                    foreach (TreeNode Brothers in e.Node.Parent.Nodes)
                    {
                        ListView_Testcon_Update(Brothers);
                    }
                }
                else
                {
                    this.listView1.Items.Clear();
                    Child_TreeNodes_Select(e.Node);
                }

                this.treeView1.SelectedNode = current_node;
            }
            Select_event_trigger = true;
            
        }
        private void Child_TreeNodes_Select(TreeNode node)
        {
            foreach (TreeNode Child_Node in node.Nodes)
            {
                Child_TreeNodes_Select(Child_Node);
            }

            if(node.Tag!=null)
            {
                ListView_Testcon_Update(node);
            }
        }

        private void ListView_Testcon_Update(TreeNode node)
        {
            Listview_updated = false;
            TestCon Testcondition = (TestCon)node.Tag;
            String[] items =
            {
                Testcondition.Test_SpecID,
                Testcondition.Parameter,
                Testcondition.Temperature,
                Testcondition.Input_Port,
                Testcondition.Output_Port,
                Testcondition.Start_Freq,
                Testcondition.Stop_Freq,
                Testcondition.Status_file,
                Testcondition.Test_Limit_L,
                Testcondition.Test_Limit_Typ,
                Testcondition.Test_Limit_U,
                node.Name,
                Convert.ToString(node.Index),
                node.FullPath
            };
            ListViewItem review_item = new ListViewItem(items);
            review_item.Checked = node.Checked;
            if (!review_item.Checked)
            {
                review_item.BackColor = Color.DarkGray;
            }
            else
            {
                review_item.BackColor = Color.Aquamarine;
            }
            this.listView1.Items.Add(review_item);

            Listview_updated = true;
        }

        private void Test_Field_Select(object sender, ItemCheckEventArgs d)
        {
            string Selected_item = (string)this.Test_Field.SelectedItem;
            bool Checked_status = false;
            
            if (d.NewValue == CheckState.Unchecked) Checked_status = false;
            if (d.NewValue == CheckState.Checked) Checked_status = true;

            this.listView1.Items.Clear();

            if (this.Root_Treeview.Text != null && Selected_item != null)
            {
                foreach (TreeNode Band_Node in this.Root_Treeview.Nodes)
                {
                    foreach (TreeNode TestItem in Band_Node.Nodes)
                    {
                        if (TestItem.Text == Selected_item)
                        {
                            Check_event_trigger = false;
                            TestItem.Checked = Checked_status;
                            Child_TreeNodes_Check(TestItem, TestItem.Checked);
                            Parent_TreeNodes_Check(TestItem, TestItem.Checked);

                            Child_TreeNodes_Select(TestItem);
                            Check_event_trigger = true;
                        }
                    }
                }
            }
        }

        private void Temp_Field_Select(object sender, ItemCheckEventArgs d)
        {
            string Selected_item = (string)this.Extream_Condition.SelectedItem;
            bool Checked_status = false;

            if (d.NewValue == CheckState.Unchecked) Checked_status = false;
            if (d.NewValue == CheckState.Checked) Checked_status = true;

            this.listView1.Items.Clear();

            if (Selected_item == "25") Selected_item = "25_Room";
            if (Selected_item == "-30") Selected_item = "-30_Low";
            if (Selected_item == "85") Selected_item = "85_Hot";


            if (this.Root_Treeview.Text != null && Selected_item != null)
            {
                foreach (TreeNode Band_Node in this.Root_Treeview.Nodes)
                {
                    foreach (TreeNode TestItem in Band_Node.Nodes)
                    {
                        foreach (TreeNode Temp in TestItem.Nodes)
                        {
                            if (Temp.Text == Selected_item)
                            {
                                Check_event_trigger = false;
                                if(Temp.Parent.Checked == true)
                                {
                                    Temp.Checked = Checked_status;
                                    Child_TreeNodes_Check(Temp, Temp.Checked);
                                    Parent_TreeNodes_Check(Temp, Temp.Checked);
                                }
                                
                                //Child_TreeNodes_Select(Temp);

                                Check_event_trigger = true;
                            }
                        }
                    }
                }
            }
        }

        private void Generate_Plan_Click(object sender, EventArgs e)
        {
            //cancel Event sub-scribed status for memory

            this.Extream_Condition.ItemCheck -= this.Temp_Field_Select;
            this.Test_Field.ItemCheck -= this.Test_Field_Select;
            this.treeView1.AfterCheck -= this.Tree_Nodes_Checked;
            this.treeView1.AfterSelect -= this.Tree_Nodes_Selected;
            this.listView1.ItemCheck -= this.Set_TreeNodes_from_ListView;
            this.listView1.ItemChecked -= this.Refresh_ListView;

            Globals.Spara_TestCon.Clear(); //remove for memory
            Globals.Expaned_Spara_Seq.Clear(); //remove for memory
            Globals.Spara_TestTrigger_Count = 0;

            //Dictionary ( temp < band < status_group < list[item] ) 

            Dictionary<string, List<Spara_Trigger_Group>> Spara_Con_Room = new Dictionary<string, List<Spara_Trigger_Group>>();
            Dictionary<string, List<Spara_Trigger_Group>> Spara_Con_Low = new Dictionary<string, List<Spara_Trigger_Group>>();
            Dictionary<string, List<Spara_Trigger_Group>> Spara_Con_Hot = new Dictionary<string, List<Spara_Trigger_Group>>();

            //Globals.Spara_config_INFO.Group_Bands
            //Globals.Spara_config_INFO.Group_def_dic

            foreach (TreeNode Band_Node in this.Root_Treeview.Nodes)
            {
                string Band = Band_Node.Text;
                bool IsLoaded = false;
                //need dictionary for manipulate
                List<Spara_Trigger_Group> Band_triggers_room = new List<Spara_Trigger_Group>();
                List<Spara_Trigger_Group> Band_triggers_Low = new List<Spara_Trigger_Group>();
                List<Spara_Trigger_Group> Band_triggers_High = new List<Spara_Trigger_Group>();

                Dictionary<string, List<TestCon>> TestCon_List_Room = new Dictionary<string, List<TestCon>>();
                Dictionary<string, List<TestCon>> TestCon_List_Low = new Dictionary<string, List<TestCon>>();
                Dictionary<string, List<TestCon>> TestCon_List_Hot = new Dictionary<string, List<TestCon>>();

                foreach (TreeNode TestItem in Band_Node.Nodes)
                {
                    string Test_ID = TestItem.Text;

                    List<TestCon> Room_TC = new List<TestCon>();
                    List<TestCon> Low_TC = new List<TestCon>();
                    List<TestCon> Hot_TC = new List<TestCon>();

                    foreach (TreeNode Temp in TestItem.Nodes)
                    {
                        string Temp_ID = Temp.Text;

                        foreach (TreeNode Testcon in Temp.Nodes)
                        {
                            if (Testcon.Checked)
                            {
                                if (Temp_ID == "25_Room")
                                {
                                    Room_TC.Add((TestCon)Testcon.Tag);
                                    IsLoaded = true;
                                }
                                else if (Temp_ID == "-30_Low")
                                {
                                    Low_TC.Add((TestCon)Testcon.Tag);
                                    IsLoaded = true;
                                }
                                else if (Temp_ID == "85_Hot")
                                {
                                    Hot_TC.Add((TestCon)Testcon.Tag);
                                    IsLoaded = true;
                                }
                            }

                        }
                    }

                    Revise_TestCons(Test_ID, Room_TC, TestCon_List_Room);
                    Revise_TestCons(Test_ID, Low_TC, TestCon_List_Low);
                    Revise_TestCons(Test_ID, Hot_TC, TestCon_List_Hot);

                    //Assign_Port_Info(Room_TC);
                    //Assign_Port_Info(Low_TC);
                    //Assign_Port_Info(Hot_TC);

                    //TestCon_List_Room.Add(Test_ID, Room_TC);
                    //TestCon_List_Low.Add(Test_ID, Low_TC);
                    //TestCon_List_Hot.Add(Test_ID, Hot_TC);

                }

                if (IsLoaded) Order_BuildTrigger(Band, TestCon_List_Room, ref Band_triggers_room);
                if (IsLoaded) Order_BuildTrigger(Band, TestCon_List_Low, ref Band_triggers_Low);
                if (IsLoaded) Order_BuildTrigger(Band, TestCon_List_Hot, ref Band_triggers_High);

                if (Band_triggers_room.Count != 0)
                {
                    Assign_Trigger_Port_Info(Band_triggers_room);
                    Spara_Con_Room.Add(Band, Band_triggers_room);
                    Globals.Spara_TestTrigger_Count += Band_triggers_room.Count;
                }

                if (Band_triggers_Low.Count != 0)
                {
                    Assign_Trigger_Port_Info(Band_triggers_Low);
                    Spara_Con_Low.Add(Band, Band_triggers_Low);
                    Globals.Spara_TestTrigger_Count += Band_triggers_Low.Count;
                }

                if (Band_triggers_High.Count != 0)
                {
                    Assign_Trigger_Port_Info(Band_triggers_High);
                    Spara_Con_Hot.Add(Band, Band_triggers_High);
                    Globals.Spara_TestTrigger_Count += Band_triggers_High.Count;
                }
            }

            if (Spara_Con_Room.Count != 0) Globals.Spara_TestCon.Add("25",Spara_Con_Room);
            if (Spara_Con_Low.Count != 0) Globals.Spara_TestCon.Add("-30",Spara_Con_Low);
            if (Spara_Con_Hot.Count != 0) Globals.Spara_TestCon.Add("85",Spara_Con_Hot);

            Globals.SPara_Plan_Generate = true;
            this.Close();
        }

        private void Order_BuildTrigger(string Band, Dictionary<string, List<TestCon>> Testcon_order, ref List<Spara_Trigger_Group> trigger_group)
        {
            if (Testcon_order.ContainsKey("IL")) { Build_trigger(Band, Testcon_order["IL"], ref trigger_group); Testcon_order.Remove("IL"); }
            if (Testcon_order.ContainsKey("Input_VSWR")) { Build_trigger(Band, Testcon_order["Input_VSWR"], ref trigger_group); Testcon_order.Remove("Input_VSWR"); }
            if (Testcon_order.ContainsKey("Input_RL")) { Build_trigger(Band, Testcon_order["Input_RL"], ref trigger_group); Testcon_order.Remove("Input_RL"); }
            if (Testcon_order.ContainsKey("Output_RL")) { Build_trigger(Band, Testcon_order["Output_RL"], ref trigger_group); Testcon_order.Remove("Output_RL"); }
            if (Testcon_order.ContainsKey("Gain_Ripple")) { Build_trigger(Band, Testcon_order["Gain_Ripple"], ref trigger_group); Testcon_order.Remove("Gain_Ripple"); }
            if (Testcon_order.ContainsKey("TX_OOB_Gain")) { Build_trigger(Band, Testcon_order["TX_OOB_Gain"], ref trigger_group); Testcon_order.Remove("TX_OOB_Gain"); }
            if (Testcon_order.ContainsKey("RX_OOB_Gain")) { Build_trigger(Band, Testcon_order["RX_OOB_Gain"], ref trigger_group); Testcon_order.Remove("RX_OOB_Gain"); }

            if (Testcon_order.ContainsKey("K_factor")) { Build_trigger(Band, Testcon_order["K_factor"], ref trigger_group); Testcon_order.Remove("K_factor"); }
            if (Testcon_order.ContainsKey("MU_factor")) { Build_trigger(Band, Testcon_order["MU_factor"], ref trigger_group); Testcon_order.Remove("MU_factor"); }
            if (Testcon_order.ContainsKey("Group_Delay")) { Build_trigger(Band, Testcon_order["Group_Delay"], ref trigger_group); Testcon_order.Remove("Group_Delay"); }
            if (Testcon_order.ContainsKey("Phase_Delta")) { Build_trigger(Band, Testcon_order["Phase_Delta"], ref trigger_group); Testcon_order.Remove("Phase_Delta"); }

            if (Testcon_order.ContainsKey("ISO:ANT, InAct_ANT")) { Build_trigger(Band, Testcon_order["ISO:ANT, InAct_ANT"], ref trigger_group); Testcon_order.Remove("ISO:ANT, InAct_ANT"); }
            if (Testcon_order.ContainsKey("ISO:ANT, ANT")) { Build_trigger(Band, Testcon_order["ISO:ANT, ANT"], ref trigger_group); Testcon_order.Remove("ISO:ANT, ANT"); }
        
            if (Testcon_order.ContainsKey("ISO:TX, RX")) { Build_trigger(Band, Testcon_order["ISO:TX, RX"], ref trigger_group); Testcon_order.Remove("ISO:TX, RX"); } //never change order with below ISOs
            if (Testcon_order.ContainsKey("ISO:RX, InAct_RX")) { Build_trigger(Band, Testcon_order["ISO:RX, InAct_RX"], ref trigger_group); Testcon_order.Remove("ISO:RX, InAct_RX"); }

            if (Testcon_order.ContainsKey("ISO:TX, InAct_RX")) { Build_trigger(Band, Testcon_order["ISO:TX, InAct_RX"], ref trigger_group); Testcon_order.Remove("ISO:TX, InAct_RX"); } //never change order with below ISOs
            if (Testcon_order.ContainsKey("ISO:InAct_RX, InAct_RX")) { Build_trigger(Band, Testcon_order["ISO:InAct_RX, InAct_RX"], ref trigger_group); Testcon_order.Remove("ISO:InAct_RX, InAct_RX"); } 
            
            if (Testcon_order.ContainsKey("ISO:ASM, InAct_ANT")) { Build_trigger(Band, Testcon_order["ISO:ASM, InAct_ANT"], ref trigger_group); Testcon_order.Remove("ISO:ASM, InAct_ANT"); }
            if (Testcon_order.ContainsKey("ISO:ASM, InAct_RX")) { Build_trigger(Band, Testcon_order["ISO:ASM, InAct_RX"], ref trigger_group); Testcon_order.Remove("ISO:ASM, InAct_RX"); }
            if (Testcon_order.ContainsKey("ISO:ASM, ASM")) { Build_trigger(Band, Testcon_order["ISO:ASM, ASM"], ref trigger_group); Testcon_order.Remove("ISO:ASM, ASM"); }
            if (Testcon_order.ContainsKey("ISO:TX, ASM")) { Build_trigger(Band, Testcon_order["ISO:TX, ASM"], ref trigger_group); Testcon_order.Remove("ISO:TX, ASM"); }

            foreach (var item in Testcon_order.Keys)
            {
                if(Testcon_order[item].Count != 0) Build_trigger(Band, Testcon_order[item], ref trigger_group);
            }
        }

        private void Assign_Trigger_Port_Info(List<Spara_Trigger_Group> Group_List)
        {
            foreach (Spara_Trigger_Group each_Group in Group_List)
            {
                List<int> Ports = new List<int>();
                string Port_Seq = "";

                foreach (TestCon each_TC in each_Group.TestCon_List)
                {
                    string first = "";
                    string Last = "";

                    if(each_TC.Spara_ID.ToUpper().Contains("S") && each_TC.Spara_ID.Trim().Length < 5)
                    {
                        first = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 1);
                        Last = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 1, 1);
                    }
                    else if(each_TC.Spara_ID.ToUpper().Contains("GDEL") && each_TC.Spara_ID.Trim().Length < 8)
                    {
                        first = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 1);
                        Last = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 1, 1);
                    }
                    else
                    {
                        first = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 4, 2);
                        Last = each_TC.Spara_ID.Substring(each_TC.Spara_ID.Length - 2, 2);
                    }

                    Ports.Add(Convert.ToInt32(first));
                    Ports.Add(Convert.ToInt32(Last));
                    Ports = Ports.Distinct().ToList();

                    switch (each_TC.Parameter.Trim())
                    {
                        //case "IL":
                        case "Phase_Delta":
                            each_TC.Spara_Searchmethod = "MIN";
                            break;
                        default:
                            each_TC.Spara_Searchmethod = "MAX";
                            break;
                    }

                    if(each_TC.Test_Name.ToUpper().Contains("MIN")) each_TC.Spara_Searchmethod = "MIN";

                }

                Ports.Sort((x, y) => x.CompareTo(y));

                int i = 1;
                SortedList<int, int> Ports_List = new SortedList<int, int>();

                foreach (int item in Ports)
                {
                    Ports_List.Add(item, i);
                    i++;
                }

                foreach (int PortNum in Ports_List.Keys)
                {
                    if (Port_Seq == "")
                    {
                        Port_Seq = PortNum.ToString();
                    }
                    else
                    {
                        Port_Seq = Port_Seq + "," + PortNum.ToString();
                    }
                }

                each_Group.Ports_Sequence = Port_Seq;
                each_Group.Ports_Assigned = Ports_List;

            }
        }

        private void Revise_TestCons(string TestID, List<TestCon> TC_List, Dictionary<string, List<TestCon>> TestCon_List)
        {
            if (TC_List.Count == 0) return;

            Assign_Port_Info(TC_List);
            //Assign_Path_Info(TC_List);
            TestCon_List.Add(TestID, TC_List);
        }

        private void Assign_Path_Info(List<TestCon> TC_List)
        {
            foreach (TestCon each_TC in TC_List)
            {

            }
        }

        private void Assign_Port_Info(List<TestCon> TC_List)
        {
            foreach (TestCon each_TC in TC_List)
            {
                int InputPort_num = 0;
                int OutputPort_num = 0;

                if (!IsEmpty(each_TC.Input_Port)) InputPort_num = Globals.Spara_config_INFO.Dic_PortNum[each_TC.Input_Port];
                if (!IsEmpty(each_TC.Output_Port)) OutputPort_num = Globals.Spara_config_INFO.Dic_PortNum[each_TC.Output_Port];

                string SID_prefix = "S";

                bool Is_1D_IN = false;
                bool Is_1D_OUT = false;
                if (InputPort_num < 10 ) Is_1D_IN = true;
                if (OutputPort_num < 10) Is_1D_OUT = true;

                if (each_TC.Parameter.ToUpper().Contains("GROUP_DELAY"))
                {
                    SID_prefix = "GDEL";
                }
                string Spara_releation = find_Spara_relation(each_TC.Parameter, InputPort_num, OutputPort_num, Is_1D_IN, Is_1D_OUT);
                each_TC.Spara_ID = SID_prefix + Spara_releation;

                if (each_TC.Parameter.Contains("_RL") || each_TC.Parameter.Contains("ISO:") || each_TC.Parameter.Contains("IL")) //Return loss, Isolation express delta result without sign
                {
                    each_TC.Spara_ConvertSign = "ON";
                }
                else
                {
                    each_TC.Spara_ConvertSign = "OFF";
                }


            }
        }
        private string find_Spara_relation(string testID, int input_Num, int Output_Num, bool Is_1D_IN, bool Is_1D_OUT)
        {
            string SparaNumber = "";
            string InputNum_F = "";
            string OutputNum_F = "";

            string Test_ID = testID;
            if (Test_ID.Contains("REV_ISO:")) Test_ID = "REV_ISO";

            if (Is_1D_IN && Is_1D_OUT)
            {
                InputNum_F = input_Num.ToString("D1");
                OutputNum_F = Output_Num.ToString("D1");
            }
            else
            {
                if (Is_1D_IN && (Test_ID == "Input_RL" || Test_ID == "Output_RL" || Test_ID == "Input_VSWR") && !Is_1D_OUT)
                {
                    InputNum_F = input_Num.ToString("D1");
                    OutputNum_F = Output_Num.ToString("D2");
                }
                else
                {
                    InputNum_F = input_Num.ToString("D2");
                    OutputNum_F = Output_Num.ToString("D2");
                }              
            }

            switch (Test_ID)
            {
                case "Input_RL":
                case "Output_RL":
                case "Input_VSWR":
                    SparaNumber = InputNum_F + InputNum_F; //S11
                    break;
                case "REV_ISO":
                    SparaNumber = OutputNum_F + InputNum_F; //S12
                    break;
                default:
                    SparaNumber = OutputNum_F + InputNum_F; //S21
                    break;
            }

            return SparaNumber;
        }

        private string find_Trigger(string Band, TestCon TC, Spara_Trigger_Group Band_triggers, string Rev_input, string Rev_output)
        {
            string find_Group = "T.B.D";
            bool IsMatched = false;

            string Test_ID = TC.Parameter;
            string Status = TC.Status_file;
            string TC_input = TC.Input_Port;
            string TC_output = TC.Output_Port;
            string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID);

            if (Band_triggers.Status_File == Status && Band_triggers.Group_TYP == SubGroup)
            {
                IsMatched = Band_triggers.Port_matching(TC_input, TC_output, Rev_input, Rev_output);
                
                if (IsMatched)
                {
                    foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                    {
                        if (Each_TC.Parameter == Test_ID &&
                            Each_TC.Input_Port == TC_input &&
                            Each_TC.Output_Port == TC_output &&
                            Each_TC.Start_Freq == TC.Start_Freq &&
                            Each_TC.Stop_Freq == TC.Stop_Freq &&
                            Each_TC.Test_SpecID == TC.Test_SpecID &&
                            Each_TC.Temperature == TC.Temperature)
                        {
                            find_Group = "IS_EXIST";
                        }
                        else
                        {
                            find_Group = "FIND";
                        }
                    }
                }
            }

            return find_Group;
        }

        private string find_Trigger(string Band, TestCon TC, Spara_Trigger_Group Band_triggers, string Tunable_OUT, bool Is_CA)
        {
            string find_Group = "T.B.D";

            string Test_ID = TC.Parameter;
            string Status = TC.Status_file;
            string TC_input = TC.Input_Port;
            string TC_output = TC.Output_Port;
            string Input_Type = Globals.Spara_config_INFO.Input_Match(Band, Test_ID);
            string Output_Type = Globals.Spara_config_INFO.Output_Match(Band, Test_ID);
            string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID, Is_CA);

            if (Band_triggers.Status_File == Status && Band_triggers.Group_TYP == SubGroup && Band_triggers.CA_Case == Tunable_OUT)
            {
                if (Band_triggers.Port_matching(TC_input, TC_output, Input_Type, Output_Type))
                {
                    foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                    {
                        if (Each_TC.Parameter == Test_ID &&
                            Each_TC.Input_Port == TC_input &&
                            Each_TC.Output_Port == TC_output &&
                            Each_TC.Start_Freq == TC.Start_Freq &&
                            Each_TC.Stop_Freq == TC.Stop_Freq &&
                            Each_TC.Test_SpecID == TC.Test_SpecID &&
                            Each_TC.Temperature == TC.Temperature)
                        {
                            find_Group = "IS_EXIST";
                        }
                        else
                        {
                            find_Group = "FIND";
                        }
                    }
                }
            }

            return find_Group;
        }

        private string find_Trigger_RX(string Band, TestCon TC, Spara_Trigger_Group Band_triggers, string Rev_input, string Rev_output, string Tunable, bool Is_CA)
        {
            string find_Group = "T.B.D";
            bool IsMatched = false;

            string Test_ID = TC.Parameter;
            string Status = TC.Status_file;
            string TC_input = TC.Input_Port;
            string TC_output = TC.Output_Port;
            string RX_GainMode = TC.LNA_Gain_Mode;
            string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID, RX_GainMode, Is_CA);

            if(!Is_CA)
            {
                if (Band_triggers.Status_File == Status && Band_triggers.Group_TYP == SubGroup && Band_triggers.RX_Mode == RX_GainMode)
                {
                    IsMatched = Band_triggers.Port_matching(TC_input, TC_output, Rev_input, Rev_output);

                    if (IsMatched)
                    {
                        foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                        {
                            if (Each_TC.Parameter == Test_ID &&
                                Each_TC.Input_Port == TC_input &&
                                Each_TC.Output_Port == TC_output &&
                                Each_TC.Start_Freq == TC.Start_Freq &&
                                Each_TC.Stop_Freq == TC.Stop_Freq &&
                                Each_TC.Test_SpecID == TC.Test_SpecID &&
                                Each_TC.Temperature == TC.Temperature &&
                                Each_TC.LNA_Gain_Mode == TC.LNA_Gain_Mode &&
                                Each_TC.Test_Name == TC.Test_Name)
                            {
                                find_Group = "IS_EXIST";
                            }
                            else
                            {
                                find_Group = "FIND";
                            }
                        }
                    }
                    else if (Test_ID.ToUpper().Contains("_RL") || Test_ID == "ISO:ANT, ANT" || Test_ID == "REV_ISO:RX, ANT")
                    {
                        IsMatched = Band_triggers.Port_matching(TC_output, TC_input, Rev_input, Rev_output);

                        if (IsMatched)
                        {
                            foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                            {
                                if (Each_TC.Parameter == Test_ID &&
                                    Each_TC.Input_Port == TC_input &&
                                    Each_TC.Output_Port == TC_output &&
                                    Each_TC.Start_Freq == TC.Start_Freq &&
                                    Each_TC.Stop_Freq == TC.Stop_Freq &&
                                    Each_TC.Test_SpecID == TC.Test_SpecID &&
                                    Each_TC.Temperature == TC.Temperature &&
                                    Each_TC.LNA_Gain_Mode == TC.LNA_Gain_Mode)
                                {
                                    find_Group = "IS_EXIST";
                                }
                                else
                                {
                                    find_Group = "FIND";
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                if (Band_triggers.Status_File == Status && Band_triggers.Group_TYP == SubGroup && Band_triggers.RX_Mode == RX_GainMode && Band_triggers.CA_Case == Tunable)
                {
                    IsMatched = Band_triggers.Port_matching(TC_input, TC_output, Rev_input, Rev_output);

                    if (IsMatched)
                    {
                        foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                        {
                            if (Each_TC.Parameter == Test_ID &&
                                Each_TC.Input_Port == TC_input &&
                                Each_TC.Output_Port == TC_output &&
                                Each_TC.Start_Freq == TC.Start_Freq &&
                                Each_TC.Stop_Freq == TC.Stop_Freq &&
                                Each_TC.Test_SpecID == TC.Test_SpecID &&
                                Each_TC.Temperature == TC.Temperature &&
                                Each_TC.LNA_Gain_Mode == TC.LNA_Gain_Mode)
                            {
                                find_Group = "IS_EXIST";
                            }
                            else
                            {
                                find_Group = "FIND";
                            }
                        }
                    }
                    else if (Test_ID.ToUpper().Contains("_RL") || Test_ID == "ISO:ANT, ANT" || Test_ID == "REV_ISO:RX, ANT")
                    {
                        IsMatched = Band_triggers.Port_matching(TC_output, TC_input, Rev_input, Rev_output);

                        if (IsMatched)
                        {
                            foreach (TestCon Each_TC in Band_triggers.TestCon_List)
                            {
                                if (Each_TC.Parameter == Test_ID &&
                                    Each_TC.Input_Port == TC_input &&
                                    Each_TC.Output_Port == TC_output &&
                                    Each_TC.Start_Freq == TC.Start_Freq &&
                                    Each_TC.Stop_Freq == TC.Stop_Freq &&
                                    Each_TC.Test_SpecID == TC.Test_SpecID &&
                                    Each_TC.Temperature == TC.Temperature &&
                                    Each_TC.LNA_Gain_Mode == TC.LNA_Gain_Mode)
                                {
                                    find_Group = "IS_EXIST";
                                }
                                else
                                {
                                    find_Group = "FIND";
                                }
                            }
                        }
                    }
                }
            }

            return find_Group;
        }

        private void Create_NewTrigger(string Band, TestCon Target_TC, ref List<Spara_Trigger_Group> Band_triggers)
        {
            string Test_ID = Target_TC.Parameter;
            string TC_input = Target_TC.Input_Port;
            string TC_output = Target_TC.Output_Port;

            if (Target_TC.Status_file != "T.B.D")
            {
                string Group = Globals.Spara_config_INFO.Get_Group(Band);
                string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID);

                Spara_Trigger_Group Spara_Trigger = new Spara_Trigger_Group();

                Spara_Trigger.Status_File = Target_TC.Status_file;
                Spara_Trigger.Group_TYP = SubGroup;

                Spara_Trigger.Direction = "";
                Spara_Trigger.Tempearature = Target_TC.Temperature;
                Spara_Trigger.Test_Input = TC_input;
                Spara_Trigger.Test_Output = TC_output;
                Spara_Trigger.TestCon_List.Add(Target_TC);
                Spara_Trigger.RX_Mode = Target_TC.LNA_Gain_Mode;
                Set_ASMPort(Target_TC.Input_Port, Target_TC.Output_Port, ref Spara_Trigger);
                Band_triggers.Add(Spara_Trigger);
            }
        }

        private void Create_NewTrigger_RX(string Band, TestCon Target_TC, ref List<Spara_Trigger_Group> Band_triggers, string Tunable_CA, bool Is_CA)
        {
            string Test_ID = Target_TC.Parameter;
            string TC_input = Target_TC.Input_Port;
            string TC_output = Target_TC.Output_Port;
            string TC_GainMode = Target_TC.LNA_Gain_Mode;

            if (Target_TC.Status_file != "T.B.D")
            {
                string Group = Globals.Spara_config_INFO.Get_Group(Band);
                string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID, TC_GainMode, Is_CA);

                Spara_Trigger_Group Spara_Trigger = new Spara_Trigger_Group();

                Spara_Trigger.Status_File = Target_TC.Status_file;
                Spara_Trigger.Group_TYP = SubGroup;

                Spara_Trigger.Direction = "";
                Spara_Trigger.Tempearature = Target_TC.Temperature;
                Spara_Trigger.Test_Input = TC_input;
                Spara_Trigger.Test_Output = TC_output;
                Spara_Trigger.TestCon_List.Add(Target_TC);
                Spara_Trigger.RX_Mode = Target_TC.LNA_Gain_Mode;
                Set_ASMPort(Target_TC.Input_Port, Target_TC.Output_Port, ref Spara_Trigger);
                Spara_Trigger.CA_Case = Tunable_CA;
                Band_triggers.Add(Spara_Trigger);
            }
        }

        private void Create_NewTrigger(string Band, TestCon Target_TC, ref List<Spara_Trigger_Group> Band_triggers, string Tunable_CA, bool IS_CA)
        {
            string Test_ID = Target_TC.Parameter;
            string TC_input = Target_TC.Input_Port;
            string TC_output = Target_TC.Output_Port;

            if (Target_TC.Status_file != "T.B.D")
            {
                string Group = Globals.Spara_config_INFO.Get_Group(Band);
                string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID, IS_CA);

                Spara_Trigger_Group Spara_Trigger = new Spara_Trigger_Group();

                Spara_Trigger.Status_File = Target_TC.Status_file;
                Spara_Trigger.Group_TYP = SubGroup;

                Spara_Trigger.Direction = "";
                Spara_Trigger.Tempearature = Target_TC.Temperature;
                Spara_Trigger.Test_Input = TC_input;
                Spara_Trigger.Test_Output = TC_output;
                Spara_Trigger.TestCon_List.Add(Target_TC);
                Spara_Trigger.RX_Mode = Target_TC.LNA_Gain_Mode;
                Set_ASMPort(Target_TC.Input_Port, Target_TC.Output_Port, ref Spara_Trigger);
                Spara_Trigger.CA_Case = Tunable_CA;

                Band_triggers.Add(Spara_Trigger);
            }
        }
        private bool Build_trigger(string Band, List<TestCon> TC_list, ref List<Spara_Trigger_Group> Band_triggers)
        {
            bool IsAssigned = false;
            TestCon Missing_part = new TestCon();

            bool Found_Group = false;
            List<string> Ports_Combination = new List<string>();

            if (TC_list.Count == 0) return IsAssigned;

            foreach (TestCon Target_TC in TC_list)
            {
                Ports_Combination.Add(Target_TC.Input_Port + "," + Target_TC.Output_Port);
                Ports_Combination = Ports_Combination.Distinct().ToList();
            }

            if(TC_list[0].Direction == "RX")
            {
                IsAssigned = Build_trigger_RX(Band, TC_list, ref Band_triggers);
                return IsAssigned;
            }


            //add "ISO:TX, ASM"


            if (TC_list[0].Parameter == "ISO:ASM, ASM")
            {
                foreach (var item in make_ActToAct_ASM_Group_order(Band, Ports_Combination, TC_list))
                {
                    Band_triggers.Add(item);
                }
                IsAssigned = true;
            }
            else if (TC_list[0].Parameter == "ISO:TX, RX") //case to determin single TRX isolation and CA case
            {
                foreach (TestCon item in TC_list)
                {
                    if (item.Status_file == "T.B.D") continue;
                    string Tuanble_file = "";

                    bool Is_CA = Define_CA_case(item, ref Tuanble_file);

                    List<string> IsFind_list = new List<string>();
                    foreach (Spara_Trigger_Group Each_Trigger in Band_triggers)
                    {
                        string IsFind = find_Trigger(Band, item, Each_Trigger, Tuanble_file, Is_CA);
                        if (IsFind == "FIND")
                        {
                            Each_Trigger.TestCon_List.Add(item);
                            if (Each_Trigger.CA_Case == "") Each_Trigger.CA_Case = Tuanble_file;
                            IsFind_list.Add(IsFind);
                        }
                    }

                    if (IsFind_list.Count == 0)
                    {
                        Create_NewTrigger(Band, item, ref Band_triggers, Tuanble_file, Is_CA);
                    }

                }
                IsAssigned = true;
            }
            //else if (TC_list[0].Parameter == "ISO:RX, InAct_RX" || TC_list[0].Parameter == "ISO:ANT, InAct_ANT")
            else if (TC_list[0].Parameter == "ISO:ANT, InAct_ANT")
            {
                foreach (TestCon item in TC_list)
                {
                    if (item.Status_file == "T.B.D") continue;
                    if (item.Input_Port.Trim() == item.Output_Port.Trim()) continue;

                    List<string> IsFind_list = new List<string>();
                    foreach (Spara_Trigger_Group Each_Trigger in Band_triggers)
                    {
                        string IsFind = find_Trigger(Band, item, Each_Trigger, "REVERSE", "FALSE");
                        if (IsFind == "FIND")
                        {
                            Each_Trigger.TestCon_List.Add(item);
                            IsFind_list.Add(IsFind);
                        }
                    }

                    if (IsFind_list.Count == 0)
                    {
                        IsFind_list.Clear();
                        string Input_Type = Globals.Spara_config_INFO.Input_Match(Band, item.Parameter);
                        string Output_Type = Globals.Spara_config_INFO.Output_Match(Band, item.Parameter);

                        foreach (Spara_Trigger_Group Each_Trigger in Band_triggers)
                        {
                            string IsFind = find_Trigger(Band, item, Each_Trigger, Input_Type, Output_Type);
                            if (IsFind == "FIND")
                            {
                                Each_Trigger.TestCon_List.Add(item);
                                IsFind_list.Add(IsFind);
                                break;
                            }
                        }

                        if (IsFind_list.Count == 0)
                        {
                            Create_NewTrigger(Band, item, ref Band_triggers);

                            //if (TC_list[0].Parameter == "ISO:RX, InAct_RX")
                            //{
                            //    string Tuanble_file = "";
                            //    bool Is_CA = Define_CA_case(item, ref Tuanble_file);
                            //    Create_NewTrigger(Band, item, ref Band_triggers, Tuanble_file, Is_CA);
                            //}
                            //else
                            //{
                            //    Create_NewTrigger(Band, item, ref Band_triggers);
                            //}
                        }
                    }
                }

                IsAssigned = true;
            }
            else if (TC_list[0].Parameter == "ISO:ASM, InAct_ANT" || TC_list[0].Parameter == "ISO:RX, InAct_RX")  //Allow duplication
            {
                foreach (TestCon item in TC_list)
                {
                    if (item.Status_file == "T.B.D") continue;
                    if (item.Input_Port.Trim() == item.Output_Port.Trim()) continue;
                    if (TC_list[0].Parameter == "ISO:RX, InAct_RX" && BAND_RXout_Check(item, "INPUT")) continue;

                    string Tuanble_file = Get_Tunable_with_OUTport(item, item.Input_Port);
                    if (TC_list[0].Parameter == "ISO:ASM, InAct_ANT") Tuanble_file = "";

                    List<string> IsFind_list = new List<string>();
                    foreach (Spara_Trigger_Group Each_Trigger in Band_triggers)
                    {
                        string Input_Type = Globals.Spara_config_INFO.Input_Match(Band, item.Parameter);
                        string Output_Type = Globals.Spara_config_INFO.Output_Match(Band, item.Parameter);

                        string IsFind = find_Trigger(Band, item, Each_Trigger, Input_Type, Output_Type);
                        if (IsFind == "FIND")
                        {   
                            Each_Trigger.TestCon_List.Add(item);
                            if (Each_Trigger.CA_Case == "") Each_Trigger.CA_Case = Tuanble_file;
                            IsFind_list.Add(IsFind);
                        }
                    }

                    if (IsFind_list.Count == 0)
                    {
                        if (TC_list[0].Parameter == "ISO:RX, InAct_RX")
                        {
                            Create_NewTrigger(Band, item, ref Band_triggers, Tuanble_file, false);
                        }
                        else
                        { 
                            Create_NewTrigger(Band, item, ref Band_triggers); 
                        }
                    }

                }
                IsAssigned = true;
            }
            else if (TC_list[0].Parameter == "ISO:TX, ASM")
            {
                foreach (TestCon Target_TC in TC_list) //Not Allow duplication && swap port to meet definition (for Out_RL)
                {
                    string Test_ID = Target_TC.Parameter;
                    List<TestCon> expanded_TC = new List<TestCon>();

                    if (Convert.ToSingle(Target_TC.Start_Freq) >= 1600f && Convert.ToSingle(Target_TC.Stop_Freq) <= 2800f) //it means testcon belong to LTE/NR "InBand"
                    {
                        //Check How Many CA bands RX..
                        string Band_string = Target_TC.Band.Trim();
                        Band_string = ((Band_string == "40") ? "40F" : Band_string);
                        Band_string = ((Band_string == "41") ? "41F" : Band_string);
                        if (Band_string.Contains('.')) Band_string = Band_string.Replace('.', 'P');

                        string Band_ID = "CA_TX_B" + Band_string;
                        string Main_BAND = ((Band_string.ToUpper().Contains("N")) ? Band_string : "B" + Band_string);

                        List<string> RX_CA_Band_Freq = new List<string>();
                        List<string> RX_CA_Band_Name = new List<string>();

                        foreach (string Band_key in Globals.IniFile.Band_CAs.Keys)
                        {
                            if (Band_key == Band_ID)
                            {
                                foreach (string RX_Bands in Globals.IniFile.Band_CAs[Band_key])
                                {
                                    if (Globals.IniFile.Frequency_table.ContainsKey("FREQ_RX_" + RX_Bands))
                                    {
                                        RX_CA_Band_Freq.Add(Globals.IniFile.Frequency_table["FREQ_RX_" + RX_Bands]);
                                        RX_CA_Band_Name.Add("FREQ_RX_" + RX_Bands);
                                    }
                                }
                            }
                        }

                        int Clone_Count = RX_CA_Band_Name.Count();

                        for (int i = 0; i < Clone_Count; i++)
                        {
                            string[] frequency = RX_CA_Band_Freq[i].Split(',');

                            TestCon NewRX_TestCon = Target_TC.Clone();
                            NewRX_TestCon.Start_Freq = frequency[0].Trim();
                            NewRX_TestCon.Stop_Freq = frequency[1].Trim();
                            NewRX_TestCon.Test_Name = NewRX_TestCon.Test_Name + " [" + RX_CA_Band_Name[i] + "]";
                            expanded_TC.Add(NewRX_TestCon);

                        }

                        if (Clone_Count == 0)
                        {
                            double Testcon_Start_F = 200000;
                            double Testcon_Stop_F = 0;
                            double InBand_Start_F = Convert.ToDouble(Target_TC.Start_Freq);
                            double InBand_Stop_F = Convert.ToDouble(Target_TC.Stop_Freq);

                            foreach (string Band_Frequency in Globals.IniFile.Frequency_table.Keys) //calculate exception frequency area
                            {
                                if(Band_Frequency.ToUpper().Contains(Main_BAND))
                                {
                                    string[] frequency = Globals.IniFile.Frequency_table[Band_Frequency].Split(',');
                                    if (Convert.ToDouble(frequency[0].Trim()) < Testcon_Start_F) Testcon_Start_F = Convert.ToDouble(frequency[0].Trim());
                                    if (Convert.ToDouble(frequency[1].Trim()) > Testcon_Stop_F) Testcon_Stop_F = Convert.ToDouble(frequency[1].Trim());
                                }
                            }

                            if (Testcon_Stop_F <= InBand_Start_F || Testcon_Start_F >= InBand_Stop_F)
                            {
                                expanded_TC.Add(Target_TC);
                            }
                            else if (Testcon_Start_F >= InBand_Start_F && Testcon_Stop_F <= InBand_Stop_F)
                            {
                                if(Band_ID.Contains("_B40A"))
                                {
                                    Testcon_Start_F = 2200d; //Exception for Lightning
                                    Testcon_Stop_F = 2485.5d; //Exception for Lightning
                                }

                                TestCon NewRX_TestCon = Target_TC.Clone();
                                NewRX_TestCon.Start_Freq = InBand_Start_F.ToString();
                                NewRX_TestCon.Stop_Freq = Testcon_Start_F.ToString();
                                expanded_TC.Add(NewRX_TestCon);

                                TestCon NewRX_TestCon2 = Target_TC.Clone();
                                NewRX_TestCon2.Start_Freq = Testcon_Stop_F.ToString();
                                NewRX_TestCon2.Stop_Freq = InBand_Stop_F.ToString();
                                expanded_TC.Add(NewRX_TestCon2);
                            }
                            else if(Testcon_Start_F <= InBand_Start_F && Testcon_Stop_F <= InBand_Stop_F)
                            {
                                TestCon NewRX_TestCon2 = Target_TC.Clone();
                                NewRX_TestCon2.Start_Freq = Testcon_Stop_F.ToString();
                                NewRX_TestCon2.Stop_Freq = InBand_Stop_F.ToString();
                                expanded_TC.Add(NewRX_TestCon2);
                            }
                            else if (Testcon_Start_F >= InBand_Start_F && Testcon_Stop_F >= InBand_Stop_F)
                            {
                                TestCon NewRX_TestCon = Target_TC.Clone();
                                NewRX_TestCon.Start_Freq = InBand_Start_F.ToString();
                                NewRX_TestCon.Stop_Freq = Testcon_Start_F.ToString();
                                expanded_TC.Add(NewRX_TestCon);
                            }
                            else if (Testcon_Start_F <= InBand_Start_F && Testcon_Stop_F >= InBand_Stop_F)
                            {
                                //Do nothing
                            }
                            
                        }

                    }
                    else
                    {
                        expanded_TC.Add(Target_TC);
                    }

                    foreach (TestCon eachTC in expanded_TC)
                    {
                        if (Band_triggers.Count != 0)
                        {
                            Found_Group = search_Trigger_group(Band, Test_ID, eachTC, ref Band_triggers);
                        }

                        if (!Found_Group)
                        {
                            Create_NewTrigger(Band, eachTC, ref Band_triggers);
                        }
                    }
                }
                IsAssigned = true;
            }
            else
            {
                if (TC_list[0].Parameter == "Gain_Ripple")
                {
                    List<TestCon> TC_Gain_list = Expand_TC_To_TXGain(TC_list);
                    TC_list = Expand_TC_with_SplitFreq(TC_list);
                    foreach (TestCon Gain_cons in TC_Gain_list)
                    {
                        TC_list.Add(Gain_cons);
                    }
                }


                foreach (TestCon Target_TC in TC_list) //Not Allow duplication && swap port to meet definition (for Out_RL)
                {
                    string Test_ID = Target_TC.Parameter;

                    if (!Test_ID.Contains("RL") && (Target_TC.Input_Port == Target_TC.Output_Port)) continue;

                    if (Band_triggers.Count != 0)
                    {
                        Found_Group = search_Trigger_group(Band, Test_ID, Target_TC, ref Band_triggers);
                    }

                    if (!Found_Group)
                    {
                        Create_NewTrigger(Band, Target_TC, ref Band_triggers);
                    }

                }
                IsAssigned = true;
            }

            return IsAssigned;
        }

        private bool BAND_RXout_Check(TestCon item, string port_position)
        {
            string port_Toinspect = "";
            string Port_COMPARE = "";

            string BAND = item.Band;
            if (!(BAND.ToUpper().Contains("B") || BAND.ToUpper().Contains("N"))) BAND = "B" + BAND;
            if (BAND.Contains('.')) BAND = BAND.Replace('.', 'P');


            if (port_position == "INPUT") port_Toinspect = item.Input_Port;
            if (port_position == "OUTPUT") port_Toinspect = item.Output_Port;

            string[] repos_Portstring = port_Toinspect.Split('_');          

            for (int i = 0; i < repos_Portstring.Length; i++)
            {
                Port_COMPARE = Port_COMPARE + repos_Portstring[i];
            }

            foreach (var OUT_Key in Globals.IniFile.Band_RXOUTs.Keys)
            {
                if(Port_COMPARE.Contains(OUT_Key))
                {
                    if(Globals.IniFile.Band_RXOUTs[OUT_Key].Contains(BAND))
                    {
                        return false;
                    }
                }
            }

            return true;
        }


        private bool IsVSWR(TestCon CurrentTest)
        {
            bool IsVSWR = false;

            if (CurrentTest.ANTIn_VSWR.Contains("3")) IsVSWR = true;
            if (CurrentTest.RXOut_VSWR.Contains("3")) IsVSWR = true;
            if (CurrentTest.ANTout_VSWR.Contains("3")) IsVSWR = true;
            if (CurrentTest.TXIn_VSWR.Contains("3")) IsVSWR = true;

            return IsVSWR;
        }

        private bool Build_trigger_RX(string Band, List<TestCon> TC_list, ref List<Spara_Trigger_Group> Band_triggers)
        {
            bool IsAssigned = false;
            List<TestCon> EXP_TC_list = new List<TestCon>();

            if (TC_list[0].Parameter == "Gain_Ripple" || TC_list[0].Parameter == "Group_Delay")
            {
                TC_list = Expand_TC_with_SplitFreq(TC_list);
            }
            else if (TC_list[0].Parameter == "Phase_Delta")
            {
                TC_list = Expand_TC_with_LMH(TC_list);
            }
            else if (TC_list[0].Parameter.Contains("RX_Gain_G"))
            {
                TC_list = Expand_TC_To_RXGain(TC_list);
            }

            EXP_TC_list = Expand_TC_with_GMode(TC_list);

            string Tuanble_file = "";
            bool Is_CA = false;

            if (true) //default RX
            {
                foreach (TestCon item in EXP_TC_list)
                {
                    if (item.Status_file == "T.B.D") continue;
                    if (item.Input_Port.Trim() == item.Output_Port.Trim()) continue;
                    if (IsVSWR(item)) continue;

                    Is_CA = Define_RX_CA_case(item, ref Tuanble_file);

                    if (Tuanble_file == "Missing_PORT") continue; //it means Port CA combination is not available

                    List<string> IsFind_list = new List<string>();
                    foreach (Spara_Trigger_Group Each_Trigger in Band_triggers)
                    {
                        string Input_Type = Globals.Spara_config_INFO.Input_Match(Band, item.Parameter);
                        string Output_Type = Globals.Spara_config_INFO.Output_Match(Band, item.Parameter);

                        string IsFind = find_Trigger_RX(Band, item, Each_Trigger, Input_Type, Output_Type, Tuanble_file, Is_CA);
                        if (IsFind == "FIND")
                        {
                            Each_Trigger.TestCon_List.Add(item);
                            IsFind_list.Add(IsFind);
                            break; //not allow duplication
                        }
                    }

                    if (IsFind_list.Count == 0)
                    {
                        Create_NewTrigger_RX(Band, item, ref Band_triggers, Tuanble_file, Is_CA);
                    }

                }
                IsAssigned = true;
            }
            return IsAssigned;
        }

        private List<TestCon> Expand_TC_with_SplitFreq(List<TestCon> TC_List)
        {
            List<TestCon> expaned_TC_list_by_FrequencyStep = new List<TestCon>();

            double StartF = Convert.ToSingle(TC_List[0].Start_Freq);
            double StopF = Convert.ToSingle(TC_List[0].Stop_Freq);
            double Range = StopF - StartF;
            double GuardBand = (Range > 20f) ? 2 : 1;
            double Step = (Range > 20f) ? 20 : 10;
            Step = (Range > 500f) ? 100 : Step;  //*Case of NR N77,N79
            GuardBand = (Range > 500f) ? 10 : GuardBand; //*Case of NR N77,N79

            SortedList<double, double> Frequencies = new SortedList<double, double>();

            double First_StartF = StartF + GuardBand;
            double First_StopF = StartF + Step;
            double Last_StartF = StopF - Step;
            double Last_StopF = StopF - GuardBand;

            Frequencies.Add(First_StartF, First_StopF);
            Frequencies.Add(Last_StartF, Last_StopF);

            while (Last_StartF > First_StopF)
            {
                First_StartF = First_StopF;
                First_StopF = First_StopF + Step;
                if(!Frequencies.ContainsKey(First_StartF)) Frequencies.Add(First_StartF, First_StopF);
            }

            foreach (TestCon item in TC_List)
            {
                foreach (var splited_Frequency in Frequencies.Keys)
                {
                    TestCon ExpTestCon_item = new TestCon();
                    ExpTestCon_item = item.Clone();
                    ExpTestCon_item.Start_Freq = Convert.ToString(splited_Frequency);
                    ExpTestCon_item.Stop_Freq = Convert.ToString(Frequencies[splited_Frequency]);
                    expaned_TC_list_by_FrequencyStep.Add(ExpTestCon_item);
                }
            }

            return expaned_TC_list_by_FrequencyStep;
        }

        private List<TestCon> Expand_TC_To_TXGain(List<TestCon> TC_List)
        {
            List<TestCon> expaned_TC_for_TXGain = new List<TestCon>();

            foreach (TestCon item in TC_List)
            {
                TestCon ExpTestCon_item = new TestCon();
                ExpTestCon_item = item.Clone();

                ExpTestCon_item.Test_Name = "TX_Gain_MAX";
                ExpTestCon_item.Parameter = "TX_Gain_MAX";
                ExpTestCon_item.Test_Limit_U = "30";
                ExpTestCon_item.Test_SpecID = "NEED_TEST";

                TestCon ExpTestCon_Min = new TestCon();
                ExpTestCon_Min = ExpTestCon_item.Clone();
                ExpTestCon_Min.Test_Name = "TX_Gain_MIN";
                ExpTestCon_Min.Parameter = "TX_Gain_MIN";
                ExpTestCon_Min.Test_Limit_U = "";
                ExpTestCon_Min.Test_Limit_L = "15";

                expaned_TC_for_TXGain.Add(ExpTestCon_item);
                expaned_TC_for_TXGain.Add(ExpTestCon_Min);
            }

            return expaned_TC_for_TXGain;
        }

        private List<TestCon> Expand_TC_To_RXGain(List<TestCon> TC_List)
        {
            List<TestCon> expaned_TC_for_RXGain = new List<TestCon>();

            foreach (TestCon item in TC_List)
            {
                expaned_TC_for_RXGain.Add(item);

                TestCon ExpTestCon_item = new TestCon();
                ExpTestCon_item = item.Clone();
                ExpTestCon_item.Test_Name = ExpTestCon_item.Test_Name.ToLower().Replace("gain","gain (min)");

                expaned_TC_for_RXGain.Add(ExpTestCon_item);                
            }

            return expaned_TC_for_RXGain;
        }

        private List<TestCon> Expand_TC_with_LMH(List<TestCon> TC_List)
        {
            List<TestCon> expaned_TC_list_LMH = new List<TestCon>();

            double Low_F = Convert.ToSingle(TC_List[0].Start_Freq);
            double High_F = Convert.ToSingle(TC_List[0].Stop_Freq);
            double Mid_F = Low_F + ((High_F - Low_F) / 2);

            SortedList<double, double> Frequencies = new SortedList<double, double>();

            Frequencies.Add(Low_F, Low_F);
            Frequencies.Add(Mid_F, Mid_F);
            Frequencies.Add(High_F, High_F);

            foreach (TestCon item in TC_List)
            {
                foreach (var splited_Frequency in Frequencies.Keys)
                {
                    TestCon ExpTestCon_item = new TestCon();
                    ExpTestCon_item = item.Clone();
                    ExpTestCon_item.Start_Freq = Convert.ToString(splited_Frequency);
                    ExpTestCon_item.Stop_Freq = Convert.ToString(Frequencies[splited_Frequency]);
                    expaned_TC_list_LMH.Add(ExpTestCon_item);
                }
            }

            return expaned_TC_list_LMH;
        }

        private List<TestCon> Expand_TC_with_GMode(List<TestCon> TC_List)
        {
            List<TestCon> expaned_TC_list_by_Gmode = new List<TestCon>();
            bool IsReconized_Gmode = false;

            foreach (TestCon EachTC in TC_List)
            {
                //EachTC.LNA_Gain_Mode
                foreach (string GainMode in GetGain_Mode(EachTC))
                {
                    TestCon NewClone_TC = new TestCon();
                    NewClone_TC = EachTC.Clone();
                    NewClone_TC.LNA_Gain_Mode = GainMode;
                    expaned_TC_list_by_Gmode.Add(NewClone_TC);
                }
            }

            return expaned_TC_list_by_Gmode;
        }

        private List<string> GetGain_Mode(TestCon EachTC)
        {
            List<string> GainMode_out = new List<string>();
            string text_GainMode = EachTC.LNA_Gain_Mode.Trim().ToUpper();
            text_GainMode = text_GainMode.Replace('[', ' ');
            text_GainMode = text_GainMode.Replace(']', ' ');
            text_GainMode = text_GainMode.Replace('(', ' ');
            text_GainMode = text_GainMode.Replace(')', ' ');

            if (text_GainMode.Contains("GX") || text_GainMode.Contains("GY"))
            {
                GainMode_out = Globals.IniFile.RX_GainModes;
                return GainMode_out;
            }

            string[] GMode_Array = text_GainMode.Trim().Split(',');

            for (int i = 0; i < GMode_Array.Length; i++)
            {
                if (IsMatched_GainMode(GMode_Array[i])) GainMode_out.Add(GMode_Array[i].Trim());
            }

            return GainMode_out;
        }

        private bool IsMatched_GainMode(string gain_mode)
        {
            bool IsMatched = false;

            foreach (string CompRXMode in Globals.IniFile.RX_GainModes)
            {
                if (gain_mode.ToUpper().Trim() == CompRXMode)
                {
                    IsMatched = true;
                    break;
                }
            }

            return IsMatched;
        }

        private string Get_Tunable_with_OUTport(TestCon item, string OUTPUT)
        {
            if (OUTPUT == "") return "";

            string BAND_Prefix = "B";
            if (item.Band.ToUpper().Contains("N")) BAND_Prefix = "";

            string TEMP_Main_Band = BAND_Prefix + item.Band;
            if (TEMP_Main_Band.ToUpper().Contains('.')) TEMP_Main_Band = TEMP_Main_Band.Replace('.', 'P');

            int Index_OUT = OUTPUT.IndexOf('_') + 1;
            string Output_port_revise = OUTPUT.Substring(Index_OUT);

            StringBuilder Tuable_f = new StringBuilder();
            Tuable_f.AppendFormat("{0}OUT{1}", TEMP_Main_Band.ToUpper(), Regex.Match(Output_port_revise, @"\d+").Value);
            
            string Tunable_OUT = Tuable_f.ToString();

            return Tunable_OUT;
        }


        private bool Define_CA_case(TestCon item, ref string Tunable)
        {
            bool Is_CA = false;

            List<string> CA_Bands = new List<string>();
            if (item.CA_Band2 != "") CA_Bands.Add(item.CA_Band2);
            if (item.CA_Band3 != "") CA_Bands.Add(item.CA_Band3);
            if (item.CA_Band4 != "") CA_Bands.Add(item.CA_Band4);

            StringBuilder Tuable_f = new StringBuilder();

            //Important Band name can't contain '.', it will be replaced as 'P'
            if (item.Band.Contains('.')) { item.Band = item.Band.Replace('.', 'P'); }

            if (item.Parameter == "ISO:RX, InAct_RX" && CA_Bands.Count == 0)
            {
                item.CA_Band2 = item.Band;
                CA_Bands.Add(item.CA_Band2);
            }

            string BAND_Prefix = "B";
            if (item.Band.ToUpper().Contains("N")) BAND_Prefix = "";

            string TEMP_Main_Band = BAND_Prefix + item.Band;

            int Index_OUT = item.Output_Port.IndexOf('_') + 1;
            string Output_port_revise = item.Output_Port.Substring(Index_OUT);

            if (CA_Bands.Count == 1 && (item.Band.Trim() == item.CA_Band2.Trim() || item.Band.Trim() == item.CA_Band3.Trim() || item.Band.Trim() == item.CA_Band4.Trim()))
            {
                Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(Output_port_revise, @"\d+").Value);
                Is_CA = false;
            }
            else if (CA_Bands.Count == 1)
            {
                List<string> Available_RXout = Get_Band_RXout(item.CA_OutputPort_List);
                string CA_output = item.Output_Port;
                string TXRX_output = "";

                if (Available_RXout.Count != 0)
                {
                    foreach (string RXout in Available_RXout)
                    {
                        if (RXout != CA_output) 
                        {
                            string[] TEMP = RXout.Trim().Split('_');
                            StringBuilder TEMP_string = new StringBuilder();

                            for (int i = 0; i < TEMP.Length; i++)
                            {
                                TEMP_string.AppendFormat(TEMP[i].Trim());
                            }

                            string TEMP_RXout = TEMP_string.ToString();

                            foreach (var OUT_Key in Globals.IniFile.Band_RXOUTs.Keys)
                            {
                                if(TEMP_RXout.Contains(OUT_Key))
                                {
                                    if (Globals.IniFile.Band_RXOUTs[OUT_Key].Contains(TEMP_Main_Band))
                                    {
                                        TXRX_output = RXout; break;
                                    }
                                }
                            }
                        }
                    }
                }
                int CA_Index = TXRX_output.IndexOf('_') + 1;
                TXRX_output = TXRX_output.Substring(CA_Index);
                CA_output = CA_output.Substring(Index_OUT);

                if (item.Parameter!= "ISO:TX, RX" && TXRX_output != "")
                {
                    Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(TXRX_output, @"\d+").Value);
                    Tuable_f.Append("."); //splitter
                }
                Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, CA_Bands[0].ToUpper(), Regex.Match(CA_output, @"\d+").Value);
                Is_CA = true;
            }
            else if (CA_Bands.Count != 0)
            {
                //List<string> Available_RXout = Get_Band_RXout(item.CA_OutputPort_List);
                List<string> Available_RXout = new List<string>();
                foreach (var Port_Descr in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                {
                    string Port_Define = Globals.Spara_config_INFO.Dic_PortDefinition[Port_Descr];
                    if (Port_Define == "RX_OUT") Available_RXout.Add(Port_Descr);
                }

                string TXRX_output = item.Output_Port;
                Available_RXout.Remove(TXRX_output);
                StringBuilder Occupied_OUTport = new StringBuilder();
                Occupied_OUTport.Append(Remove_UnderBar(TXRX_output));

                TXRX_output = TXRX_output.Substring(Index_OUT);
                Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(TXRX_output, @"\d+").Value);

                int count_find = 0;

                CA_Bands = Set_CA_Band_Order(CA_Bands);

                if (Available_RXout.Count != 0)
                {
                    for (int i = 0; i < CA_Bands.Count; i++)
                    {
                        bool found_port = false;

                        foreach (string RXout in Available_RXout)
                        {
                            string TEMP_RXout = Remove_UnderBar(RXout);

                            foreach (var OUT_Key in Globals.IniFile.Band_RXOUTs.Keys)
                            {
                                if (TEMP_RXout.Contains(OUT_Key) && !Occupied_OUTport.ToString().Contains(OUT_Key))
                                {
                                    if (Globals.IniFile.Band_RXOUTs[OUT_Key].Contains(Revise_BandPrefix(CA_Bands[i])) &&
                                        !Occupied_OUTport.ToString().Contains(OUT_Key) && !Occupied_OUTport.ToString().Contains(Revise_BandPrefix(CA_Bands[i])))
                                    {
                                        Tuable_f.Append(".");
                                        Tuable_f.AppendFormat("{0}OUT{1}", Revise_BandPrefix(CA_Bands[i]), Regex.Match(OUT_Key, @"\d+").Value);
                                        Occupied_OUTport.Append(Remove_UnderBar(TEMP_RXout));
                                        found_port = true;
                                        count_find++;
                                        break;
                                    }

                                    if (found_port) break;
                                }
                            }

                            if (found_port) break;

                        }
                    }
                }

                if (count_find == CA_Bands.Count)
                {
                    Is_CA = true;
                }
                else
                {
                    Tuable_f.Clear();
                    Tuable_f.AppendFormat("Missing_PORT");
                    Is_CA = false;
                }
            }

            Tunable = Tuable_f.ToString();

            return Is_CA;
        }

        private bool Define_RX_CA_case(TestCon item, ref string Tunable)
        {
            bool Is_CA = false;

            List<string> CA_Bands = new List<string>();
            if (item.CA_Band2 != "") CA_Bands.Add(item.CA_Band2);
            if (item.CA_Band3 != "") CA_Bands.Add(item.CA_Band3);
            if (item.CA_Band4 != "") CA_Bands.Add(item.CA_Band4);

            if (item.Band.Contains('.')) { item.Band = item.Band.Replace('.', 'P'); }

            if (item.CA_Band2 == "" && item.CA_Band3 == "" && item.CA_Band4 == "")
            {
                item.CA_Band2 = item.Band; CA_Bands.Add(item.CA_Band2);
            }

            string BAND_Prefix = "B";
            if (item.Band.ToUpper().Contains("N")) BAND_Prefix = "";
            string TEMP_Main_Band = BAND_Prefix + item.Band;

            if (item.Parameter == "ISO:ANT, ANT")
            {
                List<string> Available_RXout_ATA = new List<string>();

                foreach (var DefinedPorts in Globals.Spara_config_INFO.Dic_PortDefinition)
                {
                    if (DefinedPorts.Value.Contains("RX_OUT")) Available_RXout_ATA.Add(DefinedPorts.Key);
                }

                StringBuilder Tuable_ATA = new StringBuilder();
                Tuable_ATA.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(Available_RXout_ATA[0], @"\d+").Value);
                Tunable = Tuable_ATA.ToString();
                return Is_CA;
            }

            StringBuilder Tuable_f = new StringBuilder();

            int Index_OUT = item.Output_Port.IndexOf('_') + 1;
            string Output_port_revise = item.Output_Port.Substring(Index_OUT);

            if (CA_Bands.Count == 1 && (item.Band.Trim() == item.CA_Band2.Trim() || item.Band.Trim() == item.CA_Band3.Trim() || item.Band.Trim() == item.CA_Band4.Trim()))
            {
                Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(Output_port_revise, @"\d+").Value);
                Is_CA = false;
            }
            else if (CA_Bands.Count != 0)
            {
                //List<string> Available_RXout = Get_Band_RXout(item.CA_OutputPort_List);
                List<string> Available_RXout = new List<string>();
                foreach (var Port_Descr in Globals.Spara_config_INFO.Dic_PortDefinition.Keys)
                {
                    string Port_Define = Globals.Spara_config_INFO.Dic_PortDefinition[Port_Descr];
                    if (Port_Define == "RX_OUT") Available_RXout.Add(Port_Descr);
                }

                string TXRX_output = item.Output_Port;
                Available_RXout.Remove(TXRX_output);
                StringBuilder Occupied_OUTport = new StringBuilder();
                Occupied_OUTport.Append(Remove_UnderBar(TXRX_output));

                TXRX_output = TXRX_output.Substring(Index_OUT);
                Tuable_f.AppendFormat("{0}{1}OUT{2}", BAND_Prefix, item.Band.ToUpper(), Regex.Match(TXRX_output, @"\d+").Value);

                int count_find = 0;

                CA_Bands = Set_CA_Band_Order(CA_Bands);

                if (Available_RXout.Count != 0)
                {
                    for (int i = 0; i < CA_Bands.Count; i++)
                    {
                        bool found_port = false;

                        foreach (string RXout in Available_RXout)
                        {
                            string TEMP_RXout = Remove_UnderBar(RXout);

                            foreach (var OUT_Key in Globals.IniFile.Band_RXOUTs.Keys)
                            {
                                if (TEMP_RXout.Contains(OUT_Key) && !Occupied_OUTport.ToString().Contains(OUT_Key))
                                {
                                    if (Globals.IniFile.Band_RXOUTs[OUT_Key].Contains(Revise_BandPrefix(CA_Bands[i])) && 
                                        !Occupied_OUTport.ToString().Contains(OUT_Key) && !Occupied_OUTport.ToString().Contains(Revise_BandPrefix(CA_Bands[i])))
                                    {
                                        Tuable_f.Append(".");
                                        Tuable_f.AppendFormat("{0}OUT{1}", Revise_BandPrefix(CA_Bands[i]), Regex.Match(OUT_Key, @"\d+").Value);
                                        Occupied_OUTport.Append(Remove_UnderBar(TEMP_RXout));
                                        found_port = true;
                                        count_find++;
                                        break;
                                    }

                                    if (found_port) break;
                                }
                            }

                            if (found_port) break;

                        }
                    }
                }

                if (count_find == CA_Bands.Count)
                {
                    Is_CA = true;
                }
                else
                {
                    Tuable_f.Clear();
                    Tuable_f.AppendFormat("Missing_PORT");
                    Is_CA = false;
                }

            }

            Tunable = Tuable_f.ToString();

            return Is_CA;
        }

        private List<string> Set_CA_Band_Order(List<string> CA_BAND)
        {
            List<string> Sorted_BAND_order = new List<string>();
            Dictionary<string, int> TEMP_QUE = new Dictionary<string, int>();

            foreach (string BandID in CA_BAND)
            {
                int Matched_Count = 0;
                foreach (var OUT_Key in Globals.IniFile.Band_RXOUTs.Keys)
                {
                    if (Globals.IniFile.Band_RXOUTs[OUT_Key].Contains(Revise_BandPrefix(BandID)))
                    {
                        Matched_Count++;
                    }
                }
                TEMP_QUE.Add(BandID, Matched_Count);
            }

            int Min_Count = 10; //RX_OUT port in matched case 
            Dictionary<string, int> TEMP_QUE2 = new Dictionary<string, int>();
            
            do
            {
                TEMP_QUE2 = TEMP_QUE;
                Min_Count = 99; 

                foreach (var count in TEMP_QUE.Values)
                {
                    if (count < Min_Count) Min_Count = count;
                }

                foreach (var Band in TEMP_QUE.Keys)
                {
                    if(TEMP_QUE[Band] == Min_Count)
                    {
                        Sorted_BAND_order.Add(Band);
                        TEMP_QUE2.Remove(Band);
                        break;
                    }
                }

            } while (TEMP_QUE2.Count != 0);

            return Sorted_BAND_order;
        }

        private string Revise_BandPrefix(string band_num)
        {
            string Revised_BandID = band_num;

            string BAND_Prefix = "B";

            if (Revised_BandID.ToUpper().Contains("N")) BAND_Prefix = "";
            if (Revised_BandID.Contains('.')) { Revised_BandID = Revised_BandID.Replace('.', 'P'); }

            Revised_BandID = BAND_Prefix + Revised_BandID;

            return Revised_BandID;
        }
        private string Remove_UnderBar(string PortName)
        {
            string[] TEMP = PortName.Trim().Split('_');
            StringBuilder TEMP_string = new StringBuilder();

            for (int i = 0; i < TEMP.Length; i++)
            {
                TEMP_string.AppendFormat(TEMP[i].Trim());
            }

            string Removed_RXout = TEMP_string.ToString();

            return Removed_RXout;
        }

        private List<string> Get_Band_RXout(string CA_OUT_list)
        {
            List<string> result_ports = new List<string>();
            string Output_ports = CA_OUT_list.Trim().ToUpper();

            if (Output_ports.Contains("PRX_OUT1,2,3,4"))
            {
                Output_ports = "PRX_OUT1,PRX_OUT2,PRX_OUT3,PRX_OUT4";
            }

            foreach (string eachPort in Split_Testcon_Port(Output_ports))
            {
                if (Globals.Spara_config_INFO.Dic_AvailablePort[eachPort] != null)
                {
                    result_ports.Add(Globals.Spara_config_INFO.Dic_AvailablePort[eachPort]);
                }
            }

            return result_ports;
        }

        public List<Spara_Trigger_Group> make_ActToAct_ASM_Group_order(string Band, List<string> combination_list, List<TestCon> TC_list)
        {
            List<string> ANT_ports = new List<string>();
            List<string> ASM_ports = new List<string>();
            List<Spara_Trigger_Group> combinations = new List<Spara_Trigger_Group>();

            foreach (var item in Globals.Spara_config_INFO.Dic_PortDefinition)
            {
                if (item.Value.Contains("ANT_OUT"))
                {
                    ANT_ports.Add(item.Key);
                    continue;
                }
                if (item.Value.Contains("ASM"))
                {
                    ASM_ports.Add(item.Key);
                    continue;
                }
            }

            List<string[]> Combination = new List<string[]>();

            string[] ASMport = new string[ASM_ports.Count];

            int s = 0;

            foreach (var item in ASM_ports)
            {
                ASMport[s] = item;
                s++;
            }

            for (int i = 0; i < ASM_ports.Count; i++)
            {
                string[] Antenna = new string[ANT_ports.Count];

                for (int j = 0; j < ANT_ports.Count; j++)
                {
                    Antenna[j] = ASMport[j];
                }

                Combination.Add(Antenna);
                string Initial_ASM = ASMport[0];

                for (int k = 0; k < ASM_ports.Count; k++)
                {
                    if(k != ASM_ports.Count-1)
                    {
                        ASMport[k] = ASMport[k + 1];
                    }
                    else
                    {
                        ASMport[k] = Initial_ASM;
                    }
                        
                }
            }

            foreach (var item in Combination)
            {
                Spara_Trigger_Group test_group = new Spara_Trigger_Group();

                if (item[0] != null) test_group.ASM1 = convert_ASM_name(item[0]);
                if (item[1] != null) test_group.ASM2 = convert_ASM_name(item[1]);
                if (item[2] != null) test_group.ASM3 = convert_ASM_name(item[2]);

                foreach (TestCon each_Test in TC_list)
                {
                    if (each_Test.Status_file == "T.B.D") continue;

                    if (IsMatched(item, each_Test.Input_Port, each_Test.Output_Port))
                    {
                        if(test_group.TestCon_List.Count==0)
                        {
                            test_group.Status_File = each_Test.Status_file;
                            test_group.Group_TYP = Globals.Spara_config_INFO.Get_SubGroup(Band, each_Test.Parameter);
                            test_group.Tempearature = each_Test.Temperature;
                            test_group.Test_Input = each_Test.Input_Port;
                            test_group.Test_Output = each_Test.Output_Port;
                        }

                        test_group.TestCon_List.Add(each_Test);
                        
                    }
                    else if(IsMatched(item, each_Test.Output_Port, each_Test.Input_Port))
                    {
                        if (test_group.TestCon_List.Count == 0)
                        {
                            test_group.Status_File = each_Test.Status_file;
                            test_group.Group_TYP = Globals.Spara_config_INFO.Get_SubGroup(Band, each_Test.Parameter);
                            test_group.Tempearature = each_Test.Temperature;
                            test_group.Test_Input = each_Test.Input_Port;
                            test_group.Test_Output = each_Test.Output_Port;
                        }

                        test_group.TestCon_List.Add(each_Test);
                    }
                }
                combinations.Add(test_group);
            }

            return combinations;
        }

        private bool IsMatched(string[] ASM_table, string input_P, string output_P)
        {
            bool IsMatched = false;
            bool input_bool = false;
            bool output_bool = false;

            foreach (var item in ASM_table)
            {
                if (input_P == item) input_bool = true;
            }
            foreach (var item in ASM_table)
            {
                if (output_P == item) output_bool = true;
            }

            if(input_bool && output_bool) IsMatched = true;
            return IsMatched;
        }


        public void Set_ASMPort(string input_port, string output_port, ref Spara_Trigger_Group Spara_Trigger)
        {
            int Ant_index = 0;
            int ASM_index = 0;

            List<string> ANT_port = new List<string>();
            List<string> ASM_port = new List<string>();

            foreach (var item in Globals.Spara_config_INFO.Dic_PortDefinition)
            {
                if (item.Value.Contains("ANT_OUT"))
                {
                    ANT_port.Add(item.Key);
                    continue;
                }
                if (item.Value.Contains("ASM"))
                {
                    ASM_port.Add(item.Key);
                    continue;
                }
            }

            if (ANT_port.Count == 0 || ASM_port.Count == 0) return;

            if (Globals.Spara_config_INFO.Dic_PortDefinition[input_port] == "ASM" &&
                Globals.Spara_config_INFO.Dic_PortDefinition[output_port] == "ANT_OUT")
            {
                foreach (var each_ANT in ANT_port)
                {
                    if (output_port == each_ANT) break;
                    Ant_index++;
                }

                switch (Ant_index)
                {
                    case 0:
                        Spara_Trigger.ASM1 = convert_ASM_name(input_port);
                        break;
                    case 1:
                        Spara_Trigger.ASM2 = convert_ASM_name(input_port);
                        break;
                    case 2:
                        Spara_Trigger.ASM3 = convert_ASM_name(input_port);
                        break;

                    default:
                        break;
                }
            }
            else if (Globals.Spara_config_INFO.Dic_PortDefinition[input_port] == "ASM" &&
                     Globals.Spara_config_INFO.Dic_PortDefinition[output_port] == "ASM")
            {
                if (Spara_Trigger.ASM1 == "TERM" && Spara_Trigger.ASM2 == "TERM")
                {
                    Spara_Trigger.ASM1 = convert_ASM_name(input_port);
                    Spara_Trigger.ASM2 = convert_ASM_name(output_port);
                }
                else
                {
                    Spara_Trigger.ASM3 = convert_ASM_name(output_port);
                }

            }
            else if (Globals.Spara_config_INFO.Dic_PortDefinition[input_port] == "ASM")
            {
                Spara_Trigger.ASM1 = convert_ASM_name(input_port);
            }
            else if (Globals.Spara_config_INFO.Dic_PortDefinition[output_port] == "ASM")
            {
                Spara_Trigger.ASM1 = convert_ASM_name(output_port);
            }
        }

        private string convert_ASM_name(string port_name)
        {
            string converted_ASMportname = "";
            
            if(port_name.ToUpper().Contains("2G"))
            {
                converted_ASMportname = "GSM";
            }
            else if(port_name.ToUpper().Contains("MIMO"))
            {
                converted_ASMportname = "MIMO";
            }
            else if (port_name.ToUpper().Contains("DRX"))
            {
                converted_ASMportname = "DRX";
            }
            else if (port_name.ToUpper().Contains("LMB"))
            {
                converted_ASMportname = "LMB";
            }
            else 
            {
                converted_ASMportname = port_name;
            }

            return converted_ASMportname;
        }

        private bool search_Trigger_group(string Band, string Test_ID, TestCon Test_condition, ref List<Spara_Trigger_Group> Spara_trigger_conditions)
        {
            //this function is not allowed test condition duplication for trigger group

            bool find_Trigger_Group_to_add = false;

            if (Test_condition.Status_file != "T.B.D")
            {
                string Group = Globals.Spara_config_INFO.Get_Group(Band);
                string SubGroup = Globals.Spara_config_INFO.Get_SubGroup(Band, Test_ID);
                string Input_Type = Globals.Spara_config_INFO.Input_Match(Band, Test_ID);
                string Output_Type = Globals.Spara_config_INFO.Output_Match(Band, Test_ID);

                if (Spara_trigger_conditions.Count == 0) return false;

                foreach (var item in Spara_trigger_conditions)
                {
                    if (item.Group_TYP == SubGroup && item.Status_File == Test_condition.Status_file)
                    {
                        if (item.Port_matching(Test_condition.Input_Port, Test_condition.Output_Port, Input_Type, Output_Type))
                        {
#if false
                            foreach (var Trigger_Testcon in item.TestCon_List)
                            {
                                if (Test_ID == "Input_VSWR" && Trigger_Testcon.Parameter == "Input_VSWR") ;
                                {
                                    if (Trigger_Testcon.PA_MODE != Test_condition.PA_MODE) return false;
                                }
                            }
#endif
                            find_Trigger_Group_to_add = true;
                            item.TestCon_List.Add(Test_condition);
                            Insert_ASM(Test_condition.Input_Port, item);
                            Insert_ASM(Test_condition.Output_Port, item);
                            break;
                        }

                        if (Test_ID.ToUpper().Contains("_RL") || Test_ID == "ISO:ANT, ANT")
                        {
                            if (item.Port_matching(Test_condition.Output_Port, Test_condition.Input_Port, Input_Type, Output_Type))
                            {
                                find_Trigger_Group_to_add = true;
                                item.TestCon_List.Add(Test_condition);
                                Insert_ASM(Test_condition.Input_Port, item);
                                Insert_ASM(Test_condition.Output_Port, item);
                                break;
                            }
                        }

                    }
                }
            }

            return find_Trigger_Group_to_add;
        }

        private void Insert_ASM(string port_Target, Spara_Trigger_Group trigger_Group)
        {
            bool IsExist = false;
            string current_ASMport = convert_ASM_name(port_Target);

            if (IsEmpty(port_Target)) return;

            if (Globals.Spara_config_INFO.Dic_PortDefinition[port_Target] == "ASM")
            {
                List<string> exist_port = new List<string>();
                exist_port.Add(trigger_Group.ASM1);
                exist_port.Add(trigger_Group.ASM2);
                exist_port.Add(trigger_Group.ASM3);

                foreach (var item in exist_port)
                {
                    if(item == current_ASMport)
                    {
                        IsExist = true;
                    }
                }

                if (!IsExist && trigger_Group.ASM1 == "TERM")
                {
                    trigger_Group.ASM1 = current_ASMport;
                }
                else if (!IsExist && trigger_Group.ASM2 == "TERM")
                {
                    trigger_Group.ASM2 = current_ASMport;
                }
                else if (!IsExist && trigger_Group.ASM3 == "TERM")
                {
                    trigger_Group.ASM3 = current_ASMport;
                }
            }

           
        }

    }
}
