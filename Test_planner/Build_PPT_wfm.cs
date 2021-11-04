using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics; //Process to kill();
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Base;
using Test_Planner;

namespace S_para_planner
{
    public partial class Build_PPT_wfm : Form
    {
        public Dictionary<string, string> SnP_Directories = new Dictionary<string, string>();
        public int Excel_Proc_ID = 0;

        public Dictionary<string, List<SnpData>> Dic = new Dictionary<string, List<SnpData>>();
        public List<string> Band_list = new List<string>();

        public Build_PPT_wfm()
        {
            InitializeComponent();
            ProgressBar_Init(200);
            this.Show_Path_TCF.Text = "Please select S-para test plan";
            this.Show_Path_Unit.Text = "Please select SNP data folder tested with same plan";
            this.BTN_Build_PPT_Plan.Enabled = false;
            this.BTN_LoadUnits.Enabled = false;
        }
        public void ProgressBar_Init(int count)
        {
            this.progressBar1.Style = ProgressBarStyle.Continuous;
            this.progressBar1.Minimum = 0;
            this.progressBar1.Maximum = count;
            this.progressBar1.Step = (int)((progressBar1.Maximum - progressBar1.Minimum) / count);
            this.progressBar1.Value = 0;
            this.progressBar1.MarqueeAnimationSpeed = 1;
        }
        public void ProgBar_execute_step()
        {
            this.progressBar1.PerformStep();
        }

        private void BTN_LoadTCF_Click(object sender, EventArgs e)
        {
            string PFN_TestPlan = "";
            string PathDefault = "C:\\ProgramData\\FlexTest\\GENTLE_BREED";
            BTN_LoadTCF.BackColor = Color.LightGray;
            Show_Path_TCF.BackColor = Color.LightGray;
            this.BTN_Build_PPT_Plan.Enabled = false;
            this.BTN_LoadUnits.Enabled = false;

            try
            {
                System.Windows.Forms.OpenFileDialog OpenDialogEntity = new System.Windows.Forms.OpenFileDialog();


                OpenDialogEntity.InitialDirectory = (Directory.Exists(PathDefault) ? PathDefault : System.Environment.SpecialFolder.MyComputer.ToString());
                OpenDialogEntity.Filter = "Excel Files (.xlsx;.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                //OpenDialogEntity.Filter = "Text Files (.txt)|*.txt|Text Files (.stp)|*.stp|Text Files (.csv)|*.csv|Excel Files (.xlsx;.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                OpenDialogEntity.FilterIndex = 1;
                OpenDialogEntity.Multiselect = false;
                OpenDialogEntity.CheckFileExists = false;

                if (OpenDialogEntity.ShowDialog() == DialogResult.OK)
                {
                    PFN_TestPlan = OpenDialogEntity.FileName;
                }
                else
                {
                    return;
                }
            }
            catch
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open file from dialog");
                ErrMsg.AppendFormat("\nPlease check open file format (S-para TCF)");
                MessageBox.Show("Error on file loading", ErrMsg.ToString());
                //Environment.Exit(0);
            }

            try
            {
                Excel_File PPT_source = new Excel_File(false, PFN_TestPlan);
                this.Excel_Proc_ID = PPT_source.ProcID;
                PPT_source.Show(true);
                PPT_source.App.ScreenUpdating = true;

                int Start_row = 1;
                int Stop_row = 100000;
                int Start_col = 2; //Header list[2] = Enable "x"
                int Stop_col = 200;

                List<string> Header_list = PPT_source.Find_Header("Condition_FBAR", "Enable", "x", ref Start_row, ref Stop_row, ref Start_col, ref Stop_col);
                string[,] Full_data = PPT_source.ReadData_From_WorkSheet("Condition_FBAR", Start_row + 1, Stop_row, 1, Stop_col);

                int Index_Test_Enable = 1;
                int Index_Test_Mode = 3;
                
                for (int i = 0; i < Header_list.Count; i++)
                {
                    if (Header_list[i].Trim().ToUpper().Contains("ENABLE")) Index_Test_Enable = i;
                    if (Header_list[i].Trim().ToUpper().Contains("TEST MODE")) Index_Test_Mode = i;

                }

                ProgressBar_Init(Full_data.GetLength(0));

                SNP_structure public_data = new SNP_structure(Full_data, Header_list, Index_Test_Enable, Index_Test_Mode, this.progressBar1);

                int count_tick = this.progressBar1.Value;

                this.Dic = public_data.Dic;
                this.Band_list = public_data.Band_list;

                foreach (var item in this.Band_list)
                {
                    this.CBox_Bands.Items.Add(item, true);
                }

                Kill_Process(this.Excel_Proc_ID); //Dispose Excel file after Memory loading 
                BTN_LoadTCF.BackColor = Color.YellowGreen;
                Show_Path_TCF.BackColor = Color.GreenYellow;
                this.BTN_LoadUnits.Enabled = true;
            }
            catch
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open file from dialog");
                ErrMsg.AppendFormat("\nPlease check opend file or file name or location");
                MessageBox.Show("Error on file loading in initialization", ErrMsg.ToString());
                Kill_Process(this.Excel_Proc_ID);
            }
        }
        private void Kill_Process(int P_ID)
        {
            int ProcID = P_ID;
            Process Proc = Process.GetProcessById(ProcID);
            Proc.Kill();
        }

        private void BTN_LoadUnits_Click(object sender, EventArgs e)
        {
            string SNPdata_loot = "";
            this.CBox_UnitPath.Items.Clear();
            Show_Path_Unit.BackColor = Color.LightGray;
            BTN_LoadUnits.BackColor = Color.LightGray;

            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "  데이터가 위치한 상위 폴더에 맞추고 OK";
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                string folderPath = Path.GetDirectoryName(folderBrowser.FileName);
                SNPdata_loot = folderPath;
                Show_Path_Unit.Text = folderPath;
            }
            else
            {
                return;
            }

            int count_find = 0;

            string[] Gross_directories = Directory.GetDirectories(SNPdata_loot);
            foreach (string each_Path in Gross_directories)
            {
                bool IsSNP_Directory = false;
                string[] Files = Directory.GetFiles(each_Path);
                foreach (string each_file in Files)
                {
                    if (each_file.Contains(".s2p") ||
                       each_file.Contains(".s3p") ||
                       each_file.Contains(".s4p") ||
                       each_file.Contains(".s5p") ||
                       each_file.Contains(".s6p") ||
                       each_file.Contains(".s7p") ||
                       each_file.Contains(".s8p") ||
                       each_file.Contains(".s9p") ||
                       (each_file.Contains(".s") && each_file.Contains("p")))
                    {
                        IsSNP_Directory = true;
                        count_find++;
                        break;
                    }
                }

                if (IsSNP_Directory)
                {
                    Show_Path_Unit.BackColor = Color.GreenYellow;
                    BTN_LoadUnits.BackColor = Color.YellowGreen;

                    string Directory_name = new DirectoryInfo(each_Path).Name;
                    this.CBox_UnitPath.Items.Add(Directory_name, true);
                    this.SnP_Directories.Add(Directory_name, each_Path);
                }
            }

            if (count_find != 0) this.BTN_Build_PPT_Plan.Enabled = true;
        }

        private class SNP_structure
        {
            private int Total_row_count;
            private int Total_col_count;
            private int Data_key_enble;
            private int Header_key_enble;
            public List<string> Header = new List<string>();

            public Dictionary<string, List<SnpData>> Dic = new Dictionary<string, List<SnpData>>();
            public List<string> Band_list = new List<string>();

            public SNP_structure(string[,] Excel_2D_array, List<string> Header_list, int Index_enable, int Index_Test_Mode, ProgressBar Pbar)
            {
                clear();
                Total_row_count = Excel_2D_array.GetLength(0);
                Total_col_count = Excel_2D_array.Length / Total_row_count;
                
                Header = Header_list;

                int index_PortDefine = 0;
                int index_TestNum = 0;
                int index_SNP_Status = 0;

                for (int i = 0; i < Header_list.Count; i++)
                {
                    if (Header_list[i].Trim().ToUpper().Contains("PORT_DEFINE"))
                    {
                        index_PortDefine = i;
                    }

                    if (Header_list[i].Trim().ToUpper().Contains("TEST_NUM"))
                    {
                        index_TestNum = i;
                    }

                    if (Header_list[i].Trim().ToUpper().Contains("PARAMETER HEADER"))
                    {
                        index_SNP_Status = i;
                    }

                }

                string port_define_str = "";
                string SNP_FileNum = "";
                string Status_FileName = "";
                SnpCon SPcon = new SnpCon();

                List<SnpData> SNPdata_ListFull = new List<SnpData>();

                for (int i = 0; i < Total_row_count; i++)
                {
                    if (Excel_2D_array[i, Index_enable].Trim().ToUpper() != "X") continue;
                    if (Convert.ToInt32(Excel_2D_array[i, index_TestNum].Trim()) >= 90000) continue;

                    List<string> Read_row = new List<string>();
                    for (int j = 0; j < Total_col_count; j++)
                    {
                        Read_row.Add(Excel_2D_array[i, j]);
                    }

                    if (Excel_2D_array[i, Index_Test_Mode].Trim().ToUpper() == "DC")
                    {
                        port_define_str = Excel_2D_array[i, index_PortDefine];
                        SNP_FileNum = Excel_2D_array[i, index_TestNum];
                        Status_FileName = Excel_2D_array[i, index_SNP_Status]; ;
                        SPcon.GetCondition(Read_row, Header_list, "TXREG0B", "TXREG0C", "ASM_ANT1", "ASM_ANT2", "ASM_UAT");
                    }
                    else if(Excel_2D_array[i, Index_Test_Mode].Trim().ToUpper() == "FBAR")
                    {
                        try
                        {
                            SnpData SPdata = new SnpData();
                            SPcon.CopyCon_Data(ref SPdata);
                            SPdata.GetData(Read_row, Header_list, port_define_str, SNP_FileNum, Status_FileName);  //cond = 1, data row = n what will happen? 
                            SNPdata_ListFull.Add(SPdata);
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        
                    }

                    Pbar.PerformStep();
                }

                foreach (SnpData item in SNPdata_ListFull)
                {
                    string Band_key = item.KeySheet + "|" + item.SnpCon_TestID;
                    SearchKey_n_Insert(Band_key, item);

                    Band_list.Add(item.KeySheet);
                    Band_list = Band_list.Distinct().ToList();
                }

                //As Final result 
                //SNP_structure.Dic : full list of snp config with band + test id keys
                //SNP_structure.Band_list : for check box selection
            }
            private void clear()
            {
                this.Total_row_count = 0;
                this.Total_col_count = 0;
            }

            private void SearchKey_n_Insert(string Band_TestID_key, SnpData item)
            {
                if(this.Dic.ContainsKey(Band_TestID_key))
                {
                    this.Dic[Band_TestID_key].Add(item);
                }
                else
                {
                    List<SnpData> new_list = new List<SnpData>();
                    new_list.Add(item);
                    this.Dic.Add(Band_TestID_key, new_list);
                }
            }

        }

        public class SnpCon
        {
            public string PA_Bias_drv;
            public string PA_Bias_main;

            public string VCC_V;
            public string VBATT_V;
            public string VDDLNA_V;

            public string TRX_ON_mipi;
            public string TX_BAND_mipi;
            public string TX_INPUT_mipi;
            public string TX_OUTPUT_mipi;
            public string RX_BAND_mipi;
            public string RX_OUTPUT_mipi;
            public string RX_MODE_mipi;

            public string ASM_ANT1;
            public string ASM_ANT2;
            public string ASM_ANT3;

            public SnpCon()
            {
                clear();
            }
            public void GetCondition(List<string> Datarow, List<string> header, string PAbias_drv, string PAbias_main, string ASM_ANT1, string ASM_ANT2, string ASM_ANT3)
            {
                for (int i = 0; i < header.Count; i++)
                {
                    string Comp_str = header[i].Trim().ToUpper();
                    if (Comp_str.Contains("TRX_ON")) this.TRX_ON_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("TX_BAND")) this.TX_BAND_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("TX_INPUT")) this.TX_INPUT_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("TX_OUTPUT")) this.TX_OUTPUT_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("RX_BAND")) this.RX_BAND_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("RX_OUTPUT")) this.RX_OUTPUT_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains("LNA_MODE")) this.RX_MODE_mipi = Datarow[i].Trim();
                    if (Comp_str.Contains(PAbias_drv.Trim().ToUpper())) this.PA_Bias_drv = Datarow[i].Trim();
                    if (Comp_str.Contains(PAbias_main.Trim().ToUpper())) this.PA_Bias_main = Datarow[i].Trim();

                    if (Comp_str.Contains(ASM_ANT1.Trim().ToUpper())) this.ASM_ANT1 = Datarow[i].Trim();
                    if (Comp_str.Contains(ASM_ANT2.Trim().ToUpper())) this.ASM_ANT2 = Datarow[i].Trim();
                    if (Comp_str.Contains(ASM_ANT3.Trim().ToUpper())) this.ASM_ANT3 = Datarow[i].Trim();

                    if (Comp_str.Contains("VCC_V")) this.VCC_V = Datarow[i].Trim();
                    if (Comp_str.Contains("VBAT_V")) this.VBATT_V = Datarow[i].Trim();
                    if (Comp_str.Contains("LNAVDD_V")) this.VDDLNA_V = Datarow[i].Trim();
                }
            }

            public void CopyCon_Data(ref SnpData SnpData)
            {
                SnpData.PA_Bias_drv = this.PA_Bias_drv;
                SnpData.PA_Bias_main = this.PA_Bias_main;
                SnpData.VCC_V = this.VCC_V;
                SnpData.VBATT_V = this.VBATT_V;
                SnpData.VDDLNA_V = this.VDDLNA_V;

                SnpData.TRX_ON_mipi = this.TRX_ON_mipi;
                SnpData.TX_BAND_mipi = this.TX_BAND_mipi;
                SnpData.TX_INPUT_mipi = this.TX_INPUT_mipi;
                SnpData.TX_OUTPUT_mipi = this.TX_OUTPUT_mipi;
                SnpData.RX_BAND_mipi = this.RX_BAND_mipi;
                SnpData.RX_OUTPUT_mipi = this.RX_OUTPUT_mipi;
                SnpData.GainMode = this.RX_MODE_mipi;
                SnpData.RX_MODE_mipi = this.RX_MODE_mipi;

                SnpData.ASM_ANT1 = this.ASM_ANT1;
                SnpData.ASM_ANT2 = this.ASM_ANT2;
                SnpData.ASM_ANT3 = this.ASM_ANT3;
            }

            private void clear()
            {
                this.PA_Bias_drv = "";
                this.PA_Bias_main = "";

                this.VCC_V = "";
                this.VBATT_V = "";
                this.VDDLNA_V = "";

                this.TRX_ON_mipi = "";
                this.TX_BAND_mipi = "";
                this.TX_INPUT_mipi = "";
                this.TX_OUTPUT_mipi = "";
                this.RX_BAND_mipi = "";
                this.RX_OUTPUT_mipi = "";
                this.RX_MODE_mipi = "";

                this.ASM_ANT1 = "";
                this.ASM_ANT2 = "";
                this.ASM_ANT3 = "";
            }
        }

        public class SnpData
        {
            public string SNP_File_Name;
            public string port_define;
            public int port_define_cnt;
            public string StatusFile;

            public string KeySheet;
            public string SNPmode;

            public string Param_spec;
            public string PA_BAND;
            public string RX_OUT;
            public string GainMode;
            public string SnpCon_TestID;  //TestID
            public string SnpCon_TestName;

            public string Input_Port;
            public string Output_Port;
            public string S_param;
            public string Start_Freq;
            public string Stop_Freq;

            public string Test_limit_L;
            public string Test_limit_M;
            public string Test_limit_H;

            public string PA_Bias_drv;
            public string PA_Bias_main;

            public string VCC_V;
            public string VBATT_V;
            public string VDDLNA_V;

            public string TestNum;
            public string TRX_ON_mipi;
            public string TX_BAND_mipi;
            public string TX_INPUT_mipi;
            public string TX_OUTPUT_mipi;
            public string RX_BAND_mipi;
            public string RX_OUTPUT_mipi;
            public string RX_MODE_mipi;

            public string Temperature;

            public string ASM_ANT1;
            public string ASM_ANT2;
            public string ASM_ANT3;

            public SnpData()
            {
                clear();
            }
            private void GetFile(string Port_def)
            {
                string[] Port_desc = Port_def.Trim().Split(',');
                this.port_define = Port_def.Trim();
                this.port_define_cnt = Port_desc.Length;
                this.SNP_File_Name = GetFileName();
            }

            private string GetFileName()
            {
                StringBuilder SNP_file_name = new StringBuilder();
                SNP_file_name.AppendFormat("_{0}", (this.TestNum == "" ? "NA" : this.TestNum));
                SNP_file_name.AppendFormat("_{0}", (this.TRX_ON_mipi == "" ? "NA" : this.TRX_ON_mipi));
                
                SNP_file_name.AppendFormat("_{0}", (this.TX_BAND_mipi == "" ? "NA" : this.TX_BAND_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.TX_INPUT_mipi == "" ? "NA" : this.TX_INPUT_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.TX_OUTPUT_mipi == "" ? "NA" : this.TX_OUTPUT_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.RX_BAND_mipi == "" ? "NA" : this.RX_BAND_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.RX_OUTPUT_mipi == "" ? "NA" : this.RX_OUTPUT_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.RX_MODE_mipi == "" ? "NA" : this.RX_MODE_mipi));
                SNP_file_name.AppendFormat("_{0}", (this.Temperature == "" ? "NA" : this.Temperature));
                SNP_file_name.AppendFormat(".s{0}p", Convert.ToString(this.port_define_cnt));

                return SNP_file_name.ToString();
            }
           
            public void GetData(List<string> Datarow, List<string> header, string Port_def, string SNP_FileNum, string StatusFile_Name)
            {
                for (int i = 0; i < header.Count; i++)
                {
                    string Comp_str = header[i].Trim().ToUpper();

                    if (Comp_str.Contains("TEST_NUM")) this.TestNum = SNP_FileNum.Trim();
                    if (Comp_str.Contains("TEST MODE")) this.SNPmode = Datarow[i].Trim();
                    if (Comp_str.Contains("SPEC SHEET NAME")) this.KeySheet = Datarow[i].Trim();
                    if (Comp_str.Contains("PARA.SPEC")) this.Param_spec = Datarow[i].Trim();
                    if (Comp_str.Contains("PA_BAND")) this.PA_BAND = Datarow[i].Trim();
                    if (Comp_str.Contains("TUNABLE_BAND")) this.RX_OUT = Datarow[i].Trim();
                    //if (Comp_str.Contains("LNA_GAIN")) this.GainMode = Datarow[i].Trim();
                    if (Comp_str.Contains("PARAMETER HEADER"))
                    {
                        string[] tempArry = Datarow[i].Trim().Split(']');
                        this.SnpCon_TestID = tempArry[0].Replace("[", "").Trim();
                        this.SnpCon_TestName = tempArry[1].Trim();
                        this.StatusFile = StatusFile_Name;
                    }
                    if (Comp_str.Contains("INPUT PORT")) this.Input_Port = Datarow[i].Trim();
                    if (Comp_str.Contains("OUTPUT PORT")) this.Output_Port = Datarow[i].Trim();
                    if (Comp_str.Contains("PORT_DEFINE")) this.port_define = Datarow[i].Trim();
                    if (Comp_str.Contains("DM_S-PARAM")) this.S_param = Datarow[i].Trim();
                    if (Comp_str.Contains("START_FREQ")) this.Start_Freq = Datarow[i].Trim();
                    if (Comp_str.Contains("STOP_FREQ")) this.Stop_Freq = Datarow[i].Trim();

                    if (Comp_str.Contains("MIN_LIMIT")) this.Test_limit_L = Datarow[i].Trim();
                    if (Comp_str.Contains("TYP_LIMIT")) this.Test_limit_M = Datarow[i].Trim();
                    if (Comp_str.Contains("MAX_LIMIT")) this.Test_limit_H = Datarow[i].Trim();

                    if (Comp_str.Contains("SET_TEMP")) this.Temperature = Datarow[i].Trim();
                }

                GetFile(Port_def);
            }

            private void clear()
            {
                this.SNP_File_Name = "";
                this.port_define = "";
                this.port_define_cnt = 0;
                this.StatusFile = "";

                this.KeySheet = "";
                this.SNPmode = "";

                this.Param_spec = "";
                this.PA_BAND = "";
                this.RX_OUT = "";
                this.GainMode = "";
                this.SnpCon_TestID = "";

                this.Input_Port = "";
                this.Output_Port = "";
                this.S_param = "";
                this.Start_Freq = "";
                this.Stop_Freq = "";

                this.Test_limit_L = "";
                this.Test_limit_M = "";
                this.Test_limit_H = "";

                this.PA_Bias_drv = "";
                this.PA_Bias_main = "";

                this.VCC_V = "";
                this.VBATT_V = "";
                this.VDDLNA_V = "";

                this.TestNum = "";
                this.TRX_ON_mipi = "";
                this.TX_BAND_mipi = "";
                this.TX_INPUT_mipi = "";
                this.TX_OUTPUT_mipi = "";
                this.RX_BAND_mipi = "";
                this.RX_OUTPUT_mipi = "";
                this.RX_MODE_mipi = "";

                this.Temperature = "";

                this.ASM_ANT1 = "TERM";
                this.ASM_ANT2 = "TERM";
                this.ASM_ANT3 = "TERM";

            }
   
        }

        private void BTN_Build_PPT_Plan_Click(object sender, EventArgs e)
        {
            List<string> Target_band = Get_CheckedItem(this.CBox_Bands);
            List<string> Target_DUT = Get_CheckedItem(this.CBox_UnitPath);
            //this.SnP_Directories[Target_band[0]];

            this.BTN_LoadTCF.Enabled = false;
            this.BTN_LoadUnits.Enabled = false;
            this.BTN_Build_PPT_Plan.Enabled = false;

            Excel_File PPT_plan = new Excel_File(true);
            PPT_plan.Show(false);
            PPT_plan.App.ScreenUpdating = false;

            PPT_plan.Add_Sheet("Common");
            Build_CommonTable(PPT_plan, "Common");
            PPT_plan.Clear_Sheet();
            ProgressBar_Init(Target_band.Count * 15);

            foreach (string Band_key in Target_band)
            {
                Dictionary<string, List<SnpData>> Bands_Group = new Dictionary<string, List<SnpData>>();
                PPT_plan.Add_Sheet(Band_key);
                PPT_plan.Add_SlideHeader(Band_key);
                int SlideCnt = 1;
                int Current_row = 2;

                foreach (string Band_testID_key in this.Dic.Keys)
                {
                    if (Band_testID_key.Contains(Band_key))
                    {
                        List<SnpData> TestCon = this.Dic[Band_testID_key];
                        string[] Split_band = Band_testID_key.Split('|');
                        Bands_Group.Add(Split_band[1], TestCon);
                    }
                }

                Template_Slide InsertionLoss = Build_template(Band_key, "IL", "-2", "2", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide InputRL = Build_template(Band_key, "Input_RL", "0", "-10", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide OutputRL = Build_template(Band_key, "Output_RL", "0", "-10", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide TXGain = Build_template(Band_key, "TX_Gain_MAX", "25", "40", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide RXGain_G0 = Build_template(Band_key, "RX_Gain_G0", "13", "23", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide RXGain_CA_G0 = Build_template(Band_key, "RX_Gain_CA_G0", "13", "23", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide TXOOB = Build_template(Band_key, "TX_OOB_Gain", "-90", "40", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide RXOOB = Build_template(Band_key, "RX_OOB_Gain", "-90", "25", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_TXRX = Build_template(Band_key, "ISO:TX, RX", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_TRX_InActRX = Build_template(Band_key, "ISO:RX, InAct_RX", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_TX_InActRX = Build_template(Band_key, "ISO:TX, InAct_RX", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_InActRX_InActRX = Build_template(Band_key, "ISO:InAct_RX, InAct_RX", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_ReverseRX = Build_template(Band_key, "REV_ISO:RX, ANT", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);

                Template_Slide ISO_ANT_InActANT = Build_template(Band_key, "ISO:ANT, InAct_ANT", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_ANT_ANT = Build_template(Band_key, "ISO:ANT, ANT", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_TX_ASM = Build_template(Band_key, "ISO:TX, ASM", "-70", "30", Bands_Group, Target_DUT, this.SnP_Directories);

                Template_Slide ISO_ASM_To_ASM = Build_template(Band_key, "ISO:ASM, ASM", "-80", "20", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_ASM_To_InActANT = Build_template(Band_key, "ISO:ASM, InAct_ANT", "-80", "20", Bands_Group, Target_DUT, this.SnP_Directories);
                Template_Slide ISO_ASM_To_InActRX = Build_template(Band_key, "ISO:ASM, InAct_RX", "-80", "20", Bands_Group, Target_DUT, this.SnP_Directories);

                List<Template_Slide> SnP_SlideBuilder = new List<Template_Slide>();
                Revise_Slide(InsertionLoss, ref SnP_SlideBuilder);
                if (Band_key.Contains("_Rx"))
                {
                    Revise_Slide_RX(InputRL, ref SnP_SlideBuilder, "G0", 2); //0 = no effect, 1 = input, 2 = output, 3 = Gainmode (from return loss comparekey)
                    Revise_Slide_RX(OutputRL, ref SnP_SlideBuilder, "G0", 1); //0 = no effect, 1 = input, 2 = output, 3 = Gainmode (from return loss comparekey)
                }
                else
                {
                    Revise_Slide_CombinePair(InputRL, OutputRL, ref SnP_SlideBuilder);
                }

                Revise_Slide_Combine(TXGain, InputRL, ref SnP_SlideBuilder);
                Revise_Slide_RX(RXGain_G0, ref SnP_SlideBuilder, "G0", 2);
                Revise_Slide_RX(RXGain_CA_G0, ref SnP_SlideBuilder, "G0", 2);
                Revise_Slide(TXOOB, ref SnP_SlideBuilder);
                Revise_Slide_RX(RXOOB, ref SnP_SlideBuilder, "G0", 2);
                Split_Slide_RX(ISO_ReverseRX, ref SnP_SlideBuilder, "G0,G1,G2,G3,G4,G5", 1);

                Divide_Slide(ISO_TXRX, ref SnP_SlideBuilder, 5);
                Revise_Slide(ISO_TRX_InActRX, ref SnP_SlideBuilder);
                Revise_Slide(ISO_TX_InActRX, ref SnP_SlideBuilder);
                Revise_Slide(ISO_InActRX_InActRX, ref SnP_SlideBuilder);

                Revise_Slide(ISO_ANT_InActANT, ref SnP_SlideBuilder);
                if (Band_key.Contains("_Rx"))
                {
                    if (ISO_ANT_ANT.Pictures.Count != 0)
                    {
                        if (!ISO_ANT_ANT.Pictures[0].file_label[0].Contains("G0"))
                        {
                            Revise_Slide_RX(ISO_ANT_ANT, ref SnP_SlideBuilder, "G1", 2);
                        }
                        else
                        {
                            Revise_Slide_RX(ISO_ANT_ANT, ref SnP_SlideBuilder, "G0", 2);
                        }
                    }

                }
                else
                {
                    Revise_Slide(ISO_ANT_ANT, ref SnP_SlideBuilder);
                }

                Revise_Slide(ISO_TX_ASM, ref SnP_SlideBuilder);

                Revise_Slide(ISO_ASM_To_ASM, ref SnP_SlideBuilder);
                Revise_Slide(ISO_ASM_To_InActANT, ref SnP_SlideBuilder);
                Revise_Slide(ISO_ASM_To_InActRX, ref SnP_SlideBuilder);

                foreach (Template_Slide item in SnP_SlideBuilder)
                {
                    int SlideRow_Idx = Current_row;
                    if (item.Main_Title != "")
                    {
                        Write_Template(SlideCnt, ref Current_row, PPT_plan, Band_key, item);

                        this.progressBar1.PerformStep();

                        PPT_plan.Merge_Cell(Band_key, SlideRow_Idx, Current_row - 1, 1, 1);
                        SlideCnt++;
                    }
                }
            }

            string message = "Done : Build PPT plan";
            string caption = "All process done successfully";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            this.progressBar1.Value = this.progressBar1.Maximum;

            // Displays the MessageBox.
            result = MessageBox.Show(message, caption, buttons);
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                PPT_plan.Show(true);
                PPT_plan.App.ScreenUpdating = true;
            }

            this.BTN_LoadTCF.Enabled = true;
            this.BTN_LoadUnits.Enabled = true;
            this.BTN_Build_PPT_Plan.Enabled = true;
        }

        private void Revise_Slide(Template_Slide Slide_overPictures, ref List<Template_Slide> Mother_SlideList)
        {
            if (Slide_overPictures.Pictures.Count == 0) return;

            if (Slide_overPictures.Pictures.Count > 6)
            {
                int cnt_1 = Slide_overPictures.Pictures.Count / 6;
                float cnt_2 = Slide_overPictures.Pictures.Count % 6;
                if (cnt_2 != 0) cnt_1++;

                int Index_Pic = 0;
                for (int i = 0; i < cnt_1; i++)
                {
                    List<Template_PIC> new_PicList = new List<Template_PIC>();

                    for (int j = 0; j < 6; j++)
                    {
                        new_PicList.Add(Slide_overPictures.Pictures[Index_Pic]);
                        Index_Pic++;
                        if (Index_Pic == Slide_overPictures.Pictures.Count) break;
                    }

                    Template_Slide New_Slide = Slide_overPictures.Clone();
                    New_Slide.Pictures = new_PicList;
                    Mother_SlideList.Add(New_Slide);
                }
            }
            else if (Slide_overPictures.Pictures.Count == 4)
            {
                Slide_overPictures.slide_option = "Option4";
                Mother_SlideList.Add(Slide_overPictures);
            }
            else
            {
                Mother_SlideList.Add(Slide_overPictures);
            }
        }

        private void Revise_Slide_RX(Template_Slide Slide_overPictures, ref List<Template_Slide> Mother_SlideList, string GainMode, int SortIndex_Of_compareKey)
        {
            if (Slide_overPictures.Pictures.Count == 0) return;

            List<Template_PIC> Selected_GainOnly = new List<Template_PIC>();

            if(GainMode.Contains(','))
            {
                string[] GainModes = GainMode.Split(',');
                foreach (Template_PIC item in Slide_overPictures.Pictures)
                {
                    for (int i = 0; i < GainModes.Length; i++)
                    {
                        if (item.Compare_Key.Contains(GainModes[i].Trim())) Selected_GainOnly.Add(item);
                    }
                }
            }
            else
            {
                foreach (Template_PIC item in Slide_overPictures.Pictures)
                {
                    if (item.Compare_Key.Contains(GainMode)) Selected_GainOnly.Add(item);
                }
            }

            Template_Slide Revised_Slide = Slide_overPictures.Clone();
            Revised_Slide.Pictures = SortByKey(Selected_GainOnly, SortIndex_Of_compareKey);

            if (Revised_Slide.Pictures.Count > 6)
            {
                int cnt_1 = Revised_Slide.Pictures.Count / 6;
                float cnt_2 = Revised_Slide.Pictures.Count % 6;
                if (cnt_2 != 0) cnt_1++;

                int Index_Pic = 0;
                for (int i = 0; i < cnt_1; i++)
                {
                    List<Template_PIC> new_PicList = new List<Template_PIC>();

                    for (int j = 0; j < 6; j++)
                    {
                        new_PicList.Add(Revised_Slide.Pictures[Index_Pic]);
                        Index_Pic++;
                        if (Index_Pic == Revised_Slide.Pictures.Count) break;
                    }

                    Template_Slide New_Slide = Slide_overPictures.Clone();
                    New_Slide.Pictures = new_PicList;
                    Mother_SlideList.Add(New_Slide);
                }
            }
            else if (Revised_Slide.Pictures.Count == 4)
            {
                Revised_Slide.slide_option = "Option4";
                Mother_SlideList.Add(Revised_Slide);
            }
            else
            {
                Mother_SlideList.Add(Revised_Slide);
            }
        }

        private void Split_Slide_RX(Template_Slide Slide_overPictures, ref List<Template_Slide> Mother_SlideList, string Label, int SortIndex_Of_compareKey)
        {
            if (Slide_overPictures.Pictures.Count == 0) return;

            Template_Slide CopySlide_1 = new Template_Slide();
            Template_Slide CopySlide_2 = new Template_Slide();

            CopySlide_1.Main_Title = Slide_overPictures.Main_Title;
            CopySlide_1.Sub_Desc_Title = Slide_overPictures.Sub_Desc_Title;
            CopySlide_1.picture_cnt = Slide_overPictures.picture_cnt;
            CopySlide_1.slide_option = Slide_overPictures.slide_option;

            CopySlide_2.Main_Title = Slide_overPictures.Main_Title;
            CopySlide_2.Sub_Desc_Title = Slide_overPictures.Sub_Desc_Title;
            CopySlide_2.picture_cnt = Slide_overPictures.picture_cnt;
            CopySlide_2.slide_option = Slide_overPictures.slide_option;

            string[] Labels = Label.Split(',');
            foreach (Template_PIC item in Slide_overPictures.Pictures)
            {
                Template_PIC new_item1 = new Template_PIC();
                Template_PIC new_item2 = new Template_PIC();

                new_item1.Compare_Key = item.Compare_Key;
                new_item1.Title = item.Title;
                new_item1.Y_title = item.Y_title;
                new_item1.S_Parameter = item.S_Parameter;
                new_item1.markers = item.markers;
                new_item1.Freq_start = item.Freq_start;
                new_item1.Freq_stop = item.Freq_stop;
                new_item1.Mag_start = item.Mag_start;
                new_item1.Mag_stop = item.Mag_stop;

                new_item2.Compare_Key = item.Compare_Key;
                new_item2.Title = item.Title;
                new_item2.Y_title = item.Y_title;
                new_item2.S_Parameter = item.S_Parameter;
                new_item2.markers = item.markers;
                new_item2.Freq_start = item.Freq_start;
                new_item2.Freq_stop = item.Freq_stop;
                new_item2.Mag_start = item.Mag_start;
                new_item2.Mag_stop = item.Mag_stop;

                for (int i = 0; i < item.file_label.Count; i++)
                {
                    if (item.file_label[i].Contains("_-33_G") || item.file_label[i].Contains("_85_G")) continue;

                    bool IsMatched = false;
                    for (int j = 0; j < Labels.Length; j++)
                    {
                        if (item.file_label[i].Contains(Labels[j].Trim())) IsMatched = true;
                    }

                    if(IsMatched)
                    {
                        new_item1.SPEC_Start.Add(item.SPEC_Start[i]);
                        new_item1.SPEC_Stop.Add(item.SPEC_Stop[i]);
                        new_item1.SPEC_value.Add(item.SPEC_value[i]);
                        new_item1.file_label.Add(item.file_label[i]);
                        new_item1.file_name.Add(item.file_name[i]);
                        new_item1.file_path.Add(item.file_path[i]);
                    }
                    else
                    {
                        new_item2.SPEC_Start.Add(item.SPEC_Start[i]);
                        new_item2.SPEC_Stop.Add(item.SPEC_Stop[i]);
                        new_item2.SPEC_value.Add(item.SPEC_value[i]);
                        new_item2.file_label.Add(item.file_label[i]);
                        new_item2.file_name.Add(item.file_name[i]);
                        new_item2.file_path.Add(item.file_path[i]);
                    }
                }

                CopySlide_1.Pictures.Add(new_item1);
                CopySlide_2.Pictures.Add(new_item2);
            }

            Template_Slide Revised_CopySlide_1 = CopySlide_1.Clone();
            Revised_CopySlide_1.Pictures = SortByKey(CopySlide_1.Pictures, SortIndex_Of_compareKey);

            Template_Slide Revised_CopySlide_2 = CopySlide_2.Clone();
            Revised_CopySlide_2.Pictures = SortByKey(CopySlide_2.Pictures, SortIndex_Of_compareKey);

            Revise_Slide(Revised_CopySlide_1, ref Mother_SlideList);
            Revise_Slide(Revised_CopySlide_2, ref Mother_SlideList);
        }


        public List<Template_PIC> SortByKey(List<Template_PIC> Pictures, int index)
        {
            if (index == 0) return Pictures;

            SortedDictionary<string, List<Template_PIC>> Sorted_Dic = new SortedDictionary<string, List<Template_PIC>>();

            foreach (Template_PIC item in Pictures)
            {
                string Key = GetPort_FromKey(item.Compare_Key, index);
                
                if(!Sorted_Dic.ContainsKey(Key))
                {
                    List<Template_PIC> sub_list = new List<Template_PIC>();
                    sub_list.Add(item);
                    Sorted_Dic.Add(Key, sub_list);
                }
                else
                {
                    Sorted_Dic[Key].Add(item);
                }
            }

            List<Template_PIC> new_Pictures = new List<Template_PIC>();

            foreach (string SortedKey in Sorted_Dic.Keys)
            {
                foreach (Template_PIC item in Sorted_Dic[SortedKey])
                {
                    new_Pictures.Add(item);
                }
            }

            return new_Pictures;
        }

        private void Divide_Slide(Template_Slide Slide_overPictures, ref List<Template_Slide> Mother_SlideList, int index_split) //TX-RX ISO : 5 = RX band
        {
            if (Slide_overPictures.Pictures.Count == 0) return;

            List<Template_PIC> List_Pic = new List<Template_PIC>();
            List<string> List_Pic_Key = new List<string>();

            foreach (Template_PIC item in Slide_overPictures.Pictures)
            {
                string Split_ID = GetPort_FromKey(item.Compare_Key, index_split);
                List_Pic.Add(item);
                List_Pic_Key.Add(Split_ID);
            }

            do
            {
                List<Template_PIC> new_List = new List<Template_PIC>();
                
                SelectPic: 
                List<int> ToRemove_index = new List<int>();
                
                string Init_SplitID = List_Pic_Key[0];

                for (int i = 0; i < List_Pic.Count; i++)
                {
                    if (List_Pic_Key[i] == Init_SplitID)
                    {
                        new_List.Add(List_Pic[i]);
                        ToRemove_index.Add(i);
                    }
                }

                int remove_index = 0;

                foreach (int index in ToRemove_index)
                {
                    List_Pic.RemoveAt(index - remove_index);
                    List_Pic_Key.RemoveAt(index - remove_index);
                    remove_index++;
                }

                if (new_List.Count < 4 && List_Pic.Count != 0) goto SelectPic;

                Template_Slide New_Slide = Slide_overPictures.Clone();
                New_Slide.Pictures = new_List;
                Revise_Slide(New_Slide, ref Mother_SlideList);

            } while (List_Pic.Count != 0);
        }

        public bool CrossInOut(string CompareKey1, string CompareKey2, char Split_Char)
        {
            bool IsCombination = false;

            string[] Comp1 = CompareKey1.Trim().Split(Split_Char);
            string[] Comp2 = CompareKey2.Trim().Split(Split_Char);

            if (Comp1[1] == Comp2[1] && Comp1[2] == Comp2[2])
            {
                IsCombination = true;
            }
            else if(Comp1[1] == Comp2[2] && Comp1[2] == Comp2[1])
            {
                IsCombination = true;
            }

            return IsCombination;
        }

        public string GetPort_FromKey(string CompareKey1, int Index)
        {
            //Index = 0 : TestID
            //Index = 1 : Input Port (Spara input, Is Not Actual DUT input)
            //Index = 2 : Output Port (Spara output, Is Not Actual DUT output)

            string[] Comp1 = CompareKey1.Trim().Split('|');
            string GetPort = Comp1[Index];

            return GetPort;
        }


        private void Revise_Slide_Combine(Template_Slide Slide_overPictures, Template_Slide Slide_overPictures2, ref List<Template_Slide> Mother_SlideList)
        {
            if (Slide_overPictures.Pictures.Count == 0) return;

            List<Template_PIC> Merged_PicList = new List<Template_PIC>();
            Template_Slide New_Slide = Slide_overPictures.Clone();

            foreach (Template_PIC Pic_1 in Slide_overPictures.Pictures)
            {
                Merged_PicList.Add(Pic_1);
            }

            if (Slide_overPictures2.Pictures.Count != 0)
            {
                foreach (Template_PIC Pic_2 in Slide_overPictures2.Pictures)
                {
                    Merged_PicList.Add(Pic_2);
                }

                New_Slide.Main_Title = Slide_overPictures.Main_Title + " / " + Slide_overPictures2.Main_Title;
            }
            
            New_Slide.Pictures = Merged_PicList;
            Revise_Slide(New_Slide, ref Mother_SlideList);
        }

        private void Revise_Slide_CombinePair(Template_Slide Slide_overPictures, Template_Slide Slide_overPictures2, ref List<Template_Slide> Mother_SlideList)
        {
            if (Slide_overPictures.Pictures.Count == 0 && Slide_overPictures2.Pictures.Count == 0) return;
            if (Slide_overPictures.Pictures.Count == 0)
            {
                Revise_Slide(Slide_overPictures2, ref Mother_SlideList);
                return;
            }
            else if (Slide_overPictures2.Pictures.Count == 0)
            {
                Revise_Slide(Slide_overPictures, ref Mother_SlideList);
                return;
            }
            else if (Slide_overPictures.Pictures.Count != Slide_overPictures2.Pictures.Count)
            {
                Revise_Slide(Slide_overPictures, ref Mother_SlideList);
                Revise_Slide(Slide_overPictures2, ref Mother_SlideList);
                return;
            }


            List<Template_PIC> Merged_PicList = new List<Template_PIC>();

            Dictionary<string, Template_PIC> Merged_PicDic = new Dictionary<string, Template_PIC>();

            foreach (Template_PIC Pic_1 in Slide_overPictures.Pictures)
            {
                Merged_PicDic.Add(Pic_1.Compare_Key, Pic_1);
            }

            foreach (Template_PIC Pic_2 in Slide_overPictures2.Pictures)
            {
                Merged_PicDic.Add(Pic_2.Compare_Key, Pic_2);
            }

            do
            {
                List<Template_PIC> Sorted_PicList = new List<Template_PIC>();
                List<string> Picked_key = new List<string>();

                string init_key = Merged_PicDic.ElementAt(0).Key;
                string Port1 = GetPort_FromKey(init_key, 1);
                string Port2 = GetPort_FromKey(init_key, 2);
                string target_port = (Port1.Contains("ANT") || Port2.Contains("UAT") ? Port1 : Port2);

                foreach (string key in Merged_PicDic.Keys)
                {
                    if (GetPort_FromKey(key, 1).Contains(target_port) || GetPort_FromKey(key, 2).Contains(target_port))
                    {
                        Sorted_PicList.Add(Merged_PicDic[key]);
                        Picked_key.Add(key);
                    }
                }

                //return sorted_piclist to new slide
                Template_Slide New_Slide = Slide_overPictures.Clone();
                New_Slide.Main_Title = Slide_overPictures.Main_Title + " / " + Slide_overPictures2.Main_Title;
                New_Slide.Pictures = Sorted_PicList;
                Revise_Slide(New_Slide, ref Mother_SlideList);

                foreach (string key_remove in Picked_key)
                {
                    Merged_PicDic.Remove(key_remove);
                }

            } while (Merged_PicDic.Count != 0);
        }

        public void Write_Template(int SlideCnt, ref int Current_row, Excel_File excel_File, string sheet, Template_Slide Slide_snp)
        {
            int WIndex_row = 1;

            excel_File.Select_Sheet(sheet);
            excel_File.Write_SlideTitle(sheet, SlideCnt, ref Current_row, Slide_snp.Main_Title, Slide_snp.slide_option, false);
            excel_File.Write_SlideTitle(sheet, SlideCnt, ref Current_row, Slide_snp.Sub_Desc_Title, Slide_snp.slide_option, true);

            for (int i = 0; i < Slide_snp.Pictures.Count; i++)
            {
                Write_SlideData_TypeA(excel_File, sheet, ref Current_row, Slide_snp.Pictures[i], i);
            }

        }

        private void Write_SlideData_TypeA(Excel_File excel_File, string nSheet, ref int Current_row, Template_PIC Slide_picture, int PIC_index)
        {
            List<string> PicHeader = new List<string>();
            PicHeader.Add("Picture " + Convert.ToString(PIC_index + 1));
            PicHeader.Add("Files");
            PicHeader.Add("Path");
            Add_dummyToList(ref PicHeader, 2);
            PicHeader.Add("Name");
            Add_dummyToList(ref PicHeader, 2);
            PicHeader.Add("Label");
            Add_dummyToList(ref PicHeader, 1);

            excel_File.Select_Sheet(nSheet);
            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 2, PicHeader, 1);

            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 4, 6);
            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 7, 9);
            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 10, 11);

            Current_row++;

            int mergeRow_Picture_Idx = Current_row;
            int mergeRow_Idx = Current_row;

            for (int i = 0; i < Slide_picture.file_path.Count; i++)
            {
                List<string> Path_Name_Label = new List<string>();
                Path_Name_Label.Add(Slide_picture.file_path[i]);
                Add_dummyToList(ref Path_Name_Label, 2);
                Path_Name_Label.Add(Slide_picture.file_name[i]);
                Add_dummyToList(ref Path_Name_Label, 2);
                Path_Name_Label.Add(Slide_picture.file_label[i]);
                Add_dummyToList(ref Path_Name_Label, 1);

                excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 4, Path_Name_Label, 0);

                excel_File.Merge_Cell(nSheet, Current_row, Current_row, 4, 6);
                excel_File.Merge_Cell(nSheet, Current_row, Current_row, 7, 9);
                excel_File.Merge_Cell(nSheet, Current_row, Current_row, 10, 11);

                Current_row++;
            }

            excel_File.Merge_Cell(nSheet, mergeRow_Idx - 1, Current_row - 1, 3, 3);

            List<string> PicTitle = new List<string>();
            PicTitle.Add("Title");
            PicTitle.Add(Slide_picture.Title);

            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 3, PicTitle, 2);
            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 4, 11);
            Current_row++;

            List<string> Graph_setting = new List<string>();
            Graph_setting.Add("Graph Type");
            Graph_setting.Add("S-Parameter");
            Graph_setting.Add("Y-Axis Label");
            Graph_setting.Add("Freq Start");
            Graph_setting.Add("Freq Stop");
            Graph_setting.Add("Freq Grid");
            Graph_setting.Add("Mag Start");
            Graph_setting.Add("MagStop");
            Graph_setting.Add("Mag Grid");

            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 3, Graph_setting, 3);
            Current_row++;

            Graph_setting.Clear();
            if(Slide_picture.Y_title.ToUpper().Contains("RETURN")&& Slide_picture.Y_title.ToUpper().Contains("LOSS"))
            {
                Graph_setting.Add("Smith Chart");
            }
            else
            {
                Graph_setting.Add("Magnitude");
            }
            Graph_setting.Add(Slide_picture.S_Parameter);
            Graph_setting.Add(Slide_picture.Y_title);
            Graph_setting.Add(Slide_picture.Freq_start);
            Graph_setting.Add(Slide_picture.Freq_stop);
            Graph_setting.Add("Auto");
            Graph_setting.Add(Slide_picture.Mag_start);
            Graph_setting.Add(Slide_picture.Mag_stop);
            Graph_setting.Add("Auto");

            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 3, Graph_setting, 4);
            Current_row++;

            List<string> Add_Marker = new List<string>();
            Add_Marker.Add("Markers");
            Add_Marker.Add(as_1stringRow(Slide_picture.markers));
            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 3, Add_Marker, 2);
            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 4, 11);
            Current_row++;

            List<string> Specs = new List<string>();
            Specs.Add("Spec");
            Specs.Add("Freq Start");
            Specs.Add("Freq Stop");
            Specs.Add("Level");
            Specs.Add("Color");
            Specs.Add("Thickness");
            Specs.Add("Line Style");
            excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 3, Specs, 3);
            excel_File.Merge_Cell(nSheet, Current_row, Current_row, 10, 11);

            Current_row++;
            Specs.Clear();
            mergeRow_Idx = Current_row;

            List<string> get_Spec = remove_duplicate_SPec(Slide_picture);

            foreach (string Spec_values in get_Spec)
            {
                string[] spec_item = Spec_values.Split('|');
                Specs.Clear();
                Specs.Add(spec_item[0].Trim());
                Specs.Add(spec_item[1].Trim());
                Specs.Add(spec_item[2].Trim());
                Specs.Add("0x000000");
                Specs.Add("2 Pixel");
                Specs.Add("Solid");
                excel_File.WriteData_Row_with_formatting(nSheet, Current_row, 4, Specs, 4);
                excel_File.Merge_Cell(nSheet, Current_row, Current_row, 10, 11);
                Current_row++;
            }

            excel_File.Merge_Cell(nSheet, mergeRow_Idx - 1, Current_row - 1, 3, 3);
            excel_File.Merge_Cell(nSheet, mergeRow_Idx - 1, Current_row - 1, 10, 11);
            excel_File.Merge_Cell(nSheet, mergeRow_Picture_Idx - 1, Current_row - 1, 2, 2);           
        }

        private void Build_CommonTable(Excel_File PPT_plan, string SheetName)
        {
            List<string> TextHeader = new List<string>();

            StringBuilder ColorTable = new StringBuilder();
            ColorTable.Append("Files");
            ColorTable.Append(",Color");
            ColorTable.Append(",Thickness");
            ColorTable.Append(",Line Style");
            TextHeader.Add(ColorTable.ToString());

            for (int i = 0; i < 16; i++)
            {
                ColorTable.Clear();
                ColorTable.AppendFormat("File{0}", i + 1);
                ColorTable.Append(",Color");
                ColorTable.Append(",2 Pixel");
                ColorTable.Append(",Solid");
                TextHeader.Add(ColorTable.ToString());
            }

            int k = 1;
            foreach (string Table_Value in TextHeader)
            {
                string[] Temp = Table_Value.Split(',');
                List<string> Row_data = new List<string>();
                for (int i = 0; i < Temp.Length; i++)
                {
                    Row_data.Add(Temp[i].Trim());
                }

                if (k == 1)
                {
                    PPT_plan.WriteData_Row_with_formatting(SheetName, k, 1, Row_data, 5);
                    k++;
                }
                else
                {
                    PPT_plan.WriteData_Row_with_formatting(SheetName, k, 1, Row_data, 4);
                    k++;
                }
            }

            int d = 2;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.ForestGreen); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Blue); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Crimson); d++;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.DarkOliveGreen); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.DarkCyan); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Firebrick); d++;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Goldenrod); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.GreenYellow); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.HotPink); d++;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.LightGreen); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.LightBlue); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Tomato); d++;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.SeaGreen); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.RoyalBlue); d++;
            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.PaleVioletRed); d++;

            PPT_plan.Cformat_Color(SheetName, d, d, 2, 2, Color.Wheat); d++;

        }

        private List<string> remove_duplicate_SPec(Template_PIC Picture)
        {
            List<string> revised_List = new List<string>();
            
            for (int i = 0; i < Picture.SPEC_value.Count; i++)
            {
                StringBuilder combined_string = new StringBuilder();
                combined_string.Append(Picture.SPEC_Start[i]);
                combined_string.Append("|");
                combined_string.Append(Picture.SPEC_Stop[i]);
                combined_string.Append("|");
                combined_string.Append(Picture.SPEC_value[i]);
                revised_List.Add(combined_string.ToString());
            }

            revised_List = revised_List.Distinct().ToList();

            return revised_List;
        }

        private string as_1stringRow(List<string> list)
        {
            string result_line = "";

            foreach (var item in list)
            {
                if(result_line=="")
                {
                    result_line = item.Trim();
                }
                else
                {
                    result_line = result_line + ", " + item.Trim();
                }  
            }

            return result_line;
        }

        private void Add_dummyToList(ref List<string> List, int count)
        {
            for (int i = 0; i < count; i++)
            {
                List.Add("");
            }
        }

        private Template_Slide Build_template(string Band_key, string TestID, string Mag_Start_V, string Mag_Stop_V, Dictionary<string, List<SnpData>> Bands_Group, List<string> Target_DUT, Dictionary<string, string> SnP_Directories)
        {
            bool IsRL = (TestID.Contains("RL") ? true : false);

            Template_Slide new_slide = new Template_Slide();

            if (!Bands_Group.ContainsKey(TestID)) return new_slide;

            List<SnpData> snp_configs = Bands_Group[TestID];
            new_slide.Main_Title = (snp_configs[0].KeySheet.Replace('_', ' ') + (" : ") + snp_configs[0].SnpCon_TestName.Replace('_', ' ')).Replace("MAX", "");
            new_slide.Sub_Desc_Title = "Vbatt = " + snp_configs[0].VBATT_V + "V, Vcc1,2 = " + snp_configs[0].VCC_V
                                     + (snp_configs[0].PA_Bias_drv != "" ? "V, Vreg = 0x" + snp_configs[0].PA_Bias_drv + snp_configs[0].PA_Bias_main : "")
                                     + ", Gain Mode = " + (snp_configs[0].GainMode == "" ? "PA only" : snp_configs[0].GainMode)
                                     + "/ Measured @ Room & Over Temp";
            
            StringBuilder sub_title = new StringBuilder();
            SnpData RefSnp = snp_configs[0];
            if (TestID.Contains("ISO:ASM, ASM")) new_slide.Sub_Desc_Title = sub_title.AppendFormat("{0} (ANT1={1}, ANT2={2}, UAT={3})", new_slide.Sub_Desc_Title, RefSnp.ASM_ANT1, RefSnp.ASM_ANT2, RefSnp.ASM_ANT3).ToString();
            if (TestID.Contains("ISO:TX, RX")) new_slide.Main_Title = snp_configs[0].KeySheet.Replace('_', ' ') + (" : ") + "Active TX Input to Active Rx Isolation";
            if (TestID.Contains("REV_ISO:RX, ANT")) new_slide.Sub_Desc_Title = new_slide.Sub_Desc_Title.Replace(snp_configs[0].GainMode, "All Gain mode");
            if (TestID.Contains("ISO:TX, ASM")) new_slide.Main_Title = snp_configs[0].KeySheet.Replace('_', ' ') + (" : ") + "Active TX Input to ANT(DRX/MIMO) Isolation";
            if (TestID.Contains("RX_Gain_CA_G0")) new_slide.Main_Title = snp_configs[0].KeySheet.Replace('_', ' ') + (" : ") + "Gain at Each CA Case";

            //if (TestID.Contains("ISO:ASM, InAct_ANT")) new_slide.Sub_Desc_Title = sub_title.AppendFormat("{0} (ANT1={1}, ANT2={2}, UAT={3})", new_slide.Sub_Desc_Title, RefSnp.ASM_ANT1, RefSnp.ASM_ANT2, RefSnp.ASM_ANT3).ToString();
            //if (TestID.Contains("ISO:ASM, InAct_RX")) new_slide.Sub_Desc_Title = sub_title.AppendFormat("{0} (ANT1={1}, ANT2={2}, UAT={3})", new_slide.Sub_Desc_Title, RefSnp.ASM_ANT1, RefSnp.ASM_ANT2, RefSnp.ASM_ANT3).ToString();

            List<string> spec_ID = new List<string>();

            foreach (SnpData item in snp_configs)
            {
                string CompareKey_item = "";

                spec_ID.Add(item.Param_spec);

                if (TestID.Contains("IL")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;
                if (TestID.Contains("RX_Gain_G0")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;
                if (TestID.Contains("RX_Gain_CA_G0")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;
                if (TestID.Contains("TX_Gain_MAX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;
                if (TestID.Contains("TX_OOB_Gain")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.StatusFile;
                if (TestID.Contains("RX_OOB_Gain")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.StatusFile + "|" + item.GainMode;
                if (TestID.Contains("Gain_Ripple")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;
                if (TestID.Contains("Input_RL")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.GainMode;
                if (TestID.Contains("Output_RL")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.GainMode;
                if (TestID.Contains("ISO:TX, RX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.TX_BAND_mipi + "|" + item.RX_BAND_mipi; //4=TX band, 5=RX band
                if (TestID.Contains("ISO:RX, InAct_RX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.RX_OUTPUT_mipi;
                if (TestID.Contains("ISO:TX, InAct_RX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.TX_BAND_mipi + "|" + item.RX_BAND_mipi; //4=TX band, 5=RX band
                if (TestID.Contains("ISO:InAct_RX, InAct_RX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.TX_BAND_mipi + "|" + item.RX_BAND_mipi; //4=TX band, 5=RX band
                if (TestID.Contains("REV_ISO:RX, ANT")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.RX_OUTPUT_mipi;
                if (TestID.Contains("ISO:ANT, InAct_ANT")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param;

                if(item.KeySheet.ToUpper().Contains("_RX"))
                {
                    if (TestID.Contains("ISO:ANT, ANT")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3 + "|" + item.TX_BAND_mipi + "|" + item.GainMode;
                }
                else
                {
                    if (TestID.Contains("ISO:ANT, ANT")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3;
                }
                
                if (TestID.Contains("ISO:TX, ASM")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3 + "|" + item.Param_spec;
                if (TestID.Contains("ISO:ASM, ASM")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3;
                if (TestID.Contains("ISO:ASM, InAct_ANT")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3;
                if (TestID.Contains("ISO:ASM, InAct_RX")) CompareKey_item = item.SnpCon_TestID + "|" + item.Input_Port + "|" + item.Output_Port + "|" + item.S_param + "|" + item.RX_BAND_mipi + item.ASM_ANT1 + "|" + item.ASM_ANT2 + "|" + item.ASM_ANT3;

                if (new_slide.Pictures.Count == 0)
                {
                    Template_PIC new_pic = new Template_PIC();
                    //new_pic.Create_Picture(item, Target_DUT[0], SnP_Directories, CompareKey_item);
                    new_pic.Create_Picture_Multi(item, Target_DUT, SnP_Directories, CompareKey_item);
                    new_pic.Mag_start = Mag_Start_V;
                    new_pic.Mag_stop = Mag_Stop_V;
                    new_slide.Pictures.Add(new_pic);
                }
                else
                {
                    bool IsFind = false;

                    foreach (Template_PIC item_PIC in new_slide.Pictures)
                    {
                        if (item_PIC.Compare_Key == CompareKey_item)
                        {
                            //item_PIC.Add_ToPicture(item, Target_DUT[0], SnP_Directories);
                            item_PIC.Add_ToPicture_Multi(item, Target_DUT, SnP_Directories);
                            IsFind = true;
                            break;
                        }
                    }

                    if (!IsFind)
                    {
                        Template_PIC new_pic = new Template_PIC();
                        //new_pic.Create_Picture(item, Target_DUT[0], SnP_Directories, CompareKey_item);
                        new_pic.Create_Picture_Multi(item, Target_DUT, SnP_Directories, CompareKey_item);
                        new_pic.Mag_start = Mag_Start_V;
                        new_pic.Mag_stop = Mag_Stop_V;
                        new_slide.Pictures.Add(new_pic);
                    }
                }

            }

            spec_ID = spec_ID.Distinct().ToList();

            StringBuilder spec_IDList = new StringBuilder();

            spec_IDList.Append('(');
            
            int end_idx = 0;

            foreach (string spec_number in spec_ID)
            {
                end_idx++;
                if (spec_number.Contains("NEED_TEST")) break;

                if (spec_ID[0] == spec_number)
                {
                    spec_IDList.Append(spec_number);
                }
                else
                {
                    spec_IDList.Append("/" + spec_number);
                }
                
                if (end_idx > 2)
                {
                    spec_IDList.Append("...");
                    break;
                }
            }

            spec_IDList.Append(')');
            new_slide.Sub_Desc_Title = new_slide.Sub_Desc_Title + spec_IDList.ToString();

            foreach (Template_PIC pic_item in new_slide.Pictures)
            {
                List<string> remove_duplication = new List<string>();

                if(IsRL) pic_item.markers.Clear();

                for (int i = 0; i < pic_item.file_label.Count; i++)
                {
                    StringBuilder combined_list = new StringBuilder();
                    combined_list.Append(pic_item.file_path[i]);
                    combined_list.Append("|");
                    combined_list.Append(pic_item.file_name[i]);
                    combined_list.Append("|");
                    combined_list.Append(pic_item.file_label[i]);
                    remove_duplication.Add(combined_list.ToString());
                }

                remove_duplication = remove_duplication.Distinct().ToList();
                List<string> New_File_Path = new List<string>();
                List<string> New_File_Name = new List<string>();
                List<string> New_File_Label = new List<string>();

                foreach (string item in remove_duplication)
                {
                    string[] split_List = item.Split('|');
                    New_File_Path.Add(split_List[0].Trim());
                    New_File_Name.Add(split_List[1].Trim());
                    New_File_Label.Add(split_List[2].Trim());
                }

                pic_item.file_path = New_File_Path;
                pic_item.file_name = New_File_Name;
                pic_item.file_label = New_File_Label;

            }
            
            return new_slide;
        }



        public class Template_PIC
        {
            public List<string> file_path = new List<string>();
            public List<string> file_name = new List<string>();
            public List<string> file_label = new List<string>();

            public string Compare_Key;

            public string Title;
            public string S_Parameter;
            public string Y_title;
            public string Freq_start;
            public string Freq_stop;
            public string Mag_start;
            public string Mag_stop;

            public List<string> markers = new List<string>();
            public List<string> SPEC_Start = new List<string>();
            public List<string> SPEC_Stop = new List<string>();
            public List<string> SPEC_value = new List<string>();

            public Template_PIC()
            {
                Clear();
            }

            public void Create_Picture(SnpData item, string DUT_ID, Dictionary<string, string> SnP_Directories, string compareKey_item)
            {
                this.Compare_Key = compareKey_item;
                this.file_name.Add(DUT_ID + item.SNP_File_Name);
                this.file_path.Add(SnP_Directories[DUT_ID]);
                
                string Label = DUT_ID + "_" + item.Temperature + (item.GainMode != "" ? "_" + item.GainMode : "");
                
                if (item.SnpCon_TestID.Contains("RX_Gain_CA_G0"))
                {
                    this.file_label.Add(Label + "_" + item.RX_BAND_mipi);
                }
                else
                {
                    this.file_label.Add(Label);
                }

                Build_PicSubTitle(item); //this.Title, this.Y_title

                this.S_Parameter = item.S_param;
                this.Freq_start = item.Start_Freq.Replace("M", "").Trim();
                this.Freq_stop = item.Stop_Freq.Replace("M", "").Trim();

                Add_markers(this.Freq_start, this.Freq_stop);

                string spec_value = (item.Test_limit_L != "" ? item.Test_limit_L : item.Test_limit_H);
                
                if (item.SnpCon_TestID.Contains("ISO") && spec_value != "")
                {
                    float check_float = 0;
                    if(!float.TryParse(spec_value, out check_float))
                    {
                        spec_value = spec_value;
                    }
                    else
                    {
                        float convert_spec_value = Convert.ToSingle(spec_value.Trim());
                        spec_value = Convert.ToString(Math.Abs(convert_spec_value) * -1);
                    }
                    
                }

                this.SPEC_Start.Add(this.Freq_start);
                this.SPEC_Stop.Add(this.Freq_stop);
                this.SPEC_value.Add(spec_value);
            }

            public void Create_Picture_Multi(SnpData item, List<string> DUT_ID, Dictionary<string, string> SnP_Directories, string compareKey_item)
            {
                this.Compare_Key = compareKey_item;

                foreach (string MultiUnit_A in DUT_ID)
                {
                    this.file_name.Add(MultiUnit_A + item.SNP_File_Name);
                    this.file_path.Add(SnP_Directories[MultiUnit_A]);

                    string Label = MultiUnit_A + "_" + item.Temperature + (item.GainMode != "" ? "_" + item.GainMode : "");
                    if (item.SnpCon_TestID.Contains("RX_Gain_CA_G0"))
                    {
                        this.file_label.Add(Label + "_" + item.RX_BAND_mipi);
                    }
                    else
                    {
                        this.file_label.Add(Label);
                    }
                }

                Build_PicSubTitle(item); //this.Title, this.Y_title

                this.S_Parameter = item.S_param;
                this.Freq_start = item.Start_Freq.Replace("M", "").Trim();
                this.Freq_stop = item.Stop_Freq.Replace("M", "").Trim();

                Add_markers(this.Freq_start, this.Freq_stop);

                string spec_value = (item.Test_limit_L != "" ? item.Test_limit_L : item.Test_limit_H);

                if (item.SnpCon_TestID.Contains("ISO") && spec_value != "")
                {
                    float check_float = 0;
                    if (!float.TryParse(spec_value, out check_float))
                    {
                        spec_value = spec_value;
                    }
                    else
                    {
                        float convert_spec_value = Convert.ToSingle(spec_value.Trim());
                        spec_value = Convert.ToString(Math.Abs(convert_spec_value) * -1);
                    }

                }

                this.SPEC_Start.Add(this.Freq_start);
                this.SPEC_Stop.Add(this.Freq_stop);
                this.SPEC_value.Add(spec_value);
            }

            private void Build_PicSubTitle(SnpData item)
            {
                string Sub_title = "";
                string Y_value = "";

                switch (item.SnpCon_TestID)
                {
                    case "IL":
                        Sub_title = item.Input_Port + " Insertion Loss(" + item.Output_Port + ")";
                        Y_value = "Insertion Loss (dB)";
                        break;
                    case "Input_RL":
                        Sub_title = item.TX_BAND_mipi + " " + item.Input_Port + " Input Return Loss(" + item.Output_Port + ")";
                        Y_value = "Input Return Loss (dB)";
                        break;
                    case "Output_RL":
                        Sub_title = item.TX_BAND_mipi + " " + item.Input_Port + " Output Return Loss(" + item.Output_Port + ")";
                        Y_value = "Output Return Loss (dB)";
                        break;
                    case "TX_Gain_MAX":
                        Sub_title = item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " Gain (" + item.Output_Port + ")";
                        Y_value = "TX Gain (dB)";
                        break;
                    case "RX_Gain_G0":
                        Sub_title = item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " Gain (" + item.Input_Port + ", " + item.Output_Port + ")";
                        Y_value = "RX Gain (dB) - G0 ";
                        break;
                    case "RX_Gain_CA_G0":
                        Sub_title = item.TX_BAND_mipi + "CA Gain(" + item.Input_Port + ", " + item.RX_OUTPUT_mipi + ")";
                        Y_value = "RX Gain (dB) - G0 ";
                        break;
                    case "Gain_Ripple":
                        Sub_title = item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " Gain ripple (" + item.Output_Port + ")";
                        Y_value = "Gain (dB)";
                        break;
                    case "TX_OOB_Gain":                   
                        Sub_title = item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " OOB Gain (" + item.Output_Port + ")";
                        Y_value = "OOB Gain (dB)";
                        break;
                    case "RX_OOB_Gain":
                        Sub_title = item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " (" + item.Input_Port + ") - OOB Gain (" + item.Output_Port + ")"; 
                        Y_value = "OOB Gain (dB)";
                        break;
                    case "ISO:TX, RX":
                        Sub_title = "Active " + item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + " > " + item.RX_BAND_mipi + "RX Isolation (" + item.Output_Port + ")";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:RX, InAct_RX":
                        Sub_title = "Act " + item.TX_BAND_mipi.Replace("_"," ") + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + "> RX " + item.RX_BAND_mipi.Replace("_", "") + "(" + item.Input_Port + ") to Inact " + item.Output_Port;
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:TX, InAct_RX":
                        Sub_title = "Active " + item.TX_BAND_mipi.Replace("_", "") + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + "> Inactvie RX " + item.RX_BAND_mipi.Replace("_", "") + "(" + item.Output_Port + ")";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:InAct_RX, InAct_RX":
                        Sub_title = "Inactive RX(" + item.Input_Port + ") to Inactive RX(" + item.Output_Port + ")";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:ASM, ASM":
                        Sub_title = "Active " + item.Input_Port + "(" + Get_ANT(item.Input_Port, item) + ") To Active " + item.Output_Port + "(" + Get_ANT(item.Output_Port, item) + ") Isolation";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:ASM, InAct_ANT":
                        Sub_title = "Active " + item.Input_Port + "(" + Get_ANT(item.Input_Port, item) + ") To Inactive " + item.Output_Port + " Isolation";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:ASM, InAct_RX":
                        Sub_title = "Active " + item.Input_Port + "(" + Get_ANT(item.Input_Port, item) + ") To Inactive " + item.Output_Port + " Isolation";
                        Y_value = "Isolation (dB)";
                        break;
                    case "REV_ISO:RX, ANT":
                        Sub_title = "Reverse ISO : " + item.RX_OUTPUT_mipi + " To " + item.Output_Port + " Isolation";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:ANT, InAct_ANT":
                        Sub_title = "Active " + item.TX_BAND_mipi + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + "(" + item.Input_Port + ") To Inactive " + item.Output_Port + " Isolation";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:ANT, ANT":
                        Sub_title = "Active " + item.Input_Port +" To Active " + item.Output_Port + " (" + (item.TRX_ON_mipi == "TXRX" ? item.TX_BAND_mipi : item.RX_BAND_mipi) + (item.TRX_ON_mipi == "TXRX" ? "TX = " : "RX = ") + item.TX_OUTPUT_mipi + ")";
                        Y_value = "Isolation (dB)";
                        break;
                    case "ISO:TX, ASM":
                        Sub_title = item.TX_BAND_mipi.Replace("_", "") + (item.TRX_ON_mipi == "TXRX" ? "TX" : "RX") + "(" + item.TX_OUTPUT_mipi + ") To " + item.Output_Port + "(" + Get_ANT(item.Output_Port, item) + ")";
                        Y_value = "Isolation (dB)";
                        break;

                    default:
                        Sub_title = "Band" + item.PA_BAND + " " + item.SnpCon_TestID + "(" + item.Output_Port + ")";
                        break;
                }

                this.Title = Sub_title;
                this.Y_title = Y_value;
            }

            private string Get_ANT(string Port, SnpData item)
            {
                string ant_name = "";
                string compare_str = "";

                if (Port.Contains("2G")) compare_str = "GSM";
                if (Port.Contains("DRX")) compare_str = "DRX";
                if (Port.Contains("FROM_UAT")) compare_str = "FROM_UAT";
                if (Port.Contains("MIMO")) compare_str = "MIMO";
                if (Port.Contains("LMB")) compare_str = "LMB";

                if (compare_str == item.ASM_ANT1) ant_name = "ANT1";
                if (compare_str == item.ASM_ANT2) ant_name = "ANT2";
                if (compare_str == item.ASM_ANT3) ant_name = "UAT";

                return ant_name;
            }

            public void Add_ToPicture(SnpData item, string DUT_ID, Dictionary<string, string> SnP_Directories)
            {
                this.file_name.Add(DUT_ID + item.SNP_File_Name);
                this.file_path.Add(SnP_Directories[DUT_ID]);
                //this.file_label.Add(DUT_ID + "_" + item.Temperature + (item.GainMode != "" ? "_" + item.GainMode : "");
                
                string Label = DUT_ID + "_" + item.Temperature + (item.GainMode != "" ? "_" + item.GainMode : "");
                if (item.SnpCon_TestID.Contains("RX_Gain_CA_G0"))
                {
                    this.file_label.Add(Label + "_" + item.RX_BAND_mipi);
                }
                else
                {
                    this.file_label.Add(Label);
                }

                string start_f = item.Start_Freq.Replace("M", "").Trim();
                string stop_f = item.Stop_Freq.Replace("M", "").Trim();
                Add_markers(start_f, stop_f);

                string spec_value = (item.Test_limit_L != "" ? item.Test_limit_L : item.Test_limit_H);

                if (item.SnpCon_TestID.Contains("ISO") && spec_value != "")
                {
                    float check_float = 0;
                    if (!float.TryParse(spec_value, out check_float))
                    {
                        spec_value = spec_value;
                    }
                    else
                    {
                        float convert_spec_value = Convert.ToSingle(spec_value.Trim());
                        spec_value = Convert.ToString(Math.Abs(convert_spec_value) * -1);
                    }
                }

                this.SPEC_Start.Add(start_f);
                this.SPEC_Stop.Add(stop_f);
                this.SPEC_value.Add(spec_value);
            }

            public void Add_ToPicture_Multi(SnpData item, List<string> DUT_ID, Dictionary<string, string> SnP_Directories)
            {
                foreach (string MultiUnit_A in DUT_ID)
                {
                    this.file_name.Add(MultiUnit_A + item.SNP_File_Name);
                    this.file_path.Add(SnP_Directories[MultiUnit_A]);

                    string Label = MultiUnit_A + "_" + item.Temperature + (item.GainMode != "" ? "_" + item.GainMode : "");
                    if (item.SnpCon_TestID.Contains("RX_Gain_CA_G0"))
                    {
                        this.file_label.Add(Label + "_" + item.RX_BAND_mipi);
                    }
                    else
                    {
                        this.file_label.Add(Label);
                    }
                }

                string start_f = item.Start_Freq.Replace("M", "").Trim();
                string stop_f = item.Stop_Freq.Replace("M", "").Trim();
                Add_markers(start_f, stop_f);

                string spec_value = (item.Test_limit_L != "" ? item.Test_limit_L : item.Test_limit_H);

                if (item.SnpCon_TestID.Contains("ISO") && spec_value != "")
                {
                    float check_float = 0;
                    if (!float.TryParse(spec_value, out check_float))
                    {
                        spec_value = spec_value;
                    }
                    else
                    {
                        float convert_spec_value = Convert.ToSingle(spec_value.Trim());
                        spec_value = Convert.ToString(Math.Abs(convert_spec_value) * -1);
                    }
                }

                this.SPEC_Start.Add(start_f);
                this.SPEC_Stop.Add(stop_f);
                this.SPEC_value.Add(spec_value);
            }

            private void Add_markers(string start_freq, string stop_freq)
            {
                float I_freq = Convert.ToSingle(start_freq);
                float F_freq = Convert.ToSingle(stop_freq);
                float Range = F_freq - I_freq;
                float MID_freq = I_freq + (Range) / 2;

                if (this.markers.Count == 0)
                {
                    if (I_freq - (Range / 4) > 0f) this.markers.Add(Convert.ToString(I_freq - Range / 4));
                    this.markers.Add(Convert.ToString(I_freq));
                    this.markers.Add(Convert.ToString(MID_freq));
                    this.markers.Add(Convert.ToString(F_freq));
                    this.markers.Add(Convert.ToString(F_freq + Range / 4));
                }
                else
                {
                    this.markers.Add(Convert.ToString(I_freq));
                    this.markers.Add(Convert.ToString(F_freq));
                    this.markers = this.markers.Distinct().ToList();

                    float min_F = 900000f;
                    float max_F = 0f;

                    foreach (string frequency in this.markers)
                    {
                        float target_F = Convert.ToSingle(frequency);
                        if (target_F <= min_F) min_F = target_F;
                        if (target_F >= max_F) max_F = target_F;
                    }

                    this.Freq_start = Convert.ToString(min_F);
                    this.Freq_stop = Convert.ToString(max_F);
                }
            }

            private void Clear()
            {
                this.Compare_Key = "";
                this.Title = "";
                this.S_Parameter = "";
                this.Y_title = "";
                this.Freq_start = "";
                this.Freq_stop = "";
                this.Mag_start = "";
                this.Mag_stop = "";
                this.file_path.Clear();
                this.file_name.Clear();
                this.file_label.Clear();
                this.markers.Clear();
                this.SPEC_Start.Clear();
                this.SPEC_Stop.Clear();
                this.SPEC_value.Clear();
            }
            public Template_PIC Clone()
            {
                Template_PIC cloneType = new Template_PIC();

                cloneType.Compare_Key = this.Compare_Key;
                cloneType.Title = this.Title;
                cloneType.S_Parameter = this.S_Parameter;
                cloneType.Y_title = this.Y_title;
                cloneType.Freq_start = this.Freq_start;
                cloneType.Freq_stop = this.Freq_stop;
                cloneType.Mag_start = this.Mag_start;
                cloneType.Mag_stop = this.Mag_stop;

                cloneType.file_path = this.file_path;
                cloneType.file_name = this.file_name;
                cloneType.file_label = this.file_label;
                cloneType.markers = this.markers;
                cloneType.SPEC_Start = this.SPEC_Start;
                cloneType.SPEC_Stop = this.SPEC_Stop;
                cloneType.SPEC_value = this.SPEC_value;

                return cloneType;
            }
        }

        public class Template_Slide
        {
            public string Main_Title;
            public string Sub_Desc_Title;
            public List<Template_PIC> Pictures = new List<Template_PIC>();
            public int picture_cnt = 6;
            public string slide_option;

            public Template_Slide()
            {
                Clear();
            }

            private void Clear()
            {
                this.Main_Title = "";
                this.Sub_Desc_Title = "";
                this.slide_option = "Option6";
                this.Pictures.Clear();
            }
            public Template_Slide Clone()
            {
                Template_Slide cloneType = new Template_Slide();

                cloneType.Main_Title = this.Main_Title.ToString();
                cloneType.Sub_Desc_Title = this.Sub_Desc_Title.ToString();

                foreach (Template_PIC item in this.Pictures)
                {
                    Template_PIC new_item = new Template_PIC();
                    new_item = item;
                    cloneType.Pictures.Add(new_item);
                }

                cloneType.picture_cnt = 6;
                cloneType.slide_option = this.slide_option.ToString();

                return cloneType;
            }

        }

        private List<string> Get_CheckedItem(CheckedListBox CBox)
        {
            List<string> Checked_List = new List<string>();

            for (int i = 0; i < CBox.Items.Count; i++)
            {
                if (CBox.GetItemChecked(i)) Checked_List.Add(CBox.Items[i].ToString());
            }

            return Checked_List;
        }

        private void SelectClear_Click(object sender, EventArgs e)
        {
            if (this.CBox_Bands.Items.Count == 0)
            {

            }
            else
            {
                bool initial_item_checked = this.CBox_Bands.GetItemChecked(0);

                if(initial_item_checked)
                {
                    for (int i = 0; i < this.CBox_Bands.Items.Count; i++)
                    {
                        this.CBox_Bands.SetItemChecked(i, false);
                    }
                }
                else
                {
                    for (int i = 0; i < this.CBox_Bands.Items.Count; i++)
                    {
                        this.CBox_Bands.SetItemChecked(i, true);
                    }
                }
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.CBox_UnitPath.Items.Count == 0)
            {

            }
            else
            {
                bool initial_item_checked = this.CBox_UnitPath.GetItemChecked(0);

                if (initial_item_checked)
                {
                    for (int i = 0; i < this.CBox_UnitPath.Items.Count; i++)
                    {
                        this.CBox_UnitPath.SetItemChecked(i, false);
                    }
                }
                else
                {
                    for (int i = 0; i < this.CBox_UnitPath.Items.Count; i++)
                    {
                        this.CBox_UnitPath.SetItemChecked(i, true);
                    }
                }
            }
        }

    }
}
