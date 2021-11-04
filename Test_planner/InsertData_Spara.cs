using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics; //Process to kill();
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Base;
using Excel = Microsoft.Office.Interop.Excel; //for excel control
using Test_Planner;

namespace S_para_planner
{
    class Spara_data
    {
        public string NUM_Testplan;
        public string Sheet_Band;
        public string NUM_Spec;
        public string Temperature;

        public string TestID;
        public string Name_Test;
        public string Method_Data;
        public string Freq_Start;
        public string Freq_Stop;
        public string Mode_RXgain;
        public string In_Port_Test;
        public string Out_Port_Test;
        public string SnP_info;
        public string Data_Search;

        public string Data_Freq;
        public string Data_Value;

        public string Sign_Convert;

        public Spara_data()
        {
            clear();
        }

        public void clear()
        {
            this.NUM_Testplan = "";
            this.Sheet_Band = "";
            this.NUM_Spec = "";
            this.Temperature = "";

            this.TestID = "";
            this.Name_Test = "";
            this.Method_Data = "";
            this.Freq_Start = "";
            this.Freq_Stop = "";
            this.Mode_RXgain = "";
            this.In_Port_Test = "";
            this.Out_Port_Test = "";
            this.SnP_info = "";
            this.Data_Search = "";

            this.Data_Freq = "";
            this.Data_Value = "";
            this.Sign_Convert = "";
        }

        public void ImportData(List<string> header, List<string> Data)
        {
            if (header.Count != Data.Count) return;
            
            int Index_Data = 0;
            foreach (string item in header)
            {
                switch (item.Trim())
                {
                    case "Test_Plan_Num": 
                        this.NUM_Testplan = Data[Index_Data]; break;
                    case "Spec Sheet Name":
                        this.Sheet_Band = Data[Index_Data]; break;
                    case "Para.Spec":
                        this.NUM_Spec = Data[Index_Data]; break;
                    case "Set_Temp":
                        this.Temperature = Data[Index_Data]; break;
                    case "Parameter Header":
                        this.Name_Test = Data[Index_Data]; break;
                    case "Test Parameter":
                        this.Method_Data = Data[Index_Data]; break;
                    case "Start_Freq":
                        this.Freq_Start = Data[Index_Data]; break;
                    case "Stop_Freq":
                        this.Freq_Stop = Data[Index_Data]; break;
                    case "LNA_GAIN":
                        this.Mode_RXgain = Data[Index_Data]; break;
                    case "Input Port":
                        this.In_Port_Test = Data[Index_Data]; break;
                    case "Output Port":
                        this.Out_Port_Test = Data[Index_Data]; break;
                    case "DM_S-Param":
                        this.SnP_info = Data[Index_Data]; break;
                    case "Search_Method":
                        this.Data_Search = Data[Index_Data]; break;
                    case "Result_Freq":
                        this.Data_Freq = Data[Index_Data]; break;
                    case "Test_Result":
                        this.Data_Value = Data[Index_Data]; break;
                    case "Convert_SIGN_FOR_ISO":
                        this.Sign_Convert = Data[Index_Data]; break;
                    default:
                        break;
                }
                Index_Data++;
            }

            if (this.Name_Test != "") this.TestID = GetTestID(this.Name_Test);
        }

        private string GetTestID(string TestName)
        {
            string Test_ID = "";
            string[] TempStr = TestName.Split(']');
            Test_ID = TempStr[0].Replace("[", "").Trim();
            return Test_ID;
        }

    }

    public partial class InsertData_Spara : Form
    {
        string PathDefault = "C:\\ProgramData\\FlexTest\\GENTLE_BREED";
        public int Excel_Proc_ID = 0;
        public int Excel_Proc_ID_CM = 0;
        //public Excel.Workbook Spara_Workbook;
        //public Excel.Workbook CM_Sheet;

        public Excel_File Spara_Excel;
        public Excel_File CM_Excel;

        bool Opened_SparaData = false;
        bool Opened_CMSheet = false;

        List<string> SparaTest_Sheet = new List<string>();
        List<string> CM_Target_Sheet = new List<string>();


        public InsertData_Spara()
        {
            InitializeComponent();
            ProgressBar_Init(200);
            this.Show_Path_Spara.Text = "Please select S-para test plan";
            this.Show_Path_CM.Text = "Please select CM sheet to insert Data";
            this.BTN_Load_CM.BackColor = Color.LightGray;
            this.BTN_Load_CM.Enabled = false;
            this.BTN_Insert.Enabled = false;
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

        private void BTN_LoadData_Spara_Click(object sender, EventArgs e)
        {
            string PFN_SparaData = "";
            ProgressBar_Init(2);
            this.BTN_LoadData_Spara.BackColor = Color.Yellow;
            this.Show_Path_Spara.Text = "Loading......";
            this.Show_Path_Spara.BackColor = Color.Yellow;

            this.Opened_SparaData = false;
            this.Opened_CMSheet = false;
            this.BTN_Load_CM.Enabled = false;
            this.BTN_Insert.Enabled = false;

            try
            {
                System.Windows.Forms.OpenFileDialog OpenDialogEntity = new System.Windows.Forms.OpenFileDialog();


                OpenDialogEntity.InitialDirectory = (Directory.Exists(this.PathDefault) ? this.PathDefault : System.Environment.SpecialFolder.MyComputer.ToString());
                OpenDialogEntity.Filter = "Excel Files (.xlsx;.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                OpenDialogEntity.FilterIndex = 1;
                OpenDialogEntity.Multiselect = false;
                OpenDialogEntity.CheckFileExists = false;

                if (OpenDialogEntity.ShowDialog() == DialogResult.OK)
                {
                    PFN_SparaData = OpenDialogEntity.FileName;
                    ProgBar_execute_step();
                    this.PathDefault = PFN_SparaData;
                }
                else
                {
                    return;
                }
            }
            catch
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open Spara test data from path");
                ErrMsg.AppendFormat("\nPlease check open file location and format");
                MessageBox.Show("Error on file loading", ErrMsg.ToString());
                //Environment.Exit(0);
            }

            Excel_File SparaData = new Excel_File(false, PFN_SparaData);
            this.Excel_Proc_ID = SparaData.ProcID;
            this.BTN_LoadData_Spara.BackColor = Color.YellowGreen;
            this.Show_Path_Spara.Text = PFN_SparaData;
            this.Show_Path_Spara.BackColor = Color.AliceBlue;
            ProgBar_execute_step();
            this.BTN_Load_CM.Enabled = true;

            this.Opened_SparaData = true;

            SparaData.Show(false);
            SparaData.App.ScreenUpdating = false;

            List<string> Spara_DataSheet_List = SparaData.get_ExcelSheet_List();

            foreach (var item in Spara_DataSheet_List)
            {
                bool IsChecked = true;
                switch (item.Trim().ToUpper())
                {
                    case "CONDITION_FBAR":
                    case "TEST_DATA":
                    case "PORT_INFO":
                    case "MIPI_INFO":
                        IsChecked = false;
                        break;
                    default:
                        break;
                }

                this.CBox_SparaData.Items.Add(item, IsChecked);
            }

            this.Spara_Excel = SparaData;
        }

        private void BTN_Insert_Click(object sender, EventArgs e)
        {
            if (!Opened_SparaData & !Opened_CMSheet) return;

            List<string> Target_SampleSheet = new List<string>();
            List<string> Target_sheet_List = new List<string>();

            for (int i = 0; i < this.CBox_SparaData.Items.Count; i++)
            {
                if (this.CBox_SparaData.GetItemChecked(i)) Target_SampleSheet.Add(this.CBox_SparaData.Items[i].ToString());       
            }

            Popup_ProgressBar bar1 = new Popup_ProgressBar(Target_SampleSheet.Count);
            bar1.Show();

            this.SparaTest_Sheet = Target_SampleSheet;
            int start_row = 1;
            int end_col = 100;
            int end_row = 100000;

            List<string> Header = this.Spara_Excel.Quick_FindHeader(Target_SampleSheet[0], "Test_Plan_Num", ref start_row, ref end_col, ref end_row);

            Dictionary<string, Dictionary<string, Dictionary<string, List<Spara_data>>>> Dic_SparaData = new Dictionary<string, Dictionary<string, Dictionary<string, List<Spara_data>>>>();

            //Dictionary < Sample, Band, SpecID, Data>

            foreach (string Sample_Data_Sheet in Target_SampleSheet)
            {
                Dictionary<string, List<Spara_data>> Dic_SparaData_Bands = new Dictionary<string, List<Spara_data>>();

                StringBuilder BarMsg = new StringBuilder();
                BarMsg.AppendFormat("Load sample - {0} data from Spara plan", Sample_Data_Sheet);
                bar1.execute_step("Import Spara data", BarMsg.ToString());

                string[,] Full_data = this.Spara_Excel.ReadData_From_WorkSheet(Sample_Data_Sheet, start_row + 1, end_row, 1, end_col);

                for (int i = 0; i < (end_row - start_row); i++)
                {
                    List<string> DataRow = new List<string>();
                    
                    for (int k = 0; k < end_col; k++)
                    {
                        DataRow.Add(Full_data[i, k]);
                    }

                    Spara_data SparaData_Row = new Spara_data();
                    SparaData_Row.ImportData(Header, DataRow);

                    if (SparaData_Row.Method_Data.Trim().ToUpper() != "SETUP_TRIG" && 
                        SparaData_Row.NUM_Spec.Trim().ToUpper() != "NEED_TEST" &&
                        SparaData_Row.NUM_Spec.Trim().ToUpper() != "")
                    {
                        if (Dic_SparaData_Bands.ContainsKey(SparaData_Row.Sheet_Band))
                        {
                            Dic_SparaData_Bands[SparaData_Row.Sheet_Band].Add(SparaData_Row);
                        }
                        else
                        {
                            List<Spara_data> new_SparaList = new List<Spara_data>();
                            new_SparaList.Add(SparaData_Row);
                            Dic_SparaData_Bands.Add(SparaData_Row.Sheet_Band, new_SparaList);
                        }
                    }
                }

                Dictionary<string, Dictionary<string, List<Spara_data>>> Dic_SparaData_BandwithSpec = new Dictionary<string, Dictionary<string, List<Spara_data>>>();

                foreach (var key_band in Dic_SparaData_Bands.Keys)
                {
                    Dictionary<string, List<Spara_data>> Spec_SparaData = new Dictionary<string, List<Spara_data>>();
                    Target_sheet_List.Add(key_band);
                    Target_sheet_List = Target_sheet_List.Distinct().ToList();

                    StringBuilder BarMsg2 = new StringBuilder();
                    BarMsg2.AppendFormat("{0} : Import {1} data from Spara plan", Sample_Data_Sheet, key_band);
                    bar1.ShowMSG(BarMsg2.ToString());

                    foreach (var item in Dic_SparaData_Bands[key_band])
                    {
                        if (Spec_SparaData.ContainsKey(item.NUM_Spec))
                        {
                            Spec_SparaData[item.NUM_Spec].Add(item);
                        }
                        else
                        {
                            List<Spara_data> new_SparaListwSpec = new List<Spara_data>();
                            new_SparaListwSpec.Add(item);
                            Spec_SparaData.Add(item.NUM_Spec, new_SparaListwSpec);
                        }
                    }

                    Dic_SparaData_BandwithSpec.Add(key_band, Spec_SparaData);
                }

                Dic_SparaData.Add(Sample_Data_Sheet, Dic_SparaData_BandwithSpec);

            }
            //============================ Write to CM sheet ============================//

            Excel_File WorstCon = new Excel_File(true);
            WorstCon.Show(true);
            WorstCon.App.ScreenUpdating = true;
            WorstCon.Add_Sheet("Worst_Con");
            WorstCon.Add_Sheet("WorstConDetail");
            WorstCon.Clear_Sheet();

            List<string> WorstCondition_Num = new List<string>();
            List<string> WorstCondition_Num_Detail = new List<string>();
            List<string> WorstCondition_Num_Detail_Data = new List<string>();

            WorstCondition_Num.Add("WORST_CASES");
            WorstCondition_Num_Detail.Add("WORST_CASES_DETAIL");
            WorstCondition_Num_Detail_Data.Add("DATA");

            bar1.Init(Target_sheet_List.Count);
            int progress_step = 1;

            foreach (string Sheet_name in Target_sheet_List)
            {
                string call_sheet = "";
                bool Check_Sheet_exist = this.CM_Excel.Find_ExcelSheet(Sheet_name, ref call_sheet);
                if(Check_Sheet_exist)
                {
                    double Progress_rate = ((double)progress_step / (double)Target_sheet_List.Count) * 100;

                    StringBuilder Rate = new StringBuilder();
                    Rate.AppendFormat("Progress : {0}", Progress_rate.ToString("0.0"));

                    StringBuilder BarMsg3 = new StringBuilder();
                    BarMsg3.AppendFormat("Writing Result : data to {0}", Sheet_name);
                    bar1.execute_step(Rate.ToString(),BarMsg3.ToString());
                    progress_step++;


                    Dictionary<string, int> SpecID_DIC = new Dictionary<string, int>();
                    Dictionary<string, int> CM_Header = this.CM_Excel.Find_CM_Header_Index(call_sheet, 1, 4, 150, ref SpecID_DIC);
                    //X axis = CM_Header, Y axis = SpecID_DIC

                    Dictionary<string, Dictionary<string, List<Spara_data>>> Each_SampleData_DIC = new Dictionary<string, Dictionary<string, List<Spara_data>>>();
                    //sample number, spec ID, spara_data

                    foreach (string Sample_Data in Dic_SparaData.Keys)
                    {
                        if (Dic_SparaData[Sample_Data].Count != 0)
                        {
                            Each_SampleData_DIC.Add(Sample_Data, Dic_SparaData[Sample_Data][Sheet_name]);
                        }
                    }

                    Dictionary<string, List<string>> Spec_Rawdata_DIC = new Dictionary<string, List<string>>();
                    Dictionary<string, List<string>> Spec_Condition_DIC = new Dictionary<string, List<string>>();

                    foreach (string SpecID in SpecID_DIC.Keys)
                    {
                        List<string> Worst_data_units = new List<string>();
                        List<string> Worst_Conditions = new List<string>();

                        foreach (string UnitNum in Each_SampleData_DIC.Keys)
                        {
                            if (Each_SampleData_DIC[UnitNum].ContainsKey(SpecID))
                            {
                                List<Spara_data> Candidate_data = Each_SampleData_DIC[UnitNum][SpecID];
                                List<string> Get_MinMax = Extract_Worst_Condition(Candidate_data);

                                string Data_Selection = "";


                                if (Candidate_data[0].Sign_Convert.Trim().ToUpper() == "" ||
                                    Candidate_data[0].Sign_Convert.Trim().ToUpper() == "ON")
                                {
                                    Data_Selection = Determin_Worst_MinMax(Get_MinMax[0]);
                                }
                                else
                                {
                                    Data_Selection = Determin_Worst_MinMax_with_SignConvertOFF(Get_MinMax[0]);
                                }

                                switch (Data_Selection)
                                {
                                    case "MIN":
                                        Worst_data_units.Add(Get_MinMax[1]); // [1] Min Value in spec
                                        Worst_data_units.Add("");            // [2] Max Value in spec
                                        Worst_Conditions.Add(Get_MinMax[3]); // [3] Min Value condition // [4] Max Value condition
                                        break;
                                    case "MAX":
                                        Worst_data_units.Add("");
                                        Worst_data_units.Add(Get_MinMax[2]);
                                        Worst_Conditions.Add(Get_MinMax[4]);
                                        break;
                                    case "BOTH":
                                        Worst_data_units.Add(Get_MinMax[1]);
                                        Worst_data_units.Add(Get_MinMax[2]);
                                        Worst_Conditions.Add("Min: " + Get_MinMax[3] + " //// Max: " + Get_MinMax[4]);
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }

                        if (Worst_data_units.Count != 0) Spec_Rawdata_DIC.Add(SpecID, Worst_data_units);
                        if (Worst_Conditions.Count != 0) Spec_Condition_DIC.Add(SpecID, Worst_Conditions);
                    }

                    //Write to CM sheet Excel with Pass fail determine
                    this.CM_Excel.Write_SampleData_To_CM(call_sheet, CM_Header, SpecID_DIC, Spec_Rawdata_DIC, Spec_Condition_DIC);

                    foreach (var Spec_ID in Spec_Condition_DIC.Keys)
                    {
                        if(Spec_Condition_DIC[Spec_ID][0].Contains("/"))
                        {
                            string TempSTR = Spec_Condition_DIC[Spec_ID][0].Replace("////", "/");
                            string[] splitSTR = TempSTR.Split('/');
                            for (int i = 0; i < splitSTR.Length; i++)
                            {
                                string[] splitSubSTR = splitSTR[i].Trim().Split(',');
                                for (int j = 0; j < splitSubSTR.Length; j++)
                                {
                                    if (splitSubSTR[j].Contains("Cond_Num =")) WorstCondition_Num.Add(Regex.Match(splitSubSTR[j], @"\d+").Value);
                                }
                            }
                        }
                        else
                        {
                            string[] splitSTR = Spec_Condition_DIC[Spec_ID][0].Split(',');
                            for (int i = 0; i < splitSTR.Length; i++)
                            {
                                if (splitSTR[i].Contains("Cond_Num =")) WorstCondition_Num.Add(Regex.Match(splitSTR[i], @"\d+").Value);
                            }
                        }

                        foreach (string worst_condition in Spec_Condition_DIC[Spec_ID])
                        {
                            if (Spec_Condition_DIC[Spec_ID][0].Contains("/"))
                            {
                                string TempSTR = Spec_Condition_DIC[Spec_ID][0].Replace("////", "/");
                                string[] splitSTR = TempSTR.Split('/');

                                for (int i = 0; i < splitSTR.Length; i++)
                                {
                                    WorstCondition_Num_Detail.Add(splitSTR[i].Trim());
                                }
                            }
                            else
                            {
                                WorstCondition_Num_Detail.Add(worst_condition);
                            }
                                
                        }

                        foreach (string data in Spec_Rawdata_DIC[Spec_ID])
                        {
                            if (data != "") WorstCondition_Num_Detail_Data.Add(data);
                        }
                    }
                }
            }

            string temp = "here";
            //Write worst condition here;

            Excel.Worksheet ws_worst = WorstCon.getSheet("Worst_Con");
            WorstCon.WriteData_1D_Col(ws_worst, 1, 1, WorstCondition_Num);
            Excel.Worksheet ws_worst2 = WorstCon.getSheet("WorstConDetail");
            WorstCon.WriteData_1D_Col(ws_worst2, 1, 1, WorstCondition_Num_Detail);
            WorstCon.WriteData_1D_Col(ws_worst2, 1, 2, WorstCondition_Num_Detail_Data);

            bar1.Done(this.CM_Excel);
            this.Spara_Excel.Quit();
            //this.CM_Excel.Activate_App(true);
        }

        private string Determin_Worst_MinMax(string Test_ID)
        {
            string whichOne_IsWorst = "MIN";

            switch (Test_ID)
            {
                case "IL":
                case "Gain_Ripple":
                case "TX_OOB_Gain":
                case "RX_OOB_Gain":
                case "Group_Delay":
                case "Input_VSWR":
                    whichOne_IsWorst = "MAX";
                    break;
                case "Input_RL":
                case "Output_RL":
                case "K_factor":
                case "MU_factor":
                    whichOne_IsWorst = "MIN";
                    break;
                case "Phase_Delta":
                    whichOne_IsWorst = "BOTH";
                    break;
                default:
                    break;
            }

            if(Test_ID.Contains("RX_Gain_")) whichOne_IsWorst = "BOTH"; //RX_Gain_{Mode} or RX_Gain_CA_{Mode}

            return whichOne_IsWorst;

        }

        private string Determin_Worst_MinMax_with_SignConvertOFF(string Test_ID)
        {
            string whichOne_IsWorst = "MIN";
            if (Test_ID.ToUpper().Contains("ISO:")) whichOne_IsWorst = "MAX";

            switch (Test_ID)
            {
                case "Input_RL":
                case "Output_RL":
                case "Gain_Ripple":
                case "TX_OOB_Gain":
                case "RX_OOB_Gain":
                case "Group_Delay":
                case "Input_VSWR":
                    whichOne_IsWorst = "MAX";
                    break;
                    
                case "IL":
                case "K_factor":
                case "MU_factor":
                    whichOne_IsWorst = "MIN";
                    break;

                case "Phase_Delta":
                    whichOne_IsWorst = "BOTH";
                    break;

                default:
                    break;
            }

            if (Test_ID.Contains("RX_Gain_")) whichOne_IsWorst = "BOTH"; //RX_Gain_{Mode} or RX_Gain_CA_{Mode}
            if (Test_ID.ToUpper().Contains("ISO:")) whichOne_IsWorst = "MAX"; //in case sign_Convert == "OFF"

            return whichOne_IsWorst;

        }

        private List<string> Extract_Worst_Condition(List<Spara_data> Datas)
        {
            List<string> Final_result_worst = new List<string>();
            List<double> Data_array = new List<double>();
            List<string> Condition_array = new List<string>();

            if (Datas.Count != 0) Final_result_worst.Add(Datas[0].TestID.Trim());

            foreach (Spara_data data_Set in Datas)
            {
                Data_array.Add(Convert.ToDouble(data_Set.Data_Value));

                StringBuilder condition = new StringBuilder();
                condition.AppendFormat("{0}, Temp = {1}, Port1 = {2}, Port2 = {3}, Data Freq = {4} MHz", data_Set.NUM_Spec, data_Set.Temperature, data_Set.In_Port_Test, data_Set.Out_Port_Test, Convert.ToDouble(data_Set.Data_Freq) / 1000000);
                condition.AppendFormat(" (Range: {0} to {1})", data_Set.Freq_Start, data_Set.Freq_Stop);
                if(data_Set.Mode_RXgain!="") condition.AppendFormat(", GainMode = {0}",data_Set.Mode_RXgain);
                condition.AppendFormat(", SnP = {0}, Cond_Num = {1}", data_Set.SnP_info, data_Set.NUM_Testplan);
                string data_condition = condition.ToString();
                Condition_array.Add(data_condition);
            }

            int index_min = Data_array.IndexOf(Data_array.Min());
            int index_max = Data_array.IndexOf(Data_array.Max());

            //Element 0 = Min, Element 1 = Max, Element 2 = Condition
            Final_result_worst.Add(Convert.ToString(Data_array.Min()));
            Final_result_worst.Add(Convert.ToString(Data_array.Max()));
            Final_result_worst.Add(Condition_array[index_min]);
            Final_result_worst.Add(Condition_array[index_max]);

            return Final_result_worst;
            // [0] Test type ID
            // [1] Min Value in spec
            // [2] Max Value in spec
            // [3] Min Value condition
            // [4] Max Value condition
        }

        private void BTN_Load_CM_Click(object sender, EventArgs e)
        {
            string PFN_CM_ToWrite = "";
            ProgressBar_Init(2);
            this.BTN_Load_CM.BackColor = Color.Yellow;
            this.Show_Path_CM.Text = "Loading......";
            this.Show_Path_CM.BackColor = Color.Yellow;

            try
            {
                System.Windows.Forms.OpenFileDialog OpenDialogEntity = new System.Windows.Forms.OpenFileDialog();


                OpenDialogEntity.InitialDirectory = (Directory.Exists(this.PathDefault) ? this.PathDefault : System.Environment.SpecialFolder.MyComputer.ToString());
                OpenDialogEntity.Filter = "Excel Files (.xlsx;.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                OpenDialogEntity.FilterIndex = 1;
                OpenDialogEntity.Multiselect = false;
                OpenDialogEntity.CheckFileExists = false;

                if (OpenDialogEntity.ShowDialog() == DialogResult.OK)
                {
                    PFN_CM_ToWrite = OpenDialogEntity.FileName;
                    ProgBar_execute_step();
                }
                else
                {
                    return;
                }
            }
            catch
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open Spara test data from path");
                ErrMsg.AppendFormat("\nPlease check open file location and format");
                MessageBox.Show("Error on file loading", ErrMsg.ToString());
                //Environment.Exit(0);
            }

            Excel_File CMToWrite = new Excel_File(false, PFN_CM_ToWrite);
            this.Excel_Proc_ID_CM = CMToWrite.ProcID;
            this.BTN_Load_CM.BackColor = Color.YellowGreen;
            this.Show_Path_CM.Text = PFN_CM_ToWrite;
            this.Show_Path_CM.BackColor = Color.AliceBlue;
            this.Opened_CMSheet = true;
            ProgBar_execute_step();

            CMToWrite.Show(false);
            CMToWrite.App.ScreenUpdating = false;

            List<string> CM_SubSheet_List = CMToWrite.get_ExcelSheet_List();

            this.CM_Excel = CMToWrite;
            this.CM_Target_Sheet = CM_SubSheet_List;
            this.BTN_Insert.Enabled = true;
            this.BTN_Insert.BackColor = Color.Azure;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.CBox_SparaData.Items.Count == 0)
            {

            }
            else
            {
                bool initial_item_checked = this.CBox_SparaData.GetItemChecked(0);

                if (initial_item_checked)
                {
                    for (int i = 0; i < this.CBox_SparaData.Items.Count; i++)
                    {
                        this.CBox_SparaData.SetItemChecked(i, false);
                    }
                }
                else
                {
                    for (int i = 0; i < this.CBox_SparaData.Items.Count; i++)
                    {
                        this.CBox_SparaData.SetItemChecked(i, true);
                    }
                }
            }
        }

    }
}
