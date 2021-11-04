using System;
using System.IO; //Directory and files
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics; //Using application in system (via notepad) with Process methods
using Excel_Base;
using FlexTestLib.MsgBox;
using System.Linq;
using S_para_planner;
using System.Runtime.CompilerServices;

namespace Test_Planner
{
    public partial class Form1 : Form
    {
        public Excel_File CM_Sheet;
        public Excel_File Spara_Plan;
        public Excel_File PPT_source;
        public Excel_File PPT_Plan;
        private Timer CM_Load_Timer;
        Dictionary<string, Dictionary<string, string>> Project_Path = new Dictionary<string, Dictionary<string, string>>();

        public string CurrentProject = "";
        public string Default_CM_Path = "";
        public string Default_Port_Path = "";

        public bool Load_Initial_CM_Complete = false;

        public Form1()
        {
            InitializeComponent();

            this.Project_Path = Get_ProjectMenu();
            this.CurrentProject = Check_default_Project(this.Project_Path);
            textBox1.AppendText("Current Project = " + this.CurrentProject + "\r\n");
            
            Load_public_INI(this.Project_Path, this.CurrentProject);
        }

        private void Set_GlobalPath(string CM_Setting_Path, string Port_Setting_Path)
        {
            Globals.CMsheet_INI_Dir = CM_Setting_Path;
            Globals.Spara_Info = Port_Setting_Path;
            Globals.PortInfo_INI_Dir = Port_Setting_Path;
        }

        private string Check_default_Project(Dictionary<string, Dictionary<string, string>> Project_Path)
        {
            string find_project = "";

            foreach (string Project_name in Project_Path.Keys)
            {
                bool find_CMsetting = false;
                bool find_Portsetting = false;

                foreach (string Paths in Project_Path[Project_name].Values)
                {
                    if (Paths == this.Default_CM_Path) find_CMsetting = true;
                    if (Paths == this.Default_Port_Path) find_Portsetting = true;
                }

                if (find_CMsetting && find_Portsetting)
                {
                    find_project = Project_name;
                    return find_project;
                }
            }

            return find_project;
        }

        public void Load_public_INI(Dictionary<string, Dictionary<string, string>> Project_Path, string currentProject)
        {
            List<string> Project_Menus = new List<string>();
            foreach (var item in Project_Path.Keys)
            {
                Project_Menus.Add(item);
            }

            int ID = 1000;

            foreach (var item in Project_Menus)
            {
                ToolStripMenuItem Project_name = new ToolStripMenuItem(item);
                Project_name.Tag = ID; ID++;

                if (Project_name.Text == currentProject) Project_name.Checked = true;

                this.projectsToolStripMenuItem.DropDownItems.Add(Project_name);
                Project_name.Click += new EventHandler(Set_Project);
            }
        }

        public void Set_Project(object sender, EventArgs e)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;

            foreach (ToolStripMenuItem Dropdown_menus in this.projectsToolStripMenuItem.DropDownItems)
            {
                if (Dropdown_menus.Checked) Dropdown_menus.Checked = false;
            } 

            foreach (var Project_name in this.Project_Path.Keys)
            {
                if(item.Text == Project_name)
                {
                    bool CMsetting_exist = true;
                    bool Portsetting_exist = true;

                    if (!File.Exists(this.Project_Path[Project_name]["CM_Setting"]))
                    {
                        CMsetting_exist = false;
                    }

                    string temp = this.Project_Path[Project_name]["PORT_Setting"];

                    if (!File.Exists(this.Project_Path[Project_name]["PORT_Setting"]))
                    {
                        Portsetting_exist = false;
                    }

                    if (CMsetting_exist && Portsetting_exist)
                    {
                        item.Checked = true;
                        this.Default_CM_Path = this.Project_Path[Project_name]["CM_Setting"];
                        this.Default_Port_Path = this.Project_Path[Project_name]["PORT_Setting"];
                        this.CurrentProject = Project_name;
                        textBox1.AppendText("Change Project = " + this.CurrentProject + "\r\n");

                        Write_Default(this.Default_CM_Path, this.Default_Port_Path);

                        if (this.Load_Initial_CM_Complete)
                        {
                            StringBuilder Set_alarm = new StringBuilder();
                            Set_alarm.AppendFormat("CHANGE Project as {0}", Project_name);
                            Set_alarm.AppendFormat("\n ");
                            Set_alarm.AppendFormat("\n ------ PLEASE LOAD CM SHEET AGAIN ------");
                            Set_alarm.AppendFormat("\n ");
                            Set_alarm.AppendFormat("\n CM sheet setting : {0}", this.Default_CM_Path);
                            Set_alarm.AppendFormat("\n Port Config info : {0}", this.Default_Port_Path);
                            ClsMsgBox.Show("SET Convert option", Set_alarm.ToString());
                        }
                        else
                        {
                            StringBuilder Set_alarm = new StringBuilder();
                            Set_alarm.AppendFormat("SET Project as {0}", Project_name);
                            Set_alarm.AppendFormat("\n ");
                            Set_alarm.AppendFormat("\n CM sheet setting : {0}", this.Default_CM_Path);
                            Set_alarm.AppendFormat("\n Port Config info : {0}", this.Default_Port_Path);
                            ClsMsgBox.Show("SET Convert option", Set_alarm.ToString());
                        }
                    }
                    else
                    {
                        StringBuilder Set_alarm = new StringBuilder();
                        Set_alarm.AppendFormat("FAIL to SET Project as {0}", Project_name);
                        Set_alarm.AppendFormat("\n ");
                        if (!CMsetting_exist) Set_alarm.AppendFormat("\n (NOT EXIST) CM sheet setting : {0}", this.Default_CM_Path);
                        if (!Portsetting_exist) Set_alarm.AppendFormat("\n (NOT EXIST) Port Config info : {0}", this.Default_Port_Path);
                        ClsMsgBox.Show("ERROR to SET Convert option", Set_alarm.ToString());
                    }

                    

                    break;
                }
                
            }
        }

        private void Write_Default(string CM_setting_Path, string Port_Setting_Path)
        {
            string PathDefault = @"C:\ProgramData\FlexTest\GENTLE_BREED\GentleBreed_default.ini";

            if(File.Exists(PathDefault))
            {
                using (FileStream fs = new FileStream(PathDefault, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    string newContents = "";

                    StreamReader streamReader = new StreamReader(fs);
                    string currentContents = streamReader.ReadToEnd();

                    currentContents = currentContents.Replace('\r', ' ');
                    string[] Lines = currentContents.Split('\n');

                    for (int i = 0; i < Lines.Count(); i++)
                    {
                        if(Lines[i].Contains("DEFAULT_PATH_CM"))
                        {
                            Lines[i] = "DEFAULT_PATH_CM = " + CM_setting_Path;
                        }
                        if(Lines[i].Contains("DEFAULT_PATH_PORT"))
                        {
                            Lines[i] = "DEFAULT_PATH_PORT = " + Port_Setting_Path;
                        }
                        newContents = newContents + Lines[i] + "\r\n";
                    }

                    fs.SetLength(0);

                    StreamWriter writer = new StreamWriter(fs);
                    writer.Write(newContents);
                    writer.Close();
                    streamReader.Close();
                }
            }
        }

        private Dictionary<string, Dictionary<string, string>> Get_ProjectMenu()
        {
            string PathDefault = @"C:\ProgramData\FlexTest\GENTLE_BREED\GentleBreed_default.ini";
            Dictionary<string, Dictionary<string, string>> Project_Path = new Dictionary<string, Dictionary<string, string>>();

            try
            {
                using (StreamReader sr = new StreamReader(PathDefault))
                {
                    string Line;
                    string Key;
                    string Val;
                    string Project_Key = "TBD";
                    
                    Dictionary<string, string> Temp_Project_Paths = new Dictionary<string, string>();

                    while ((Line = sr.ReadLine()) != null)
                    {
                        bool Valid_Line = (Line.Contains('=') && !IsComment(Line));
                        bool Valied_Project = (Line.Contains('[') && Line.Contains(']') && !IsComment(Line));
                        bool IsValid = Valid_Line || Valied_Project;
                        if (!IsValid) continue;

                        if (Valied_Project)
                        {
                            string temp_Line = Line;
                            if (temp_Line.Contains("END"))
                            {
                                Dictionary<string, string> Project_Paths = new Dictionary<string, string>();

                                foreach (var item in Temp_Project_Paths.Keys)
                                {
                                    Project_Paths.Add(item, Temp_Project_Paths[item]);
                                }

                                Project_Path.Add(Project_Key, Project_Paths);

                                Project_Key = "";
                                Temp_Project_Paths.Clear();
                            }
                            else
                            {
                                this.RemoveComment(ref temp_Line);
                                temp_Line = temp_Line.Replace('[', ' ');
                                temp_Line = temp_Line.Replace(']', ' ');
                                temp_Line = temp_Line.Trim();
                                Project_Key = temp_Line;
                            }
                        }
                        else if (Valid_Line)
                        {
                            string[] Substr = Line.Split('=');
                            Key = Substr[0].Trim();
                            Val = Substr[1].Trim();
                            this.RemoveComment(ref Val);

                            if (Key.ToUpper().Contains("DEFAULT_PATH_CM"))
                            {
                                this.Default_CM_Path = Val;
                                //this.CurrentProject = Project_Key;
                            }
                            else if (Key.ToUpper().Contains("DEFAULT_PATH_PORT"))
                            {
                                this.Default_Port_Path = Val;
                                //this.CurrentProject = Project_Key;
                            }
                            else
                            {
                                if (Key.ToUpper().Contains("CM"))
                                {
                                    Key = "CM_Setting";
                                    Temp_Project_Paths.Add(Key, Val);
                                }
                                if (Key.ToUpper().Contains("PORT"))
                                {
                                    Key = "PORT_Setting";
                                    Temp_Project_Paths.Add(Key, Val);
                                }
                            }
                        }
                    }                   
                }
            }
            catch (Exception)
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open default setting <GentleBreed_default.ini>");
                ErrMsg.AppendFormat("\nPlease check file exist in ");
                ErrMsg.AppendFormat("\n<C:\\ProgramData\\FlexTest\\GENTLE_BREED\\GentleBreed_default.ini>");
                ClsMsgBox.Show("Error on file loading in initialization", ErrMsg.ToString());
                Environment.Exit(0);
                throw;
            }

            
            return Project_Path;

        }

        private void button1_Click(object sender, EventArgs e) //Spara Test planner
        {
            Test_Planner.Testsetting TestSetup = new Test_Planner.Testsetting();
            TestSetup.ShowForm(); //>>Globals.Spara_TestCon

            if(Globals.SPara_Plan_Generate)
            {
                Build_Spara Pilot_plan = new Build_Spara();
                Pilot_plan.GeneratePlan(Globals.Spara_TestCon);
                Globals.SPara_Plan_Generate = false;
            }

        }

        private void TxTest_button_Click(object sender, EventArgs e) //Tx Test planner
        {
            Test_Planner.TxPlan_builder TX_testPlanner = new Test_Planner.TxPlan_builder();
            TX_testPlanner.ShowForm(); //>>Globals.Spara_TestCon

        }

        public void Init_ProgressBar(object sender, EventArgs e, int count)
        {
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = count;
            progressBar1.Step = (int)((progressBar1.Maximum - progressBar1.Minimum)/count);
            progressBar1.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
        }

        public void Init_ProgressBar(int count)
        {
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = count;
            progressBar1.Step = (int)((progressBar1.Maximum - progressBar1.Minimum) / count);
            progressBar1.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
        }

        public void Progress_perform(string text)
        {
            textBox1.AppendText("Write to summary =" + text + " ..done..\r\n");
            progressBar1.PerformStep();
        }

        public void Progress_perform_Spara(string text)
        {
            textBox1.AppendText(text + "\r\n");
            progressBar1.PerformStep();
        }

        private int Check_CM_ShowData(List<int> CM_DataList, int index)
        {
            if(CM_DataList.Count != 0)
            {
                return CM_DataList[index];
            }
            else
            {
                return 0;
            }
        }
        private string Check_CM_ShowData(List<string> CM_DataList, int index)
        {
            if (CM_DataList.Count != 0)
            {
                return CM_DataList[index];
            }
            else
            {
                return "";
            }
        }
        private string SEL_Check_CM_ShowData(string TXorRX, List<string> TX_DataList, List<string> RX_DataList, int index)
        {
            if (TXorRX.Trim().ToUpper().Contains("_RX") && RX_DataList.Count != 0)
            {
                return RX_DataList[index];
            }
            else if(TXorRX.Trim().ToUpper().Contains("TX") && TX_DataList.Count != 0)
            {
                return TX_DataList[index];
            }
            else
            {
                return "";
            }
        }
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(Globals.LoadCM_completed && !Globals.Kill_CM_Sheet) CM_Sheet.Quit();
            Environment.Exit(0);
        }

        private void setupConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Process.Start(Globals.INI_Dir);
            Process.Start("wordpad",Globals.CMsheet_INI_Dir);
        }

        private void portConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("wordpad", Globals.PortInfo_INI_Dir);
        }

        private void loadCMWithDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string PFN = "";
            string PathDefault = "C:\\ProgramData\\FlexTest\\GENTLE_BREED";
            string INI_PFN = "";

            Set_GlobalPath(this.Default_CM_Path, this.Default_Port_Path);

            //Initialize global variable routine for repeatable execution.
            Globals.DUT_CM.Clear();
            Globals.IniFile.clear();
            Globals.IniFile.LoadINI(this.Default_CM_Path);
            Globals.LoadCM_completed = false;
            Globals.Kill_CM_Sheet = false;
            this.Spara_button.Enabled = false;

            Globals.Spara_TestDic.Clear();
            Globals.TX_TestDic.Clear();
            Globals.RX_TestDic.Clear();
            Globals.DC_TestDic.Clear();
            Globals.NOISE_TestDic.Clear();
            Globals.Expaned_Spara_Seq.Clear();

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
                    PFN = OpenDialogEntity.FileName;
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
                ErrMsg.AppendFormat("\nPlease check opend file or file name or location");
                ClsMsgBox.Show("Error on file Path loading from dialog window", ErrMsg.ToString());
                Environment.Exit(0);
            }

            try
            {
                int BandIndex = Globals.IniFile.Band_Sheet_Name.Count * 2;
                this.Init_ProgressBar(sender, e, BandIndex);

                CM_Sheet = new Excel_File(false, PFN);
                //CM_Sheet.Show(false);
                int Proc_ID = CM_Sheet.ProcID;

                List<string> Defined_Bands = Globals.IniFile.Band_Sheet_Name;


                //Build Condition Sheet per CM header in Band
                foreach (string current_sheet in Defined_Bands)
                {
                    textBox1.AppendText("Loading sheet of " + current_sheet + " from CM spec..\r\n");
                    progressBar1.PerformStep();
                    Globals.DUT_CM.Add(current_sheet, CM_Sheet.ExcelUpdate_LoadSheetFromCM(CM_Sheet, current_sheet, Globals.IniFile));
                }
                progressBar1.PerformStep();

                Globals.Kill_CM_Sheet = true;
                this.Load_Initial_CM_Complete = true;
                CM_Sheet.Quit(); //Dispose CM sheet after Memory loading 
                //--------------------------------------------------------------------------------------------------------------------------
                //ReBuild Null Condition test name and order key index
                foreach (string current_sheet in Defined_Bands) //re-arrange condition value of each CM sheet
                {
                    if (current_sheet == "Band 11_Rx") Proc_ID = Proc_ID;

                    textBox1.AppendText("Re-Align contents of " + current_sheet + " from CM spec..\r\n");
                    TestCon Get_condition = new TestCon();
                    CMstructure sortMemory = new CMstructure();
                    progressBar1.PerformStep();
                    if (current_sheet == "1.6GHz_Tx") { string stop_here = ""; }

                    sortMemory.Revise_MemoryCM_Full(current_sheet, ref Globals.DUT_CM); //add Test name to null item, and key index
                }
                //--------------------------------------------------------------------------------------------------------------------------

                Globals.LoadCM_completed = true;

                //Globals.Spara_TestDic;
                //Globals.TX_TestDic;
                //Globals.RX_TestDic;
                //Globals.NOISE_TestDic;
                //Globals.DC_TestDic;

                //Globals.DUT_CM
                List<string> band_1 = Globals.IniFile.Band_Sheet_Name;


                //if (checkBox1.Checked)
                if (true)
                {
                    DataTable dt = new DataTable();
                    dt.Columns.AddRange(new DataColumn[]
                    {
                    new DataColumn("Sheet"),
                    new DataColumn("Index"),
                    new DataColumn("Spec ID"),
                    new DataColumn("Test Name"),
                    new DataColumn("Band"),
                    new DataColumn("Param"),
                    new DataColumn("Direction"),
                    new DataColumn("PA MODE"),
                    new DataColumn("LNA MODE"),
                    new DataColumn("Input Port"),
                    new DataColumn("Output Port"),
                    new DataColumn("Start_Freq"),
                    new DataColumn("Stop_Freq"),
                    new DataColumn("Temperature"),
                    new DataColumn("ANT VSWR"),
                    new DataColumn("Limit_L"),
                    new DataColumn("Limit_Typ"),
                    new DataColumn("Limit_U"),
                    new DataColumn("Sample#1_Min"),
                    new DataColumn("Sample#1_Max"),
                    new DataColumn("Sample#2_Min"),
                    new DataColumn("Sample#2_Max"),
                    new DataColumn("Sample#3_Min"),
                    new DataColumn("Sample#3_Max"),
                    new DataColumn("Condition_worst")
                    });

                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 80; //Sheet
                    dataGridView1.Columns[1].Width = 50; //Row_Index
                    dataGridView1.Columns[2].Width = 100; //Spec_ID
                    dataGridView1.Columns[3].Width = 150; //Test Name
                    dataGridView1.Columns[4].Width = 60; //Band
                    dataGridView1.Columns[5].Width = 60; //Param
                    dataGridView1.Columns[6].Width = 60; //Direction
                    dataGridView1.Columns[7].Width = 60; //PA MODE
                    dataGridView1.Columns[8].Width = 60; //LNA MODE
                    dataGridView1.Columns[9].Width = 100; //Input Port
                    dataGridView1.Columns[10].Width = 100; //Output Port
                    dataGridView1.Columns[11].Width = 70; //Start_Freq
                    dataGridView1.Columns[12].Width = 70; //Stop_Freq
                    dataGridView1.Columns[13].Width = 60; //Temperature
                    dataGridView1.Columns[14].Width = 60; //ANT VSWR
                    dataGridView1.Columns[15].Width = 60; //Limit_L
                    dataGridView1.Columns[16].Width = 60; //Limit_Typ
                    dataGridView1.Columns[17].Width = 60; //Limit_U
                    dataGridView1.Columns[18].Width = 80; //Sample#1_Min_value
                    dataGridView1.Columns[19].Width = 80; //Sample#1_Max_value
                    dataGridView1.Columns[20].Width = 80; //Sample#2_Min_value
                    dataGridView1.Columns[21].Width = 80; //Sample#2_Max_value
                    dataGridView1.Columns[22].Width = 80; //Sample#3_Min_value
                    dataGridView1.Columns[23].Width = 80; //Sample#3_Max_value
                    dataGridView1.Columns[24].Width = 100; //Testcondition

                    foreach (var BandDic_Keys in Globals.DUT_CM.Keys)
                    {
                        for (int i = 0; i < Globals.DUT_CM[BandDic_Keys].Test_SpecID.Count(); i++)
                        {
                            dt.Rows.Add(BandDic_Keys,
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Key_Index, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Test_SpecID, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Test_Name, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Band, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Parameter, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Direction, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].PA_MODE, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].LNA_Gain_Mode, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Input_Port, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Output_Port, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Start_Freq, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Stop_Freq, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Temperature, i),
                            SEL_Check_CM_ShowData(BandDic_Keys, Globals.DUT_CM[BandDic_Keys].ANTout_VSWR, Globals.DUT_CM[BandDic_Keys].ANTIn_VSWR, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Test_Limit_L, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Test_Limit_Typ, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Test_Limit_U, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_1_min, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_1_max, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_2_min, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_2_max, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_3_min, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Sample_3_max, i),
                            Check_CM_ShowData(Globals.DUT_CM[BandDic_Keys].Worst_Condition_text, i)
                            );
                        }
                    }
                }

                textBox1.AppendText("---------- Done Loading CM sheet ----------\r\n");

                //CM_Load_Timer.Stop();
                progressBar1.PerformStep();
                Globals.LoadCM_completed = true;
                this.Spara_button.Enabled = true;
                this.Table_Make.Enabled = true;
                this.TxTest_button.Enabled = true;
            }
            catch (Exception err)
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:Open CM Sheet Exception = {0}", err);
                ErrMsg.AppendFormat("\nPlease check opened file or file name or location");
                ClsMsgBox.Show("Error on Excel file loading", ErrMsg.ToString());
                CM_Sheet.Quit(); //kill without save
                Environment.Exit(0);
            }
        }

        private void closeWindowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Globals.LoadCM_completed && !Globals.Kill_CM_Sheet) CM_Sheet.Quit();
            Environment.Exit(0);
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
        private void sparaPPTPlannerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        
        private void Table_Make_Click(object sender, EventArgs e)
        {
            CMstructure sortMemory = new CMstructure();
            List<string> Defined_Bands = Globals.IniFile.Band_Sheet_Name;


            //Globals.Spara_TestDic;
            //Globals.TX_TestDic;
            //Globals.RX_TestDic;
            //Globals.NOISE_TestDic;
            //Globals.DC_TestDic;

            this.Init_ProgressBar(sender, e, 171); //report step

            //List<Table_data> Data_per_ID = new List<Table_data>();

            string PFN = Globals.Path_Default + "Summary_Table.xlsx";
            Excel_File TableFile_Excel = new Excel_File(true, PFN);

            List<string> Exception_band_keyword = new List<string>();
            Exception_band_keyword.Add("ASM");

            TableFile_Excel.App.ScreenUpdating = false;
            //TableFile_Excel.App.ScreenUpdating = true;

            if (!TableFile_Excel.LoadError)
            {
                //TableFile_Excel.Worksheet_DefaultName("summary"); //only use when create new file (not from format copy)
                int Row_index = 1;

                if (true)
                {

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["Input_RL"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["Input_RL"], "TX", "APT", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["CW_P2dB"], "TX", "APT", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["CW_P2dB"], "TX", "APT", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["CW_P2dB"], "TX", "APT", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_ET"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_ET"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_ET"], "TX", "ET", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_APT"], "TX", "APT", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_APT"], "TX", "APT", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LowGain_APT"], "TX", "APT", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Gain_Slope"], "TX", "All", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["Gain_Ripple"], "TX", "All", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["TX_OOB_Gain"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["Current_ET"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LowCurrent_APT"], "TX", "APT", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LowCurrent_APT"], "TX", "APT", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_ACP"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_ACP"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_ACP"], "TX", "ET", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE_ACP"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE_ACP"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE_ACP"], "TX", "ET", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_ACP"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_ACP"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_ACP"], "TX", "ET", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_EVM"], "TX", "All", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_EVM"], "TX", "All", "OT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_EVM"], "TX", "APT", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NR_EVM"], "TX", "APT", "OT", "VSWR", Exception_band_keyword));


                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE_EVM"], "TX", "All", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE_EVM"], "TX", "All", "OT", "VSWR", Exception_band_keyword));


                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_EVM"], "TX", "All", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_EVM"], "TX", "All", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_EVM"], "TX", "All", "RT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["LTE/NR_SEM"], "TX", "All", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["WCDMA_SEM"], "TX", "All", "RT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["HAR_2"], "TX", "ET", "OT", "VSWR", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["HAR_3"], "TX", "ET", "OT", "VSWR", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["HAR_4"], "TX", "ET", "OT", "VSWR", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["HAR_5"], "TX", "ET", "OT", "VSWR", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["ANT_NOISE"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NS_05"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NS_03"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NS_04"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["NS_21"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["FCC_Emission"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["KOR_SPE_LTE/NR"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["VER_SPE_LTE/NR"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["SPE_LTE/NR"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["SPE_WCDMA"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    //Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["SPE_TDSCDMA"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["RXBN"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["RXBN_ASM_MIMO"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["RXBN_ASM_DRX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    //Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["RXBN_ASM_LMB"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["TXL"], "TX", "ET", "RT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["TXL"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.NOISE_TestDic["TXL_DPX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.TX_TestDic["TXL_ASM"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:TX, RX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:TX, InAct_RX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:TX, ASM"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:RX, InAct_RX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:InAct_RX, InAct_RX"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:ANT, ANT"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.Spara_TestDic["ISO:ANT, InAct_ANT"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));

                    Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "TX_summary", ref Row_index, Globals.DC_TestDic["LEAK_RF_DC"], "TX", "ET", "OT", "50ohm", Exception_band_keyword));
                }
                TableFile_Excel.Select_Sheet("RX_summary");
                Row_index = 1;

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Input_RL"], "RX", "All_GainMode", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Input_RL"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Output_RL"], "RX", "All_GainMode", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Output_RL"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Phase_Delta"], "RX", "All_GainMode", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["K_factor"], "RX", "All_GainMode", "OT", "VSWR", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["MU_factor"], "RX", "All_GainMode", "OT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G0"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G0"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G1"], "RX", "G1", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G1"], "RX", "G1", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G2"], "RX", "G2", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G2"], "RX", "G2", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G3"], "RX", "G3", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G3"], "RX", "G3", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G4"], "RX", "G4", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G4"], "RX", "G4", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G5"], "RX", "G5", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G5"], "RX", "G5", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G6"], "RX", "G6", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["Current_G6"], "RX", "G6", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G0"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G0"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G0"], "RX", "G0", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G1"], "RX", "G1", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G1"], "RX", "G1", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G1"], "RX", "G1", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G2"], "RX", "G2", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G2"], "RX", "G2", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G2"], "RX", "G2", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G3"], "RX", "G3", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G3"], "RX", "G3", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G3"], "RX", "G3", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G4"], "RX", "G4", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G4"], "RX", "G4", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G4"], "RX", "G4", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G5"], "RX", "G5", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G5"], "RX", "G5", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G5"], "RX", "G5", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G6"], "RX", "G6", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G6"], "RX", "G6", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_G6"], "RX", "G6", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G0"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G1"], "RX", "G1", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G2"], "RX", "G2", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G3"], "RX", "G3", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G4"], "RX", "G4", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G5"], "RX", "G5", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["NF_CA_G6"], "RX", "G6", "RT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.NOISE_TestDic["NFR_MIPI"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.NOISE_TestDic["NFR"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.NOISE_TestDic["NFR"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.NOISE_TestDic["NFR"], "RX", "G0", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G0"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G0"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G0"], "RX", "G0", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G1"], "RX", "G1", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G1"], "RX", "G1", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G1"], "RX", "G1", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G2"], "RX", "G2", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G2"], "RX", "G2", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G2"], "RX", "G2", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G3"], "RX", "G3", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G3"], "RX", "G3", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G3"], "RX", "G3", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G4"], "RX", "G4", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G4"], "RX", "G4", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G4"], "RX", "G4", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G5"], "RX", "G5", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G5"], "RX", "G5", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G5"], "RX", "G5", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G6"], "RX", "G6", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G6"], "RX", "G6", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_G6"], "RX", "G6", "RT", "VSWR", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G0"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G1"], "RX", "G1", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G2"], "RX", "G2", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G3"], "RX", "G3", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G4"], "RX", "G4", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G5"], "RX", "G5", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_Gain_CA_G6"], "RX", "G6", "RT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Gain_Ripple"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G0"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G1"], "RX", "G1", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G2"], "RX", "G2", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G3"], "RX", "G3", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G4"], "RX", "G4", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G5"], "RX", "G5", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_P1dB_G6"], "RX", "G6", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G0"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G1"], "RX", "G1", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G2"], "RX", "G2", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G3"], "RX", "G3", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G4"], "RX", "G4", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G5"], "RX", "G5", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_IIP3_G6"], "RX", "G6", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.RX_TestDic["RX_EVM"], "RX", "All_GainMode", "OT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["Group_Delay"], "RX", "All_GainMode", "OT", "50ohm", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["REV_ISO:RX, ANT"], "RX", "G0-G5", "RT", "TBD", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["REV_ISO:RX, ANT"], "RX", "G6", "OT", "TBD", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["ISO:ANT, ANT"], "RX", "All_GainMode", "OT", "TBD", Exception_band_keyword));

                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["RX_OOB_Gain"], "RX", "G0", "RT", "50ohm", Exception_band_keyword));
                Progress_perform(sortMemory.Summary_Table(TableFile_Excel, "RX_summary", ref Row_index, Globals.Spara_TestDic["RX_OOB_Gain"], "RX", "G0", "OT", "50ohm", Exception_band_keyword));

                TableFile_Excel.App.ScreenUpdating = true;

                StringBuilder Msg = new StringBuilder();
                Msg.AppendFormat("TX RX summarize done");
                Msg.AppendFormat("\nPlease check opend file and copy to other location");
                ClsMsgBox.Show("Summarize done", Msg.ToString());

                TableFile_Excel.Release_Excel_Resource();
            }
        }

        private void BTN_Build_PPT_Click(object sender, EventArgs e)
        {

            Build_PPT_wfm BuildPPT_UI = new Build_PPT_wfm();
            BuildPPT_UI.Show();

        }

        private void BTN_GrabSnp_Click(object sender, EventArgs e)
        {
            GrabSnPs Grab_snpfiles = new GrabSnPs();
            Grab_snpfiles.Show();
        }

        public int temp_Excel_ID = 0;

        private void BTN_INSERT_SPARA_Click(object sender, EventArgs e)
        {
            InsertData_Spara Data_Spara_tool = new InsertData_Spara();
            Data_Spara_tool.Show();
        }

        private void BTN_Create_Worst_Click(object sender, EventArgs e)
        {
            string WorstTable = "";
            string PathDefault = "C:\\ProgramData\\FlexTest\\GENTLE_BREED";

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
                    WorstTable = OpenDialogEntity.FileName;
                }
                else
                {
                    return;
                }
            }
            catch
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error:During open worst list file");
                ErrMsg.AppendFormat("\nPlease check open file format (Worst test condition number list)");
                MessageBox.Show("Error on file loading", ErrMsg.ToString());
                //Environment.Exit(0);
            }

            try
            {
                Excel_File WorstTableExcel = new Excel_File(false, WorstTable);
                this.temp_Excel_ID = WorstTableExcel.ProcID;
                WorstTableExcel.Show(true);
                WorstTableExcel.App.ScreenUpdating = true;

                int initRow = 20;
                int initCol = 20;

                string[,] Rowdata = WorstTableExcel.ReadData_From_WorkSheet("Worst_Con", 1, initRow, 1, initCol);

                for (int i = 0; i < initRow; i++)
                {
                    for (int j = 0; j < initCol; j++)
                    {
                        if(Rowdata[i,j].ToUpper().Contains("WORST_CASES"))
                        {
                            initRow = i + 2;
                            initCol = j + 1;
                            break;
                        }
                    }

                }

                string[,] Full_data = WorstTableExcel.ReadData_From_WorkSheet("Worst_Con", initRow, 100000, initCol, initCol);

                List<string> CaseNumber_list = new List<string>();
                int Null_Cell = 0;

                for (int k = 0; k < 100000; k++)
                {
                    if (Full_data[k, 0] == "") Null_Cell++;
                    CaseNumber_list.Add(Full_data[k, 0]);
                    if (Null_Cell > 10) break;
                }

                Kill_Process(this.temp_Excel_ID); //Dispose Excel file after Memory loading 

                string Acting_Plan = "";

                System.Diagnostics.Process[] AfterExcelProcess;
                AfterExcelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                if(AfterExcelProcess.Length!=0)
                {
                    for (int i = 0; i < AfterExcelProcess.Length; i++)
                    {
                        if(AfterExcelProcess[i].Id!= this.temp_Excel_ID) Kill_Process(AfterExcelProcess[i].Id);
                    }
                }
                else
                {
                    string debug = "null";
                }

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
                        Acting_Plan = OpenDialogEntity.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
                catch
                {
                    StringBuilder ErrMsg = new StringBuilder();
                    ErrMsg.AppendFormat("Error:During open worst list file");
                    ErrMsg.AppendFormat("\nPlease check open file format (Worst test condition number list)");
                    MessageBox.Show("Error on file loading", ErrMsg.ToString());
                    //Environment.Exit(0);
                }

                Excel_File TableFile_Excel = new Excel_File(false, Acting_Plan);
                int index_Row = 1;

                string[,] Full_Condition = TableFile_Excel.ReadData_From_WorkSheet("Condition_FBAR", 3, 100000, 1, 4);
                int Last_DC_index = 0;

                for (int i = 0; i < 100000; i++)
                {
                    string DC_FBAR = Full_Condition[i, 3];
                    string TestNum = Full_Condition[i, 2].Trim();
                    string Enable = Full_Condition[i, 1];
                    string Range_Index = Full_Condition[i, 0];

                    if (DC_FBAR.Contains("DC"))
                    {
                        Last_DC_index = i;
                    }
                    else if(DC_FBAR.Contains("FBAR"))
                    {
                        if (CaseNumber_list.Contains(TestNum))
                        {
                            Full_Condition[i, 1] = "x";
                            Full_Condition[Last_DC_index, 1] = "x";
                        }                      
                    }

                    if (Range_Index.ToUpper().Contains("#STOP"))
                    {
                        TableFile_Excel.WriteData_ToArray("Condition_FBAR", 3, 1, Full_Condition, false);
                        break;
                    }
                }
                




            }
            catch
            {

            }
        }
        private void Kill_Process(int P_ID)
        {
            int ProcID = P_ID;
            Process Proc = Process.GetProcessById(ProcID);
            Proc.Kill();
        }

    }
}
