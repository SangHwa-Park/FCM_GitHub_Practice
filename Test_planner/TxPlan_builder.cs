using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using Excel_Base;
using Test_Planner;


namespace Test_Planner
{
    public partial class TxPlan_builder : Form
    {
        public CMstructure sortMemory = new CMstructure();
        public int Excel_Proc_ID = 0;

        public TxPlan_builder()
        {
            InitializeComponent();
        }
        public void ShowForm()
        {
            this.Interface_Setting();
            //this.Listview_Header_initialize();
            //this.Set_treeview_with_Expanded_Seq(Globals.Expaned_Spara_Seq);

            this.ShowDialog();
        }

        public void Interface_Setting()
        {
            List<string> Defined_Bands = Globals.IniFile.Band_Sheet_Name;
            Dictionary<string, List<string>> TX_RF_List = Globals.TX_TestDic;


            //Globals.DUT_CM
            //Globals.IniFile

            Dictionary<string, List<TestCon>> LTE_ACP = new Dictionary<string, List<TestCon>>();
            Dictionary<string, List<TestCon>> NR_ACP = new Dictionary<string, List<TestCon>>();
            Dictionary<string, List<TestCon>> WCDMA_ACP = new Dictionary<string, List<TestCon>>();
            Dictionary<string, List<TestCon>> CDMA_ACP = new Dictionary<string, List<TestCon>>();
            Dictionary<string, List<TestCon>> TDSCDMA_ACP = new Dictionary<string, List<TestCon>>();

            LTE_ACP = GetTestcon("LTE_ACP", TX_RF_List);
            NR_ACP = GetTestcon("NR_ACP", TX_RF_List);
            WCDMA_ACP = GetTestcon("WCDMA_ACP", TX_RF_List);

            


            Dictionary<string, List<TestCon>> RXBN_ref = new Dictionary<string, List<TestCon>>();
            RXBN_ref = GetTestcon("RXBN", Globals.NOISE_TestDic);

        }
        private void BTN_LoadTxSet_Click(object sender, EventArgs e)
        {
            string APT_plan_setting = "";
            string PathDefault = "C:\\ProgramData\\FlexTest\\GENTLE_BREED";
            BTN_LoadTxSet.BackColor = Color.LightGray;
            Show_Path_TxSet.BackColor = Color.LightGray;

            #region open_dialog(open_excel)
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
                    APT_plan_setting = OpenDialogEntity.FileName;
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
                ErrMsg.AppendFormat("\nPlease check open file format (Plan setting file)");
                MessageBox.Show("Error on file loading", ErrMsg.ToString());
                //Environment.Exit(0);
            }
            #endregion

            Excel_File PPT_source = new Excel_File(false, APT_plan_setting);
            this.Excel_Proc_ID = PPT_source.ProcID;
            PPT_source.Show(true);
            PPT_source.App.ScreenUpdating = true;

            int Start_row = 1;
            int Stop_row = 100000;
            int Start_col = 1; //Header list[2] = Enable "x"
            int Stop_col = 200;

            //List<string> Header_list = PPT_source.Find_Header("Support_Band", "Enable", "v", ref Start_row, ref Stop_row, ref Start_col, ref Stop_col);
            string[,] Support_Band = PPT_source.ReadData_From_WorkSheet("Support_Band", Start_row + 1, Stop_row, 1, Stop_col);
            string[,] Waveform_lists = PPT_source.ReadData_From_WorkSheet("Waveform_List", Start_row + 1, Stop_row, 1, Stop_col);
            string[,] Bias_Table = PPT_source.ReadData_From_WorkSheet("ET_APT_Bias Table", Start_row + 1, Stop_row, 1, Stop_col);
            string[,] Extra_Table = PPT_source.ReadData_From_WorkSheet("General_Setting", Start_row + 1, Stop_row, 1, Stop_col);

            int Index_Test_Enable = 1;
            int Index_Test_Mode = 3;
            /*
            for (int i = 0; i < Header_list.Count; i++)
            {
                if (Header_list[i].Trim().ToUpper().Contains("ENABLE")) Index_Test_Enable = i;
                if (Header_list[i].Trim().ToUpper().Contains("TEST MODE")) Index_Test_Mode = i;

            }
            */

        }

        private Dictionary<string, List<TestCon>> GetTestcon(string Get_Index, Dictionary<string, List<string>> TX_RF_List)
        {
            List<string> Band_ListIndex = new List<string>();

            foreach (string TestID_key in TX_RF_List.Keys)
            {
                if(Get_Index.Trim().ToUpper()==TestID_key.Trim().ToUpper())
                {
                    Band_ListIndex = TX_RF_List[TestID_key];
                    break;
                }
            }

            Dictionary<string, List<TestCon>> ITEM_testcon = new Dictionary<string, List<TestCon>>();

            if (Band_ListIndex.Count!=0)
            {
                foreach (string item in Band_ListIndex)
                {
                    string[] splite_ID = item.Split(',');
                    string Sheet_Band = splite_ID[0].Trim();
                    int Spec_rowIndex = Convert.ToInt32(splite_ID[1].Trim());
                    
                    TestCon TestCon = new TestCon();
                    TestCon = sortMemory.Getcondition_by_Index(Sheet_Band, Spec_rowIndex);

                    if(ITEM_testcon.ContainsKey(Sheet_Band))
                    {
                        ITEM_testcon[Sheet_Band].Add(TestCon);
                    }
                    else
                    {
                        List<TestCon> new_list = new List<TestCon>();
                        new_list.Add(TestCon);
                        ITEM_testcon.Add(Sheet_Band, new_list);
                    }
                }
            }
            
            return ITEM_testcon;
        }
    }
}
