using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO; //Directory and file control
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace S_para_planner
{
    public partial class GrabSnPs : Form
    {
        public string SNPdata_loot = "";
        string PathDefault = System.IO.Directory.GetCurrentDirectory();
        public bool Status_OK = false;
        public Dictionary<string, string> SubDir = new Dictionary<string, string>();
        public string Default_Data_path = @"C:\Grap_TEMP_SNP_Data";

        public GrabSnPs()
        {
            InitializeComponent();
            BTN_Grab_SNP.Enabled = false;
        }

        private void BTN_selectPath_Click(object sender, EventArgs e)
        {
            try
            {
                Status_OK = false;
                SubDir = new Dictionary<string, string>();
                listBox1.Items.Clear();

                OpenFileDialog folderBrowser = new OpenFileDialog();
                folderBrowser.ValidateNames = false;
                folderBrowser.CheckFileExists = false;
                folderBrowser.CheckPathExists = true;
                // Always default to Folder Selection.
                folderBrowser.FileName = "데이터가 위치한 상위 폴더에 맞추고 OK";
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = Path.GetDirectoryName(folderBrowser.FileName);
                    SNPdata_loot = folderPath;
                    TBox_FilePath.Text = folderPath;
                    TBox_FilePath.BackColor = Color.GreenYellow;
                    BTN_selectPath.BackColor = Color.YellowGreen;
                    BTN_selectPath.Text = "경로 지정됨";
                }

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
                            break;
                        }
                    }

                    if (IsSNP_Directory)
                    {
                        string Directory_name = new DirectoryInfo(each_Path).Name;
                        SubDir.Add(Directory_name, each_Path);
                        listBox1.Items.Add(Directory_name + " : " + each_Path);
                        Status_OK = true;
                        BTN_Grab_SNP.Enabled = true;
                        BTN_Grab_SNP.BackColor = Color.LightGoldenrodYellow;
                    }
                }
            }
            catch
            {

            }
        }

        private void BTN_Grab_SNP_Click(object sender, EventArgs e)
        {
            if (Text_TestNum.Text.Trim() != "" && Status_OK)
            {
                string TestNum_Text = Text_TestNum.Text.Trim();
                int TestNum = Convert.ToInt32(TestNum_Text);

                if (Check_Data_erase.Checked)
                {
                    bool ExistDataFolder = System.IO.Directory.Exists(Default_Data_path);
                    if (ExistDataFolder)
                    {
                        string[] OldeFiles = Directory.GetFiles(Default_Data_path);
                        foreach (string each_file in OldeFiles)
                        {
                            System.IO.File.Delete(each_file);
                        }
                    }
                }

                foreach (string Paths in SubDir.Values)
                {
                    SortedDictionary<int, string> SNP_File_found = new SortedDictionary<int, string>();

                    string[] Files = Directory.GetFiles(Paths);
                    foreach (string each_file in Files)
                    {
                        string fileName = new DirectoryInfo(each_file).Name;
                        string[] fileName_arry = fileName.Split('_');
                        int Current_Num = 0;

                        for (int i = 0; i < fileName_arry.Length; i++)
                        {
                            if (int.TryParse(fileName_arry[i], out Current_Num))
                            {
                                break;
                            }
                        }

                        if (Current_Num != 0)
                        {
                            if (!SNP_File_found.ContainsKey(Current_Num))
                            {
                                SNP_File_found.Add(Current_Num, each_file);
                            }
                        }
                    }

                    if (SNP_File_found.Count != 0)
                    {
                        int Last_Num = 0;

                        foreach (int SNP_TestNum in SNP_File_found.Keys)
                        {
                            if (SNP_TestNum < TestNum)
                            {
                                Last_Num = SNP_TestNum;
                            }
                            else if (SNP_TestNum == TestNum)
                            {
                                if (SNP_File_found.ContainsKey(SNP_TestNum))
                                {
                                    Send_SNP(SNP_File_found[SNP_TestNum]);
                                }
                                break;
                            }
                            else if (SNP_TestNum > TestNum)
                            {
                                if (SNP_File_found.ContainsKey(Last_Num))
                                {
                                    Send_SNP(SNP_File_found[Last_Num]);
                                }
                                break;
                            }

                        }
                    }
                }

                bool IsExist = System.IO.Directory.Exists(Default_Data_path);
                if (IsExist)
                {
                    System.Diagnostics.Process.Start("explorer.exe", Default_Data_path);
                }
            }
        }
        private void Send_SNP(string Target_File_path)
        {
            bool IsExist = System.IO.Directory.Exists(Default_Data_path);
            if (!IsExist)
            {
                System.IO.Directory.CreateDirectory(Default_Data_path);
            }

            string fileName = new DirectoryInfo(Target_File_path).Name;
            string Destination_File = Default_Data_path + "\\" + fileName;
            System.IO.File.Copy(Target_File_path, Destination_File, true);

        }
    }
}
