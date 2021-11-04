using System;
using System.Collections.Generic;
using System.Diagnostics; //for system process ID read
using System.Reflection;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing; //it has system font lib
using Excel = Microsoft.Office.Interop.Excel; //for excel control
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using FlexTestLib.MsgBox; //attach Flextest MSG project for reference

namespace Excel_Base
{
    //================================== External Excel Method ==================================//
    public class Excel_HeaderItem
    {
        public int nColumn = 0;
        public bool Show = true;
        public string Name = "";
        public string Unit = "";
        public int? NumDecPlc = 2;
        public string Comment = "";

        public Excel_HeaderItem(int _nColumn, bool _Show, string _Name, string _Unit = "", int? _NumDecPlc = 2, string _Comment = "")
        {
            this.nColumn = _nColumn;
            this.Show = _Show;
            this.Name = _Name;
            this.Unit = _Unit;
            this.NumDecPlc = _NumDecPlc;
            this.Comment = _Comment;
        }
    }

    //================================== Core Excel Method (using Microsoft.Office.Interop.Excel) ==================================//
    public class Excel_File
    {
        public int ProcID;
        public Excel.Application App;
        public Excel.Workbook Workbook;
        public Excel.Worksheet Worksheet;
        public string WorkbookPFN;
        public Excel.Range CelRngXY;
        public bool LoadError = false;

        

        public List<Excel_HeaderItem> SheetTestCon_Header = new List<Excel_HeaderItem>();
        public Excel_File(bool Create_New, string File_Path) //basic class
        {
            string PFN = File_Path;
            bool Error_Code;
            this.AppLoad(Create_New); //Argument = visible : false, Load application = means load Excel program itself 

            if (Create_New)
            {
                //this.Workbook_New(File_Path);
                this.Workbook_NewFormat(File_Path);
                Error_Code = this.LoadError;
            }
            else
            {
                Workbook_Open(PFN, true); //open workbook after Excel program loading
                this.Show(true);
                Error_Code = this.LoadError;
            }
            
        }
        public Excel_File(bool Create_New) //basic class
        {
            this.AppLoad(Create_New); //Argument = visible : false, Load application = means load Excel program itself 
            
            //foreach (Excel.Workbook wb in this.App.Workbooks)
            //{
            //    wb.Close(false);
            //}

            this.App.Workbooks.Add(Type.Missing);
            this.Workbook = this.App.Workbooks.get_Item(1);
            for (int nSheet = 1; nSheet <= this.App.Worksheets.Count; nSheet++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            }
            this.LoadError = false;
        }

        public Excel_File(string File_Path)
        {
            this.LoadError = true;
            File_Path = File_Path + ".xlsx";
            
            bool Error_Code;
            int file_num = 1;
            File_Path = revise_FileName(File_Path, file_num);
            try
            {
                this.AppLoad(false);
                this.App.Workbooks.Add(Type.Missing);
                this.Workbook = this.App.Workbooks.get_Item(1);
                this.Workbook.SaveAs(File_Path);
                this.WorkbookPFN = File_Path;
                this.Workbook = this.App.Workbooks.get_Item(1);
                this.LoadError = false;
            }
            catch
            {
                string Error = "";
                Error = "ERROR during Create New excel : file name error";
                this.LoadError = true;
                ClsMsgBox.Show("ERROR", Error);
            }
            
            Error_Code = this.LoadError;
        }

        public void Create_SheetName(List<string> sheet_list, string active_sheetName)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);

            for (int nSheet = 1; nSheet <= sheet_list.Count; nSheet++)
            {
                try
                {
                    Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                    ws.Name = sheet_list[nSheet - 1];
                }
                catch
                {
                    this.Workbook.Worksheets.Add(Missing.Value, this.Workbook.Worksheets[nSheet - 1], Missing.Value, Missing.Value);
                    Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                    ws.Name = sheet_list[nSheet - 1];
                }
            }

            this.Select_Sheet(active_sheetName);
        }

        private string revise_FileName(string File_Path, int order)
        {
            if (System.IO.File.Exists(File_Path))
            {
                StringBuilder filename = new StringBuilder();
                string[] temp = File_Path.Split('.');

                if (temp[0].EndsWith(")"))
                {
                    temp[0] = temp[0].Substring(0, temp[0].Length - 3);
                }

                filename.AppendFormat("{0}({1}).xlsx", temp[0], order);
                order++;
                File_Path = revise_FileName(filename.ToString(), order);
            }
            else
            {
                return File_Path;
            }

            return File_Path;
        }

        public void Show(bool Visible)
        {
            this.App.Visible = Visible;          
        }
        public void Activate_App(bool Visible)
        {
            this.App.Visible = Visible;
            this.App.ScreenUpdating = Visible;
        }

        private void AppLoad(bool Visible) //Open Blank Excel application
        {
            this.App = new Excel.Application(); //Load program and patch program attribute
            this.App.Visible = Visible;
            this.App.ReferenceStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1; //cell style = fixed cell address
            this.App.DisplayAlerts = false;
            this.ProcID = ClsProcess.GetLastProcID("Excel"); //process ID patched using system library
            this.App.SheetsInNewWorkbook = 3;           
            this.WorkbookPFN = "";
        }

        public void Workbook_NewFormat(string File_Path) //Creat new workbook (not sheet)
        {
            string Error = "";
            try
            {
                string default_format = @"C:\ProgramData\FlexTest\GENTLE_BREED\Summary_Table_format";
                this.LoadError = false;
                string ErrMsg = "";
                bool AllRqdSheetsFound = false;

                try
                {
                    foreach (Excel.Workbook wb in this.App.Workbooks)
                    {
                        wb.Close(false);
                    }

                    this.App.Workbooks.Open(default_format, false);
                    if (System.IO.File.Exists(File_Path))
                    {
                        System.IO.File.Delete(File_Path);
                    }
                    this.Workbook = this.App.Workbooks.get_Item(1);
                    //this.Worksheet = this.App.Worksheets.get_Item("Testcon");
                    //this.Worksheet.Cells[1, 1] = "ID";
                    this.Workbook.SaveAs(File_Path);
                    this.WorkbookPFN = File_Path;
                    this.Workbook = this.App.Workbooks.get_Item(1);

                    for (int nSheet = 1; nSheet <= this.App.Worksheets.Count; nSheet++)
                    {
                        Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                        //if (ws.Name.Trim().ToUpper() == SheetTestCfg_NameRqd.Trim().ToUpper()) SheetTestCfg_nSheetFound = nSheet;                   
                    }

                    AllRqdSheetsFound = true;
                    this.LoadError = false;
                  
                }
                catch
                {
                    Error = "ERROR : Please close Opened summary table first()";
                    this.LoadError = true;
                    ClsMsgBox.Show("ERROR", Error);
                }
            }
            catch (Exception)
            {
                Error = "ERROR during Workbook_New()";
                this.LoadError = true;
                ClsMsgBox.Show("ERROR", Error);
            }
        }

        public void Worksheet_DefaultName(string sheetname)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws_TX = (Excel.Worksheet)this.App.Worksheets["Sheet1"];
            ws_TX.Name = "TX_" + sheetname;
            Excel.Worksheet ws_RX = (Excel.Worksheet)this.App.Worksheets["Sheet2"];
            ws_RX.Name = "RX_" + sheetname;
        }

        public void Release_Excel_Resource()
        {
            //Marshal.ReleaseComObject(this.Worksheet);
            Marshal.ReleaseComObject(this.Workbook);
            Marshal.ReleaseComObject(this.App);
        }

        public void Workbook_Open(string PFN, bool CheckForErrors) //Open exist file
        {
            this.LoadError = false;
            string ErrMsg = "";
            bool AllRqdSheetsFound = false;

            try
            {
                foreach (Excel.Workbook wb in this.App.Workbooks)
                {
                    wb.Close(false);
                }

                this.App.Workbooks.Open(PFN, false); //Excel Lib method

                this.WorkbookPFN = PFN;
                this.Workbook = this.App.Workbooks.get_Item(1);

                for (int nSheet = 1; nSheet <= this.App.Worksheets.Count; nSheet++)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                    //if (ws.Name.Trim().ToUpper() == SheetTestCfg_NameRqd.Trim().ToUpper()) SheetTestCfg_nSheetFound = nSheet;                   
                }

                AllRqdSheetsFound = true;
                this.LoadError = false;
            }
            catch
            {
                ErrMsg = "ERROR during 'WorkBook_Open()'";
                this.LoadError = true;
                ClsMsgBox.Show("ERROR", ErrMsg);
            }
        }

        public List<string> get_ExcelSheet_List()
        {
            List<string> Sheet_Names = new List<string>();

            this.Workbook = this.App.Workbooks.get_Item(1);
            for (int nSheet = 1; nSheet <= this.App.Worksheets.Count; nSheet++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                Sheet_Names.Add(ws.Name);
            }

            return Sheet_Names;
        }

        public bool Find_ExcelSheet(string Sheet_Name, ref string actual_sheet_name)
        {
            bool Is_Exist = false;

            this.Workbook = this.App.Workbooks.get_Item(1);
            for (int nSheet = 1; nSheet <= this.App.Worksheets.Count; nSheet++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
                if(ws.Name.Trim().ToUpper()==Sheet_Name.Trim().ToUpper())
                {
                    actual_sheet_name = ws.Name;
                    Is_Exist = true;
                    return Is_Exist;
                }
            }

            return Is_Exist;
        }

        public Excel.Worksheet getSheet(string nSheet)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            return ws;
        }
        
        public void Select_Sheet(string sheet_name)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[sheet_name];
        }
        public void Delete_Sheet(string sheet_name)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[sheet_name];
            ws.Delete();
        }
        public void Clear_Sheet()
        {
            List<string> default_sheet = new List<string>();
            default_sheet.Add("Sheet1");
            default_sheet.Add("Sheet2");
            default_sheet.Add("Sheet3");

            foreach (string Sheet_name in default_sheet)
            {
                this.Workbook = this.App.Workbooks.get_Item(1);
                Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[Sheet_name];
                ws.Delete();
            }
        }

        public void Add_Sheet(string sheet_name)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = this.Workbook.Worksheets.Add(After: this.Workbook.Sheets[this.Workbook.Sheets.Count]) as Excel.Worksheet;
            ws.Name = sheet_name;
        }

        public void Add_SlideHeader(string sheet_name)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[sheet_name];
            List<string> Subject_Header = new List<string>();
            Subject_Header.Add("Slides");
            Subject_Header.Add("Properties");
            WriteData_1D_Row(sheet_name, 1, 1, Subject_Header);
            
            System.Object Cel1 = ws.Cells[1, 2];  //for PPT slide generation
            System.Object Cel2 = ws.Cells[1, 11];
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.Merge();

            Rformat_Width(ws, 1, 11, 15);
            Rformat_Color(ws, 1, 11, Excel.XlThemeColor.xlThemeColorAccent5, 0);
        }

        public void Merge_Cell(string sheet_name, int row, int row_end, int col, int col_end)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[sheet_name];
            System.Object Cel1 = ws.Cells[row, col];  //for PPT slide generation
            System.Object Cel2 = ws.Cells[row_end, col_end];
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.Merge();
            draw_LineEdge(ws, Cel1, Cel2);
        }

        public void Write_SlideTitle(string sheet, int SlideCnt, ref int Row, string Title, string Slide_Option, bool IsSub)
        {
            List<string> Array_String = new List<string>();

            if(IsSub)
            {
                Array_String.Add("");
                Array_String.Add("Subtitle");
                Array_String.Add(Title);
            }
            else
            {
                Array_String.Add("Slide" + Convert.ToString(SlideCnt) + (Slide_Option.Contains("Option4") ? "\nOption4" : ""));
                Array_String.Add("Title");
                Array_String.Add(Title);
            }

            var ValueArray = new object[1, Array_String.Count];

            for (var col = 0; col <= (Array_String.Count - 1); col++)
            {
                ValueArray[0, col] = Array_String[col];
            }

            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[sheet];
            System.Object Cel1 = ws.Cells[Row, 1];  //for PPT slide generation
            System.Object Cel2 = ws.Cells[Row, 1 + Array_String.Count - 1];
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.Font.Name = "Arial";
            RowRng.Font.Size = 10;
            RowRng.Font.FontStyle = "Bold";
            RowRng.Value2 = ValueArray;
            
            Cformat_Color(ws, Row, 2, Excel.XlThemeColor.xlThemeColorAccent5, 0);
            
            Cel1 = ws.Cells[Row, 1];  
            Cel2 = ws.Cells[Row, 2];
            RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            Cel1 = ws.Cells[Row, 3];  
            Cel2 = ws.Cells[Row, 11];
            RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.Merge();

            string[] temp = Title.Split('/');
            int Length_index = temp[0].Length;
            int Length_total = Title.Length;

            if (IsSub)
            {
                RowRng.Font.Italic = true;
                RowRng.Characters[1, Length_index].Font.Color = Color.Red;
                RowRng.Characters[Length_index + 1, Length_total].Font.Color = Color.BlueViolet;
            }

            Cel1 = ws.Cells[Row, 1];  
            Cel2 = ws.Cells[Row, 11];
            RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);

            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            Row++;
        }

        public void WriteData_Row_with_formatting(string nSheet, int nRow, int nColFirst, List<string> ValueList, int type)
        {
            //Slide plan generation purpose
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = ws.Cells[nRow, nColFirst];
            System.Object Cel2 = ws.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            RowRng.Value2 = ValueArray;
            RowRng.Font.Name = "Arial";
            RowRng.Font.Size = 10;
            draw_LineEdge(ws, Cel1, Cel2);

            if (type == 1) // 0 == no formatting
            {
                //Blue background, white bold text, align to center with different combination
                //col2 = bluc, col 3 = skyblue, col 4 to end = gray
                RowRng.Font.Color = Color.White;
                RowRng.Font.FontStyle = "Bold";
                RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Cel1 = ws.Cells[nRow, nColFirst];
                Cel2 = ws.Cells[nRow, nColFirst];
                RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);

                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0; //blue

                Cel1 = ws.Cells[nRow, nColFirst + 1];
                Cel2 = ws.Cells[nRow, nColFirst + 1];
                RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);

                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0.5; //Light blue

                Cel1 = ws.Cells[nRow, nColFirst + 2];
                Cel2 = ws.Cells[nRow, nColFirst + ValueList.Count - 1];
                RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);

                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0.8;
            }
            else if (type == 2)
            {
                Cel1 = ws.Cells[nRow, nColFirst];
                Cel2 = ws.Cells[nRow, nColFirst];
                RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
                RowRng.Font.Color = Color.White;
                RowRng.Font.FontStyle = "Bold";
                RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0.5; //Light blue

                Cel1 = ws.Cells[nRow, nColFirst + 1];
                Cel2 = ws.Cells[nRow, nColFirst + 1];
                RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
                RowRng.Font.FontStyle = "Bold";
            }
            else if (type == 3)
            {
                RowRng.Font.Color = Color.White;
                RowRng.Font.FontStyle = "Bold";
                RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0.5; //Light blue
            }
            else if (type == 4)
            {
                //RowRng.Font.FontStyle = "Bold";
                RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else if (type == 5)
            {
                RowRng.Font.Color = Color.White;
                RowRng.Font.FontStyle = "Bold";
                RowRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                RowRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5;
                RowRng.Interior.TintAndShade = 0; //blue
            }

        }

        private void draw_LineEdge(Excel.Worksheet ws, System.Object Cel1, System.Object Cel2)
        {
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);

            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
        }

        public void ExcelWrite(string nSheet, int row, int col, List<string> Row_data)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            CellValSet(ws, row, col, Row_data);
        }
        public void ExcelWrite_Data(string nSheet, int row, int col, List<string> Row_data, List<string> Header_List)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            CellValSet_Data(ws, row, col, Row_data, Header_List);
        }

        public void ExcelWrite_Header(string nSheet, int row, int col, List<string> Row_data)
        {
            this.Workbook = this.App.Workbooks.get_Item(1);
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            CellValSet_Header(ws, row, col, Row_data);
        }

        

#if (false)
        public void CellValSet(int nSheet, int nRow, int nColFirst, List<string> ValueList)
        {
            // Note: This method is significantly slower than the overload which receives the Worksheet argument directly.

            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            this.CellValSet(ws, nRow, nColFirst, ValueList);
        }
#endif

        public Excel.Range GetRangeXL(Excel.Worksheet ws, int nRow1, int nCol1, int nRow2, int nCol2)
        {
            Excel.Range CellRng1 = (Excel.Range)ws.Cells[nRow1, nCol1];
            Excel.Range CellRng2 = (Excel.Range)ws.Cells[nRow2, nCol2];
            Excel.Range NewRng = this.App.get_Range((object)CellRng1, (object)CellRng2);
            return NewRng;
        }
        public Excel.Range GetRangeXL(Excel.Worksheet ws, int nRow, int nCol)
        {
            Excel.Range NewRng = this.GetRangeXL(ws, nRow, nCol, nRow, nCol);
            return NewRng;
        }

        public Excel.Range GetLineRangeXL(Excel.Worksheet ws, int nRow, int length)
        {
            Excel.Range NewRng = this.GetRangeXL(ws, nRow, 1, nRow, length);
            return NewRng;
        }

        public void Cformat_LineColor(Excel.Worksheet Sheet, int nRow, int length, Microsoft.Office.Interop.Excel.XlThemeColor xlThemeColor, double tintandshade_value)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, 1, nRow, length);
            RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
            RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
            RowRng_cell.Interior.ThemeColor = xlThemeColor;
            RowRng_cell.Interior.TintAndShade = tintandshade_value;
            RowRng_cell.Interior.PatternTintAndShade = 0;

            //  Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262  //Light Gray
            //  Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629 //Light Green
            //  Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943 //Light Orange
            //  Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943 //Orange
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314 //Light blue
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629 //blue
            //  Excel.XlThemeColor.xlThemeColorLight1, 1 //White?
        }

        public void Cformat_PartialColor(Excel.Worksheet Sheet, int nRow, int StartCol, int StopCol, Microsoft.Office.Interop.Excel.XlThemeColor xlThemeColor, double tintandshade_value)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, StartCol, nRow, StopCol);
            RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
            RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
            RowRng_cell.Interior.ThemeColor = xlThemeColor;
            RowRng_cell.Interior.TintAndShade = tintandshade_value;
            RowRng_cell.Interior.PatternTintAndShade = 0;

            //  Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262  //Light Gray
            //  Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629 //Light Green
            //  Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943 //Light Orange
            //  Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943 //Orange
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314 //Light blue
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629 //blue
            //  Excel.XlThemeColor.xlThemeColorLight1, 1 //White?
        }

        public void Cformat_Color(Excel.Worksheet Sheet, int nRow, int nCol, Microsoft.Office.Interop.Excel.XlThemeColor xlThemeColor, double tintandshade_value)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, nCol, nRow, nCol);
            RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
            RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
            RowRng_cell.Interior.ThemeColor = xlThemeColor;
            RowRng_cell.Interior.TintAndShade = tintandshade_value;
            RowRng_cell.Interior.PatternTintAndShade = 0;
            RowRng_cell.Font.Color = Color.White;

            //  Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262  //Light Gray
            //  Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629 //Light Green
            //  Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943 //Light Orange
            //  Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943 //Orange
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314 //Light blue
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629 //blue
            //  Excel.XlThemeColor.xlThemeColorLight1, 1 //White?
        }

        public void Cformat_Width(Excel.Worksheet Sheet, int nRow, int nCol, int Width)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, nCol, nRow, nCol);
            RowRng_cell.ColumnWidth = Width;
        }
        public void Cformat_Color(string Sheet, int nRow, int nRow_end, int nCol, int nCol_end, Color color)
        {
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[Sheet];
            Excel.Range RowRng_cell = this.GetRangeXL(ws, nRow, nCol, nRow_end, nCol_end);
            RowRng_cell.Interior.Color = color;
        }

        public void Rformat_Color(Excel.Worksheet Sheet, int nRow, int nCol, Microsoft.Office.Interop.Excel.XlThemeColor xlThemeColor, double tintandshade_value)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, 1, nRow, nCol);
            RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
            RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
            RowRng_cell.Interior.ThemeColor = xlThemeColor;
            RowRng_cell.Interior.TintAndShade = tintandshade_value;
            RowRng_cell.Interior.PatternTintAndShade = 0;

            //  Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262  //Light Gray
            //  Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629 //Light Green
            //  Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943 //Light Orange
            //  Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943 //Orange
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314 //Light blue
            //  Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629 //blue
            //  Excel.XlThemeColor.xlThemeColorLight1, 1 //White?
        }

        public void Rformat_Width(Excel.Worksheet Sheet, int nRow, int nCol, int Width)
        {
            Excel.Range RowRng_cell = this.GetRangeXL(Sheet, nRow, 1, nRow, nCol);
            RowRng_cell.ColumnWidth = Width;
            RowRng_cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            RowRng_cell.NumberFormat = "General"; //option = "0", "General", "0.00";
            RowRng_cell.Font.Name = "Arial";
            RowRng_cell.Font.FontStyle = "Bold";
            RowRng_cell.Font.Size = 10;
            RowRng_cell.Font.Color = Color.White;
        }

        public int Cell_WriteTitle(Excel.Worksheet Sheet, string Value, int nRow, int nCol, int length)
        {
            if (Value == null) return nRow;
            int next_row = nRow;

            var ValueArray = new object[1, length];

            for (var col = 0; col <= (length - 1); col++)
            {
                if (col == nCol - 1)
                {
                    ValueArray[0, col] = Value;
                }
                else
                {
                    ValueArray[0, col] = "";
                }
                
            }

            System.Object Cel1 = Sheet.Cells[nRow, 1];
            System.Object Cel2 = Sheet.Cells[nRow, length - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Bold";
            RowRng.Font.Size = 12;
            RowRng.Font.Color = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 1;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;

            RowRng.Value2 = ValueArray;

            next_row++;

            return next_row;
        }

        public int Cell_WriteTrigger(Excel.Worksheet Sheet, int nRow, int nColFirst, List<List<string>> Group_rows)
        {
            int return_row = nRow + Group_rows.Count;
            var ValueArray = new object[Group_rows.Count, Group_rows[0].Count];

            int row_index = 0; 
            foreach (List<string> Each_row in Group_rows)
            {
                int col_index = 0;
                foreach (string Value_Col in Each_row)
                {
                    ValueArray[row_index, col_index] = Value_Col;
                    col_index++;
                }
                row_index++;
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[return_row - 1, nColFirst + Group_rows[0].Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);
            RowRng.Value2 = ValueArray;

            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Regular";
            RowRng.Font.Size = 10;
            RowRng.Font.Strikethrough = false;
            RowRng.Font.Subscript = false;
            RowRng.Font.OutlineFont = false;
            RowRng.Font.Shadow = false;
            RowRng.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            RowRng.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 0;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;
            RowRng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            RowRng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            RowRng.ColumnWidth = 8;
            RowRng.RowHeight = 15;

            return return_row;
        }

        public void Cell_WriteHeader(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList, bool orientation_90)
        {
            bool Orient = orientation_90;

            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Bold";
            RowRng.Font.Size = 11;
            RowRng.Font.Strikethrough = false;
            RowRng.Font.Subscript = false;
            RowRng.Font.OutlineFont = false;
            RowRng.Font.Shadow = false;
            RowRng.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            RowRng.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 0;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;
            RowRng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            RowRng.VerticalAlignment = XlVAlign.xlVAlignBottom;
            if (Orient)
            {
                RowRng.Orientation = 90;
            }
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            RowRng.Value2 = ValueArray;
        }

        public void CellValSet(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList)
        {
            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);
            
            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Regular";
            RowRng.Font.Size = 10;
            RowRng.Font.Strikethrough = false;
            RowRng.Font.Subscript = false;
            RowRng.Font.OutlineFont = false;
            RowRng.Font.Shadow = false;
            RowRng.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            RowRng.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 0;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;


            string Low_limit = (string)ValueArray[0, 12];
            string Upper_limit = (string)ValueArray[0, 14];

            string Min_S1 = (string)ValueArray[0, 16];
            string Max_S1 = (string)ValueArray[0, 17];
            string Min_S2 = (string)ValueArray[0, 18];
            string Max_S2 = (string)ValueArray[0, 19];
            string Min_S3 = (string)ValueArray[0, 20];
            string Max_S3 = (string)ValueArray[0, 21];

            RowRng.Value2 = ValueArray;

            Cell_format_B(Sheet, nRow, 1, "C", "General");
            Cell_format_B(Sheet, nRow, 2, "C", "General");
            Cell_format_B(Sheet, nRow, 3, "C", "General");
            Cell_format_B(Sheet, nRow, 4, "C", "General");
            Cell_format_B(Sheet, nRow, 5, "C", "General");
            Cell_format_B(Sheet, nRow, 6, "C", "General");
            Cell_format_B(Sheet, nRow, 7, "C", "General");
            Cell_format_B(Sheet, nRow, 8, "C", "General");
            Cell_format_B(Sheet, nRow, 9, "C", "General");
            Cell_format_B(Sheet, nRow, 10, "C", "General");
            Cell_format_B(Sheet, nRow, 11, "C", "General");
            Cell_format_B(Sheet, nRow, 12, "C", "General");
            Cell_format_B(Sheet, nRow, 13, "C", "0.00"); //Limit_L
            Cell_format_B(Sheet, nRow, 14, "C", "0.00"); //Limit_TYP
            Cell_format_B(Sheet, nRow, 15, "C", "0.00"); //Limit_U
            Cell_format_B(Sheet, nRow, 16, "C", "General"); //unit
            Cell_format_C(Sheet, nRow, 17, "C", "0.00", Low_limit, Min_S1, "min"); //S1_min
            Cell_format_C(Sheet, nRow, 18, "C", "0.00", Upper_limit, Max_S1, "max"); //S1_max
            Cell_format_C(Sheet, nRow, 19, "C", "0.00", Low_limit, Min_S2, "min"); //S2_min
            Cell_format_C(Sheet, nRow, 20, "C", "0.00", Upper_limit, Max_S2, "max"); //S2_max
            Cell_format_C(Sheet, nRow, 21, "C", "0.00", Low_limit, Min_S3, "min"); //S3_min
            Cell_format_C(Sheet, nRow, 22, "C", "0.00", Upper_limit, Max_S3, "max"); //S3_max
            Cell_format_B(Sheet, nRow, 23, "L", "General"); //condition
        }

        public void CellValSet_Data(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList, List<string> Header_list)
        {
            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Regular";
            RowRng.Font.Size = 10;
            RowRng.Font.Strikethrough = false;
            RowRng.Font.Subscript = false;
            RowRng.Font.OutlineFont = false;
            RowRng.Font.Shadow = false;
            RowRng.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            RowRng.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 0;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            string Low_limit = "";
            string Upper_limit = "";
            
            RowRng.Value2 = ValueArray;
            int nCol = 1;

            for (int i = 0; i < Header_list.Count; i++)
            {
                if (Header_list[i].Contains("Limit_L"))
                {
                    Low_limit = (string)ValueArray[0, i];
                    Cell_format_B(Sheet, nRow, nCol, "C", "0.00");
                }
                else if (Header_list[i].Contains("Limit_U"))
                {
                    Upper_limit = (string)ValueArray[0, i];
                    Cell_format_B(Sheet, nRow, nCol, "C", "0.00");
                }
                else if (Header_list[i].Contains("_Max"))
                {
                    string data_max = (string)ValueArray[0, i];
                    Cell_format_C(Sheet, nRow, nCol, "C", "0.00", Upper_limit, data_max, "max"); //max
                }
                else if (Header_list[i].Contains("_Min"))
                {
                    string data_min = (string)ValueArray[0, i];
                    Cell_format_C(Sheet, nRow, nCol, "C", "0.00", Low_limit, data_min, "min"); //min
                }
                else if (Header_list[i].Contains("Description"))
                {
                    Cell_format_B(Sheet, nRow, nCol, "L", "General"); //condition
                }
                else
                {
                    Cell_format_B(Sheet, nRow, nCol, "C", "General");
                }
                nCol++;
            }
        }

        public void CellValSet_Header(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList)
        {
            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);
            RowRng.NumberFormat = "0.00";
            RowRng.Font.Name = "Calibri";
            RowRng.Font.FontStyle = "Regular";
            RowRng.Font.Size = 10;
            RowRng.Font.Strikethrough = false;
            RowRng.Font.Subscript = false;
            RowRng.Font.OutlineFont = false;
            RowRng.Font.Shadow = false;
            RowRng.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            RowRng.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            RowRng.Font.TintAndShade = 0;
            RowRng.Font.ThemeFont = XlThemeFont.xlThemeFontNone;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            RowRng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            
            RowRng.Value2 = ValueArray;

            int nCol = 1;

            foreach (string item in ValueList)
            {
                string Header_string = item.Trim();
            
                if (Header_string.Contains("Spec ID")) Cell_format_A(Sheet, nRow, nCol, 15, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("Test Name")) Cell_format_A(Sheet, nRow, nCol, 36, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("Band")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
                if (Header_string.Contains("CA")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
                if (Header_string.Contains("Param_ID")) Cell_format_A(Sheet, nRow, nCol, 12, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("Temp")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629); //Light Green
                if (Header_string.Contains("VSWR")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
                if (Header_string.Contains("PA MODE")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
                if (Header_string.Contains("LNA GAIN")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
                if (Header_string.Contains("SIGNAL")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("Pout_dBm")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("WAVEFORM")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("SIGNAL")) Cell_format_A(Sheet, nRow, nCol, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
                if (Header_string.Contains("_Freq")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue
                if (Header_string.Contains("Limit_")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
                if (Header_string.Contains("Typical")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
                if (Header_string.Contains("Unit")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
                if (Header_string.Contains("S1_")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue
                if (Header_string.Contains("S2_")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
                if (Header_string.Contains("S3")) Cell_format_A(Sheet, nRow, nCol, 10, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
                if (Header_string.Contains("Description")) Cell_format_A(Sheet, nRow, nCol, 40, Excel.XlThemeColor.xlThemeColorLight1, 1); //White?
                nCol++;
            }

            //Cell_format_A(Sheet, nRow, 1, 15, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 2, 36, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 3, 8, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
            //Cell_format_A(Sheet, nRow, 4, 12, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 5, 8, Excel.XlThemeColor.xlThemeColorAccent6, 0.599963377788629); //Light Green
            //Cell_format_A(Sheet, nRow, 6, 10, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
            //Cell_format_A(Sheet, nRow, 7, 8, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
            //Cell_format_A(Sheet, nRow, 8, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 9, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 10, 8, Excel.XlThemeColor.xlThemeColorLight1, 0.499984740745262); //Light Gray
            //Cell_format_A(Sheet, nRow, 11, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue
            //Cell_format_A(Sheet, nRow, 12, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue

            //Cell_format_A(Sheet, nRow, 13, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
            //Cell_format_A(Sheet, nRow, 14, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
            //Cell_format_A(Sheet, nRow, 15, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue
            //Cell_format_A(Sheet, nRow, 16, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.599963377788629); //blue

            //Cell_format_A(Sheet, nRow, 17, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue
            //Cell_format_A(Sheet, nRow, 18, 10, Excel.XlThemeColor.xlThemeColorAccent5, 0.799981688894314); //Light blue

            //Cell_format_A(Sheet, nRow, 19, 10, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange
            //Cell_format_A(Sheet, nRow, 20, 10, Excel.XlThemeColor.xlThemeColorAccent4, 0.399945066682943); //Light Orange

            //Cell_format_A(Sheet, nRow, 21, 10, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
            //Cell_format_A(Sheet, nRow, 22, 10, Excel.XlThemeColor.xlThemeColorAccent2, 0.399945066682943); //Orange
            //Cell_format_A(Sheet, nRow, 23, 40, Excel.XlThemeColor.xlThemeColorLight1, 1); //White?

            
        }

        

        public void Cell_format_A(Excel.Worksheet Sheet, int Row, int Col, float Width, Microsoft.Office.Interop.Excel.XlThemeColor xlThemeColor, double tintandshade_value)
        {
            System.Object Cel1 = Sheet.Cells[Row, Col];
            System.Object Cel2 = Sheet.Cells[Row, Col];
            Excel.Range RowRng_cell = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
            RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
            RowRng_cell.Interior.ThemeColor = xlThemeColor;
            RowRng_cell.Interior.TintAndShade = tintandshade_value;
            RowRng_cell.Interior.PatternTintAndShade = 0;
            RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            RowRng_cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            RowRng_cell.ColumnWidth = Width;
            
        }

        public void Cell_format_B(Excel.Worksheet Sheet, int Row, int Col, string align, string number_format)
        {
            System.Object Cel1 = Sheet.Cells[Row, Col];
            System.Object Cel2 = Sheet.Cells[Row, Col];
            Excel.Range RowRng_cell = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            if (align.ToUpper().Trim().Contains("C")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (align.ToUpper().Trim().Contains("L")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            if (align.ToUpper().Trim().Contains("R")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            RowRng_cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RowRng_cell.NumberFormat = number_format.Trim(); //option = "0", "General", "0.00";
        }

        public void Cell_format_C(Excel.Worksheet Sheet, int Row, int Col, string align, string number_format,string Limit, string data, string order)
        {
            System.Object Cel1 = Sheet.Cells[Row, Col];
            System.Object Cel2 = Sheet.Cells[Row, Col];
            Excel.Range RowRng_cell = (Excel.Range)Sheet.get_Range(Cel1, Cel2);

            if (align.ToUpper().Trim().Contains("C")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (align.ToUpper().Trim().Contains("L")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            if (align.ToUpper().Trim().Contains("R")) RowRng_cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            RowRng_cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RowRng_cell.NumberFormat = number_format.Trim(); //option = "0", "General", "0.00";

            bool IsPass = false;
            
            try
            {
                double test_limit = Convert.ToDouble(Limit.Trim());
                double test_data = Convert.ToDouble(data.Trim());
                if(order.Trim().ToUpper().Contains("MIN"))
                {
                    if(test_limit <= test_data)
                    {
                        IsPass = true;
                    }
                }
                else
                {
                    if (test_limit > test_data)
                    {
                        IsPass = true;
                    }
                }

                if (IsPass) 
                { 
                    RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
                    RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                    RowRng_cell.Interior.Color = 5296274; //Yellow Green
                    RowRng_cell.Interior.TintAndShade = 0;
                    RowRng_cell.Interior.PatternTintAndShade = 0;
                }
                else
                {
                    RowRng_cell.Interior.Pattern = XlPattern.xlPatternSolid;
                    RowRng_cell.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                    RowRng_cell.Interior.Color = 255; //RED
                    RowRng_cell.Interior.TintAndShade = 0;
                    RowRng_cell.Interior.PatternTintAndShade = 0;
                }
            }
            catch
            {

            }

            
            
        }

        public Band_Condition ExcelUpdate_LoadSheetFromCM(Excel_File Excel, string nSheet, Excel_Base.TestConfig INIfile)
        {
            int Start_row = 0;
            int End_row = 0;
            int End_col = 0;
            int sample_1 = 0;
            int Worst_Col = 0;

            Band_Condition Band_Info = new Band_Condition();
            Excel.Worksheet ws = (Excel.Worksheet)Excel.App.Worksheets[nSheet];


            //find column index with INI match? 
            //step 1 : find header row, column index, row count

            List<string> Header_List = new List<string>();

            Find_Index(ws, INIfile.Header_Start_ID, ref Start_row, ref End_row, ref End_col, ref sample_1, ref Worst_Col, ref Header_List, INIfile);

            INIfile.Selected_Headers = Header_List;

            //step 2 : load column data from row +1, transfer to Band_info.Members
            int Index = 1;
            Start_row = ++Start_row ; //grap row data from next row after header
            End_row = --End_row; //To just before "END" row index

            foreach (string Col_Header in Header_List)
            {
                if (EXACT_Matching_Header(Col_Header, INIfile.Test_SpecID)) Band_Info.Test_SpecID = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Band)) Band_Info.Band = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.CA_Band2)) Band_Info.CA_Band2 = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.CA_Band3)) Band_Info.CA_Band3 = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.CA_Band4)) Band_Info.CA_Band4 = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Description)) Band_Info.Test_Name = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Parameter))
                {
                    Band_Info.Parameter = GetKey_GetValue(ws, Index, Start_row, End_row);
                    Band_Info.Direction = GetKey_GetValue(ws, Index, Start_row, End_row); //Add Dummy data to new direction col(TX,RX indicator) for array size adjustment.
                    for (int i = 0; i < Band_Info.Direction.Count; i++)
                    {
                        Band_Info.Direction[i] = "";
                    }
                }

                if (Matching_Header(Col_Header, INIfile.Input_Port)) Band_Info.Input_Port = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.Output_Port)) Band_Info.Output_Port = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.LNA_Gain_Mode)) Band_Info.LNA_Gain_Mode = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (Matching_Header(Col_Header, INIfile.Vbatt)) Band_Info.Vbatt = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.Vdd_LNA)) Band_Info.Vdd_LNA = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (Matching_Header(Col_Header, INIfile.TXIn_VSWR)) Band_Info.TXIn_VSWR = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.ANTout_VSWR)) Band_Info.ANTout_VSWR = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.ANTIn_VSWR)) Band_Info.ANTIn_VSWR = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.RXOut_VSWR)) Band_Info.RXOut_VSWR = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (EXACT_Matching_Header(Col_Header, INIfile.Temperature)) Band_Info.Temperature = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (Matching_Header(Col_Header, INIfile.Start_Freq)) Band_Info.Start_Freq = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.Stop_Freq)) Band_Info.Stop_Freq = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.IBW)) Band_Info.IBW = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (EXACT_Matching_Header(Col_Header, INIfile.PA_MODE)) Band_Info.PA_MODE = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.TXBand_In_RXtest)) Band_Info.TXBand_In_RXtest = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.Target_Pout)) Band_Info.Target_Pout = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (Matching_Header(Col_Header, INIfile.Signal_Standard)) Band_Info.Signal_Standard = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Waveform_Category)) Band_Info.Waveform_Category = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (EXACT_Matching_Header(Col_Header, INIfile.Test_Limit_L)) Band_Info.Test_Limit_L = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Test_Limit_Typ)) Band_Info.Test_Limit_Typ = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Test_Limit_U)) Band_Info.Test_Limit_U = GetKey_GetValue(ws, Index, Start_row, End_row);

                if (EXACT_Matching_Header(Col_Header, INIfile.Unit)) Band_Info.Unit = GetKey_GetValue(ws, Index, Start_row, End_row);
                if (EXACT_Matching_Header(Col_Header, INIfile.Compliance)) Band_Info.Compliance = GetKey_GetValue(ws, Index, Start_row, End_row);

                Index++;
            }

            if( sample_1 != 0) //mean if find "sample data" column in CM sheet
            {
                Band_Info.Sample_1_min = GetKey_GetValue(ws, sample_1, Start_row, End_row);
                Band_Info.Sample_1_max = GetKey_GetValue(ws, sample_1+1, Start_row, End_row);
                Band_Info.Sample_2_min = GetKey_GetValue(ws, sample_1+2, Start_row, End_row);
                Band_Info.Sample_2_max = GetKey_GetValue(ws, sample_1+3, Start_row, End_row);
                Band_Info.Sample_3_min = GetKey_GetValue(ws, sample_1+4, Start_row, End_row);
                Band_Info.Sample_3_max = GetKey_GetValue(ws, sample_1+5, Start_row, End_row);
                Band_Info.Worst_Condition_text = GetKey_GetValue(ws, Worst_Col, Start_row, End_row);
            }
            //repeat!
            
            return Band_Info;
        }

        public bool Matching_Header(string CM_HeaderName, List<string> ListToCompare)
        {
            bool IsMatched = false;

            foreach (string INI_Listed_Name in ListToCompare)
            {
                if (CM_HeaderName.Trim().ToUpper().Contains(INI_Listed_Name.Trim().ToUpper())) return true;
            }

            return IsMatched;
        }

        public bool EXACT_Matching_Header(string CM_HeaderName, List<string> ListToCompare)
        {
            bool IsMatched = false;

            foreach (string INI_Listed_Name in ListToCompare)
            {
                if (CM_HeaderName.Trim().ToUpper() == INI_Listed_Name.Trim().ToUpper()) return true;
            }

            return IsMatched;
        }

        public List<string> GetKey_GetValue(Excel.Worksheet ws, int Val_col, int Row_start, int Row_stop)
        {
            int numRows = Row_stop;
            var Val_Array = this.GetRangeVals(ws, Row_start, Row_stop, Val_col, Val_col);

            string Key = "";
            string Val = "";

            List<string> List_Data = new List<string>();

            for (int Row = 0; Row <= Val_Array.GetUpperBound(0); Row++)
            {
                try
                {
                    if (Row == 0)
                    {
                        List_Data.Clear();
                    }
                    Val = Val_Array[Row, 0]; //current List is 1D array from indexed Column, so Col = 0;
                    List_Data.Add(Val);
                }
                catch (Exception)
                {
                    int current_Col = Val_col;
                    int current_Row = Row;
                    string current_value = Val_Array[Row, Val_col];

                    StringBuilder ErrMsg = new StringBuilder();
                    ErrMsg.AppendFormat("Error:Load cell failure from GetKey_GetValue(...)");
                    ErrMsg.AppendFormat("\nPlease check below infomation");
                    ErrMsg.AppendFormat("\nRow Num :" + current_Row.ToString() + "/ Col Num :" + current_Col.ToString());
                    ErrMsg.AppendFormat("\nValue :" + current_value);
                    ClsMsgBox.Show("Error on Load Band data from sheet",ErrMsg.ToString());
                    Environment.Exit(0);
                }

            }

            return List_Data;
        }

        public string[,] ReadData_From_WorkSheet(Excel.Worksheet Temp_Worksheet, int Start_Row, int Stop_Row, int Start_Col, int Stop_Col)
        {
            //위의 함수와는 다른 함수이며, 현재 이 프로그램에서 이 함수는 쓰이지 않고 있다. 이 함수를 쓰게되면, 위의 함수 보다 속도가 더 빠르다..차이는?

            int R1 = (Start_Row <= Stop_Row ? Start_Row : Stop_Row); //StartRow와 StopRaw 중에 작은 것을 R1으로 지정
            int R2 = (R1 == Start_Row ? Stop_Row : Start_Row); //나머지를 R2로 지정

            int C1 = (Start_Col <= Stop_Col ? Start_Col : Stop_Col); //상동
            int C2 = (C1 == Start_Col ? Stop_Col : Start_Col); //상동

            int NumRows = R2 - R1 + 1;
            int NumCols = C2 - C1 + 1;

            System.Object Cel1 = Temp_Worksheet.Cells[R1, C1];
            System.Object Cel2 = Temp_Worksheet.Cells[R2, C2];
            Excel.Range RowRng = (Excel.Range)Temp_Worksheet.get_Range(Cel1, Cel2); //지정된 범위의 cell data를 가져온다. 
            System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;


            string[,] Result_Arrays = new string[NumRows, NumCols]; //결과로 내보낼, 문자열 타입의 다차원 배열을 선언
            System.Object Values; //위의 Cell data가 System.Object[,] 다차원 array로 되어있으므로, array에서 개별 data를 꺼낼때 잠시 담아둘 system.object type

            for (int Row = 0; Row < NumRows; Row++)
            {
                for (int Col = 0; Col < NumCols; Col++)
                {
                    Values = ObjVals[Row + 1, Col + 1]; //ObjVals 에서 각각의 array 한개에 들어있는 값을 Value로 담는다. 
                    Result_Arrays[Row, Col] = (Values == null ? "" : Values.ToString()); //Value의 값이 null 이면 공백 문자("")를,아니면 결과 array에 string 으로 바꾸어서 담는다.
                }
            }

            return Result_Arrays;

        }

        public string[,] ReadData_From_WorkSheet(string Sheet, int Start_Row, int Stop_Row, int Start_Col, int Stop_Col)
        {
            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            int R1 = (Start_Row <= Stop_Row ? Start_Row : Stop_Row); //StartRow와 StopRaw 중에 작은 것을 R1으로 지정
            int R2 = (R1 == Start_Row ? Stop_Row : Start_Row); //나머지를 R2로 지정

            int C1 = (Start_Col <= Stop_Col ? Start_Col : Stop_Col); //상동
            int C2 = (C1 == Start_Col ? Stop_Col : Start_Col); //상동

            int NumRows = R2 - R1 + 1;
            int NumCols = C2 - C1 + 1;

            System.Object Cel1 = this.Worksheet.Cells[R1, C1];
            System.Object Cel2 = this.Worksheet.Cells[R2, C2];
            Excel.Range RowRng = (Excel.Range)this.Worksheet.get_Range(Cel1, Cel2); //지정된 범위의 cell data를 가져온다. 
            System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;


            string[,] Result_Arrays = new string[NumRows, NumCols]; //결과로 내보낼, 문자열 타입의 다차원 배열을 선언
            System.Object Values; //위의 Cell data가 System.Object[,] 다차원 array로 되어있으므로, array에서 개별 data를 꺼낼때 잠시 담아둘 system.object type

            for (int Row = 0; Row < NumRows; Row++)
            {
                for (int Col = 0; Col < NumCols; Col++)
                {
                    Values = ObjVals[Row + 1, Col + 1]; //ObjVals 에서 각각의 array 한개에 들어있는 값을 Value로 담는다. 
                    Result_Arrays[Row, Col] = (Values == null ? "" : Values.ToString()); //Value의 값이 null 이면 공백 문자("")를,아니면 결과 array에 string 으로 바꾸어서 담는다.
                }
            }

            return Result_Arrays;

        }

        public List<string> ReadRow_asList(string Sheet, int Start_Row, int Start_Col, int Stop_Col)
        {
            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            List<string> Read1D = new List<string>();

            int R1 = Start_Row; //StartRow와 StopRaw 중에 작은 것을 R1으로 지정
            int R2 = Start_Row; //나머지를 R2로 지정

            int C1 = (Start_Col <= Stop_Col ? Start_Col : Stop_Col); //상동
            int C2 = (C1 == Start_Col ? Stop_Col : Start_Col); //상동

            int NumRows = R2 - R1 + 1;
            int NumCols = C2 - C1 + 1;

            System.Object Cel1 = this.Worksheet.Cells[R1, C1];
            System.Object Cel2 = this.Worksheet.Cells[R2, C2];
            Excel.Range RowRng = (Excel.Range)this.Worksheet.get_Range(Cel1, Cel2); //지정된 범위의 cell data를 가져온다. 
            System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;


            string[,] Result_Arrays = new string[NumRows, NumCols]; //결과로 내보낼, 문자열 타입의 다차원 배열을 선언
            System.Object Values; //위의 Cell data가 System.Object[,] 다차원 array로 되어있으므로, array에서 개별 data를 꺼낼때 잠시 담아둘 system.object type

            for (int Row = 0; Row < NumRows; Row++)
            {
                for (int Col = 0; Col < NumCols; Col++)
                {
                    Values = ObjVals[Row + 1, Col + 1]; //ObjVals 에서 각각의 array 한개에 들어있는 값을 Value로 담는다. 
                    Result_Arrays[Row, Col] = (Values == null ? "" : Values.ToString()); //Value의 값이 null 이면 공백 문자("")를,아니면 결과 array에 string 으로 바꾸어서 담는다.
                }
            }

            for (int i = 0; i < NumCols; i++)
            {
                Read1D.Add(Result_Arrays[0, i]);
            }

            return Read1D;

        }

        public List<string> ReadCol_asList(string Sheet, int Start_Row, int Stop_Row, int Start_Col)
        {
            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            List<string> Read1D = new List<string>();

            int R1 = (Start_Row <= Stop_Row ? Start_Row : Stop_Row); //StartRow와 StopRaw 중에 작은 것을 R1으로 지정
            int R2 = (R1 == Start_Row ? Stop_Row : Start_Row); //나머지를 R2로 지정

            int C1 = Start_Col; //상동
            int C2 = Start_Col; //상동

            int NumRows = R2 - R1 + 1;
            int NumCols = C2 - C1 + 1;

            System.Object Cel1 = this.Worksheet.Cells[R1, C1];
            System.Object Cel2 = this.Worksheet.Cells[R2, C2];
            Excel.Range RowRng = (Excel.Range)this.Worksheet.get_Range(Cel1, Cel2); //지정된 범위의 cell data를 가져온다. 
            System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;


            string[,] Result_Arrays = new string[NumRows, NumCols]; //결과로 내보낼, 문자열 타입의 다차원 배열을 선언
            System.Object Values; //위의 Cell data가 System.Object[,] 다차원 array로 되어있으므로, array에서 개별 data를 꺼낼때 잠시 담아둘 system.object type

            for (int Row = 0; Row < NumRows; Row++)
            {
                for (int Col = 0; Col < NumCols; Col++)
                {
                    Values = ObjVals[Row + 1, Col + 1]; //ObjVals 에서 각각의 array 한개에 들어있는 값을 Value로 담는다. 
                    Result_Arrays[Row, Col] = (Values == null ? "" : Values.ToString()); //Value의 값이 null 이면 공백 문자("")를,아니면 결과 array에 string 으로 바꾸어서 담는다.
                }
            }

            for (int i = 0; i < NumRows; i++)
            {
                Read1D.Add(Result_Arrays[i, 0]);
            }

            return Read1D;

        }

        public void Write_SampleData_To_CM(string nSheet, Dictionary<string,int> Header_Index, Dictionary<string, int> Spec_IDX, Dictionary<string,List<string>> Sample_Data, Dictionary<string, List<string>> Worst_condition)
        {
            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];

            foreach (var CMSheet_SpecNum in Spec_IDX.Keys)
            {
                if(Sample_Data.ContainsKey(CMSheet_SpecNum))
                {
                    int ColNum = Header_Index["Spec_Index"] + 1;
                    int LoLimitNum = Header_Index["LoLimit_Index"] + 1;
                    int HiLimitNum = Header_Index["UpLimit_Index"] + 1;
                    System.Object Cel1 = ws.Cells[Spec_IDX[CMSheet_SpecNum] + 1, ColNum];
                    System.Object Cel2 = ws.Cells[Spec_IDX[CMSheet_SpecNum] + 1, HiLimitNum];
                    Excel.Range RowRng = ws.get_Range(Cel1, Cel2); //지정된 범위의 cell data를 가져온다. 
                    System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;

                    int Sample1_index = Header_Index["Sample_Index"] + 1;
                    int RowNum = Spec_IDX[CMSheet_SpecNum] + 1;

                    List<string> RawData = Sample_Data[CMSheet_SpecNum];
                    this.WriteData_1D_Row(ws, RowNum, Sample1_index, RawData);

                    double Lo_limit = -9999d;
                    double Hi_limit = 9999d;
                    try
                    {
                        if (ObjVals[1, HiLimitNum] != null) Hi_limit = Convert.ToDouble(ObjVals[1, HiLimitNum]);
                        if (ObjVals[1, HiLimitNum - 2] != null) Lo_limit = Convert.ToDouble(ObjVals[1, HiLimitNum - 2]);
                    }
                    catch
                    {

                    }

                    int index_min = 1;
                    foreach (string Result_MinMax in RawData)
                    {
                        string PassFail = "default";

                        if (index_min % 2 == 0)
                        {
                            if (Hi_limit != 9999 && Result_MinMax.Trim() != "") 
                            {
                                if (Convert.ToDouble(Result_MinMax.Trim()) <= Hi_limit)
                                {
                                    PassFail = "Pass";
                                    RowRng[1, Header_Index["Sample_Index"] + index_min].Interior.Color = 3407718;
                                }
                            }
                            if (Hi_limit != 9999 && Result_MinMax.Trim() != "" )
                            {
                                if (Convert.ToDouble(Result_MinMax.Trim()) > Hi_limit)
                                {
                                    PassFail = "Fail";
                                    RowRng[1, Header_Index["Sample_Index"] + index_min].Interior.Color = Color.Red;
                                }
                            }

                        }
                        else
                        {
                            if (Lo_limit != -9999 && Result_MinMax.Trim() != "" )
                            {
                                if (Convert.ToDouble(Result_MinMax.Trim()) >= Lo_limit)
                                {
                                    PassFail = "Pass";
                                    RowRng[1, Header_Index["Sample_Index"] + index_min].Interior.Color = 3407718;
                                }
                            }
                            if (Lo_limit != -9999 && Result_MinMax.Trim() != "" )
                            {
                                if (Convert.ToDouble(Result_MinMax.Trim()) < Lo_limit)
                                {
                                    PassFail = "Fail";
                                    RowRng[1, Header_Index["Sample_Index"] + index_min].Interior.Color = Color.Red;
                                }
                            }
                        }

                        index_min++;
                    }

                    List<string> Condition = new List<string>();
                    if (Worst_condition[CMSheet_SpecNum].Count != 0) Condition.Add(Worst_condition[CMSheet_SpecNum][0]);
                    int condition_index = Header_Index["Condition_Index"] + 1;

                    this.WriteData_1D_Row(ws, RowNum, condition_index, Condition);
                }
            }
        }

        public void WriteData_1D_Row(string nSheet, int nRow, int nColFirst, List<string> ValueList)
        {
            // Note: This method is significantly slower than the overload which receives the Worksheet argument directly.

            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            this.WriteData_1D_Row(ws, nRow, nColFirst, ValueList);
        }

        public void WriteData_1D_Row(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList)
        {
            var ValueArray = new object[1, ValueList.Count];

            for (var col = 0; col <= (ValueList.Count - 1); col++)
            {
                ValueArray[0, col] = ValueList[col];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow, nColFirst + ValueList.Count - 1];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);
            RowRng.Value2 = ValueArray;
            
        }

        public void WriteData_1D_Col(string nSheet, int nRow, int nColFirst, List<string> ValueList)
        {
            // Note: This method is significantly slower than the overload which receives the Worksheet argument directly.

            Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            this.WriteData_1D_Col(ws, nRow, nColFirst, ValueList);
        }

        public void WriteData_1D_Col(Excel.Worksheet Sheet, int nRow, int nColFirst, List<string> ValueList)
        {
            var ValueArray = new object[ValueList.Count, 1];

            for (var row = 0; row <= (ValueList.Count - 1); row++)
            {
                ValueArray[row, 0] = ValueList[row];
            }

            System.Object Cel1 = Sheet.Cells[nRow, nColFirst];
            System.Object Cel2 = Sheet.Cells[nRow + ValueList.Count - 1, nColFirst];
            Excel.Range RowRng = (Excel.Range)Sheet.get_Range(Cel1, Cel2);
            RowRng.Value2 = ValueArray;
        }

        public void WriteData_ToArray(string Sheet, int nRow, int nColFirst, string[,] data_array, bool float_type)
        {
            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            int Rc = data_array.GetLength(0);
            int Cc = data_array.Length / Rc;

            var ValueArray = new object[Rc, Cc];
            if (float_type)
            {
                for (var Row = 0; Row < Rc; Row++)
                {
                    for (var Col = 0; Col < Cc; Col++)
                    {
                        ValueArray[Row, Col] = Convert.ToSingle(data_array[Row, Col]);
                    }
                }
            }
            else
            {
                ValueArray = data_array;
            }

            System.Object Cel1 = this.Worksheet.Cells[nRow, nColFirst];
            System.Object Cel2 = this.Worksheet.Cells[nRow + Rc - 1, nColFirst + Cc - 1];
            Excel.Range RowRng = this.Worksheet.get_Range(Cel1, Cel2);
            RowRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            RowRng.NumberFormat = "0.000"; //option = "0", "General", "0.00";
            RowRng.Value2 = ValueArray;
        }

        public string[,] GetRangeVals(Excel.Worksheet ws, int nRow1, int nRow2, int nCol1, int nCol2)
        {
            // Reorder Rows and Cols (may not be necessary)
            int R1 = (nRow1 <= nRow2 ? nRow1 : nRow2);
            int R2 = (R1 == nRow1 ? nRow2 : nRow1);
            int C1 = (nCol1 <= nCol2 ? nCol1 : nCol2);
            int C2 = (C1 == nCol1 ? nCol2 : nCol1);
            int NumRows = R2 - R1 + 1;
            int NumCols = C2 - C1 + 1;

            // Query Excel Range Values (objects)
            //Excel.Worksheet ws = (Excel.Worksheet)this.App.Worksheets[nSheet];
            System.Object Cel1 = ws.Cells[R1, C1];
            System.Object Cel2 = ws.Cells[R2, C2];
            Excel.Range RowRng = (Excel.Range)ws.get_Range(Cel1, Cel2);
            System.Object[,] ObjVals = (System.Object[,])RowRng.Value2;

            // Convert Object Array into String Array (note that obj array is '1-based', and structured as 'row,col')
            System.Object oVal;
            string[,] StrVals = new string[NumRows, NumCols];

            for (int Row = 0; Row < NumRows; Row++)
            {
                for (int Col = 0; Col < NumCols; Col++)
                {
                    oVal = ObjVals[Row + 1, Col + 1];
                    StrVals[Row, Col] = (oVal == null ? "" : oVal.ToString());
                }
            }

            return StrVals;
        }

        public void Find_Index(Excel.Worksheet ws, string Key_string, ref int Start_row, ref int End_row, ref int End_Col, ref int first_sample_Header, ref int condition_header, ref List<string> Header_list, Excel_Base.TestConfig INIfile)
        {
            var Row_Array = this.GetRangeVals(ws, 1, 1000, 1, 1); //find 1st column, 1~1000 row to find header, end row
            bool first_Sample_found = false;
            
            End_row = 0;

            try
            {
                for (int Row = 0; Row <= Row_Array.GetUpperBound(0); Row++)
                {
                    string Row_read = Row_Array[Row, 0];

                    if (Row_read.ToUpper().Contains(Key_string.Trim().ToUpper()))
                    {
                        Start_row = Row + 1; //Set CM sheet row_start_index
                    }

                    if (Row_read.ToUpper().Contains("END"))
                    {
                        End_row = Row + 1; //Set CM sheet row_stop_index
                    }
                }

                bool Contain_sample_data = false;
                var SampleCol_Array = this.GetRangeVals(ws, 1, 1, 1, 500);
                for (int sampleCol = 0; sampleCol < SampleCol_Array.GetUpperBound(1); sampleCol++)
                {
                    string Col_read = SampleCol_Array[0, sampleCol];
                    //Header_list.Add(Col_read); //Build CM sheet Header

                    //if (Col_read.Trim().ToUpper().Contains("COMPLIANCE"))
                    if (Col_read.Trim().ToUpper().Contains("#CONDITION_START_INDEX"))
                    {
                        condition_header = sampleCol + 1;
                    }

                    if ((!first_Sample_found) && Col_read.Trim().ToUpper().Contains("SAMPLE#1"))
                    {
                        first_sample_Header = sampleCol + 1;
                        first_Sample_found = true;
                    }
                }

                var Col_Array = this.GetRangeVals(ws, Start_row, Start_row, 1, 500);

                for (int Col = 0; Col < Col_Array.GetUpperBound(1); Col++)
                {
                    string Col_read = Col_Array[0, Col];
                    Header_list.Add(Col_read); //Build CM sheet Header

                    //if (Col_read.Trim().ToUpper().Contains("COMPLIANCE"))
                    if(Matching_Header(Col_read, INIfile.Compliance))
                    {
                        End_Col = Col + 1; //Set CM sheet Col_stop_index
                        return;
                    }
                }
                //if can't find index then it will hit this statement

                if (End_Col == 0) End_Col = 500;
                if (End_row == 0) End_row = 1000;
            }
            catch (Exception)
            {
                StringBuilder ErrMsg = new StringBuilder();
                ErrMsg.AppendFormat("Error: Find Header Index failure in void Find_Index(...)");
                ClsMsgBox.Show("Error on Load Header sheet", ErrMsg.ToString());
                Environment.Exit(0);
            }
        }

        public List<string> Quick_FindHeader(string Sheet, string key_string, ref int Start_Row, ref int End_Col, ref int End_RoW)
        {
            List<string> Header_items = new List<string>();

            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];
            
            string[,] Array_2D = ReadData_From_WorkSheet(this.Worksheet, Start_Row, 100, 1, End_Col);

            for (int i = 0; i < 100; i++) //i=row
            {
                for (int j = 0; j < End_Col; j++) //j=col
                {
                    if(Array_2D[i,j].Trim().ToUpper() == key_string.Trim().ToUpper())
                    {
                        for (int k = 0; k < End_Col; k++)
                        {
                            if(Array_2D[i, k]!="") Header_items.Add(Array_2D[i, k]);
                        }

                        Start_Row = i + 1;
                        End_Col = Header_items.Count;

                        string[,] Array_EndRow = ReadData_From_WorkSheet(this.Worksheet, Start_Row, 150000, j, j);

                        for (int l = 0; l < Array_EndRow.Length; l++)
                        {
                            if (Array_EndRow[l, 0] == "")
                            {
                                End_RoW = l + Start_Row; break;
                            }
                        }

                        return Header_items;
                    }
                }
            }

            return Header_items;
        }

        public Dictionary<string,int> Find_CM_Header_Index(string Sheet, int Start_Row, int End_RoW, int End_Col, ref Dictionary<string,int> SpecID_DIC)
        {
            Dictionary<string, int> Header_info = new Dictionary<string, int>();

            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            string[,] Array_2D = ReadData_From_WorkSheet(this.Worksheet, Start_Row, End_RoW, 1, End_Col);

            for (int i = 0; i < End_RoW; i++) //i=row
            {
                for (int j = 0; j < End_Col; j++) //j=col
                {
                    if (Array_2D[i, j].Trim().ToUpper() == "SPEC NUMBER")
                    {
                        Header_info.Add("Spec_Index", j);
                    }

                    if (Array_2D[i, j].Trim().ToUpper() == "SAMPLE#1" ||
                        Array_2D[i, j].Trim().ToUpper() == "SAMPLE#01" ||
                        Array_2D[i, j].Trim().ToUpper() == "SAMPLE #1")
                    {
                        Header_info.Add("Sample_Index", j);
                    }

                    if (Array_2D[i, j].Trim().ToUpper() == "LOWER LIMIT")
                    {
                        Header_info.Add("LoLimit_Index", j);
                    }

                    if (Array_2D[i, j].Trim().ToUpper() == "UPPER LIMIT")
                    {
                        Header_info.Add("UpLimit_Index", j);
                    }

                    if (Array_2D[i, j].Trim().ToUpper() == "#CONDITION_START_INDEX")
                    {
                        Header_info.Add("Condition_Index", j);
                    }
                }
            }

            if(Header_info.ContainsKey("Spec_Index"))
            {
                string[,] Array_EndRow = ReadData_From_WorkSheet(this.Worksheet, 1, 10000, Header_info["Spec_Index"] + 1, Header_info["Spec_Index"] + 1);

                Dictionary<string, int> SpecList = new Dictionary<string, int>();

                int EndRow_Index = 0;
                bool UnderSpecID = false;

                for (int i = 0; i < 9999; i++)
                {  
                    if (UnderSpecID)
                    {
                        if (Array_EndRow[i, 0].Trim().ToUpper() == "END")
                        {
                            break;
                        }
                        else
                        {
                            if (!SpecList.ContainsKey(Array_EndRow[i, 0]))
                            {
                                SpecList.Add(Array_EndRow[i, 0], i);
                                EndRow_Index = i;
                            }
                        }
                    }

                    if (Array_EndRow[i, 0].Trim().ToUpper() == "SPEC NUMBER") UnderSpecID = true;

                }

                SpecID_DIC = SpecList;
                Header_info.Add("EndRow_Index", EndRow_Index);
            }

            return Header_info;
        }

        public List<string> Find_Header(string Sheet, string index_ID, string keyTofind_LastRow, ref int Header_start_row, ref int Header_end_row, ref int Header_start_col, ref int Header_end_col)
        {
            this.Worksheet = (Excel.Worksheet)this.App.Worksheets[Sheet];

            List<string> Header_items = new List<string>();

            string[,] Index_Col = ReadData_From_WorkSheet(this.Worksheet, Header_start_row, Header_end_row, Header_start_col, Header_start_col);
            int Is_emptyCell_1 = 0;

            for (int i = 0; i < Index_Col.Length; i++)
            {
                if (Index_Col[i, 0] == index_ID)
                {
                    Header_start_row = i + 1;
                }

                if (Index_Col[i, 0].Trim().ToUpper() == keyTofind_LastRow.Trim().ToUpper())
                {
                    Header_end_row = i + 1;
                }
                else if (Is_emptyCell_1 == 100000)
                {
                    Header_end_row = i - 100000;
                    break;
                }
                else if (Index_Col[i, 0].Trim().ToUpper() == "")
                {
                    Header_end_row = i + 1;
                    Is_emptyCell_1++;
                }
                else
                {
                    Header_end_row = i + 1;
                    Is_emptyCell_1 = 0;
                }
            }

            if (Header_start_row != 0)
            {
                string[,] Index_row = ReadData_From_WorkSheet(this.Worksheet, Header_start_row, Header_start_row, 1, 500);
                int Is_emptyCell_2 = 0;

                for (int i = 0; i < Index_row.Length; i++)
                {
                    if (Is_emptyCell_2 == 10)
                    {
                        Header_end_col = i - 10;
                        break;
                    }
                    else if (Index_row[0, i] == "")
                    {
                        Is_emptyCell_2++;
                    }
                    else
                    {
                        Header_end_col = i;
                        Is_emptyCell_2 = 0;
                    }
                }
            }

            string[,] Header = ReadData_From_WorkSheet(this.Worksheet, Header_start_row, Header_start_row, 1, Header_end_col);
            for (int i = 0; i < Header.Length; i++)
            {
                Header_items.Add(Header[0, i]);
            }

            return Header_items;

        }

        public void Workbook_Format()
        {
            Excel.Workbook wb = this.App.ActiveWorkbook;
            this.Workbook_Format(wb);
        }

        public void Workbook_Format(Excel.Workbook wb)
        {
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            Excel.Range rng = ws.get_Range(ws.Cells.get_Item(9, 1), ws.Cells.get_Item(9, 38));
            rng.Font.Bold = true;
            rng.Orientation = 90;
        }

        public void Quit()
        {
            int ProcID = this.ProcID;
            //this.App.Quit();
            ClsProcess.KillProcByID(ProcID);
        }
        public void Quit(int P_ID)
        {
            int ProcID = P_ID;
            //this.App.Quit();
            ClsProcess.KillProcByID(ProcID);
        }


        //================================== System Method (using System.Diagnostics) ==================================//

        public class ClsProcess
        {
            public static int GetLastProcID(string ProcName)
            {
                Process[] ProcAry = Process.GetProcessesByName(ProcName);
                int LastProcID = 0;
                DateTime LastTime = new DateTime(2000, 1, 1);

                foreach (Process Proc in ProcAry)
                {
                    if (Proc.StartTime.CompareTo(LastTime) > 0)
                    {
                        LastTime = Proc.StartTime;
                        LastProcID = Proc.Id;
                    }
                }
                return LastProcID;
            }

            public static void KillProcByName(string ProcName)
            {
                Process[] ProcAry = Process.GetProcessesByName(ProcName);
                foreach (Process Proc in ProcAry)
                {
                    Proc.Kill();
                }
                //System.Media.SystemSounds.Exclamation.Play();
            }

            public static void KillProcByID(int ProcID)
            {
                Process Proc = Process.GetProcessById(ProcID);
                Proc.Kill();
            }
        }

    }
}
