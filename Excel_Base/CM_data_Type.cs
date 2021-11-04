using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Base
{
    public class Band_Condition
    {
        public List<int> Key_Index;
        public List<string> Test_Name;
        public List<string> Test_SpecID;
        public List<string> Band;
        public List<string> CA_Band2;
        public List<string> CA_Band3;
        public List<string> CA_Band4;
        public List<string> Direction;
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

        public List<string> Sample_1_min;
        public List<string> Sample_1_max;
        public List<string> Sample_2_min;
        public List<string> Sample_2_max;
        public List<string> Sample_3_min;
        public List<string> Sample_3_max;

        public List<string> Worst_Condition_text;

        public Band_Condition()
        {
            this.clear();
        }

        public void clear()
        {
            this.Key_Index = new List<int>();
            this.Test_Name = new List<string>();
            this.Test_SpecID = new List<string>();
            this.Band = new List<string>();
            this.CA_Band2 = new List<string>();
            this.CA_Band3 = new List<string>();
            this.CA_Band4 = new List<string>();
            this.Direction = new List<string>();
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

            this.Sample_1_min = new List<string>();
            this.Sample_1_max = new List<string>();
            this.Sample_2_min = new List<string>();
            this.Sample_2_max = new List<string>();
            this.Sample_3_min = new List<string>();
            this.Sample_3_max = new List<string>();

            this.Worst_Condition_text = new List<string>();
        }

    }
}
