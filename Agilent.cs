using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Agilent.Ag3446x.Interop;
using AutoTest;

namespace AgilentMultimeter
{
    public interface DMMInterface
    {
        string ResourceName { get; set; }
        string ID { get; set; }
        double Range { get; set; }
        double NPLC { get; set; }
        bool InputImpedance { get; set; }
        bool NullState { get; set; }
        double NullValue { get; set; }
        double Constant { get; set; }
        double RawValue { get; set; }
        bool Run { get; set; }
        Thread Work { get; set; }

        void DMM_SetID(string text);
        void DMM_SelectAutoRange(int AutoZeroID);
        void DMM_SelectRange(int SelectedIndex);
        void DMM_SelectNPLC(int SelectedIndex);
    }

    //public interface AgilentInterface
    //{
    //    Error Open(DMMInterface inter);
    //    //void Close(DMMInterface DMM1);
    //    void GetData(DMMInterface DMM1, out double[] Data);
    //    Error getError(string title, DMMInterface driver1);
    //}

}
