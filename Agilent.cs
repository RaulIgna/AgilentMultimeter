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

    public class AgilentInterface
    {
        public static Error Open(DMMInterface inter, bool IsNewMultimeter)
        {
            if(IsNewMultimeter)
            {
                return Agilent_34465A_LIB.Agilent_34465A_LIB.Open(inter);
            }
            else
            {
                return Agilent_34411A_LIB.Agilent_34411A_LIB.Open(inter);
            }
        }
        public static void Close(DMMInterface DMM1, bool IsNewMultimeter)
        {
            if(IsNewMultimeter)
            {
                Agilent_34465A_LIB.Agilent_34465A_LIB.Close(DMM1);
            }
            else
            {
                Agilent_34411A_LIB.Agilent_34411A_LIB.Close(DMM1);
            }
        }
        public static void GetData(DMMInterface DMM1, bool IsNewMultimeter,out double? Data)
        {
            if(IsNewMultimeter)
            {
                Agilent_34465A_LIB.Agilent_34465A_LIB.GetData(DMM1,out Data);
            }
            else
            {
                Agilent_34411A_LIB.Agilent_34411A_LIB.GetData(DMM1, out Data);
            }
        }
        public static Error getError(string title, DMMInterface driver1,bool IsNewMultimeter)
        {
            if(IsNewMultimeter)
            {
                return Agilent_34465A_LIB.Agilent_34465A_LIB.getError(title, driver1);
            }
            else
            {
                return Agilent_34411A_LIB.Agilent_34411A_LIB.getError(title, driver1);
            }
        }
    }

}
