using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Agilent.Ag3446x.Interop;
using Agilent.AgAPS.Interop;
using AutoTest;

namespace AgilentMultimeter
{
    public class RP795A
    {
        public string ResourceName { get; set; }
        public string ID { get; set; }
        public bool Run { get; set; }
        public double OVPLimit { get; set; }
        public double VoltageLimit { get; set; }
        public double VoltageLevel { get; set; }
        public double CurrentLevel { get; set; }

        public Agilent.AgAPS.Interop.AgAPS Driver { get; set; }
    
        public RP795A(string lID)
        {
            ResourceName = "USB0::0x2A8D::0x2802::MY6300" + lID + "::0::INSTR";
            ID = lID;

        }
    }

    public class Keysight_7945A_LIB
    {
        private static Error err_rt = new Error(false, "Default Rx thread error");

        public static Error Open(RP795A RP)
        {

            string pOptionString = "Cache=false, InterchangeCheck=false, QueryInstrStatus=true";
            bool pIdQuery = true;
            bool pReset = true;

            try
            {
                RP.Driver.Close();
                RP.Driver = null;
            }
            catch { }

            Error er = new Error();
            try
            {
                RP.Driver = new Agilent.AgAPS.Interop.AgAPS();
                RP.Driver.Initialize(RP.ResourceName,pIdQuery,pReset,pOptionString);

                RP.Driver.Output.Voltage.Level = RP.VoltageLevel;
                RP.Driver.Output.Current.Level = RP.CurrentLevel;


            }
            catch { }


            return new Error(true, "No error");
        }
    }
}
