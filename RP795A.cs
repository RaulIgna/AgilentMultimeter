using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Agilent.Ag3446x.Interop;
using Agilent.AgAPS.Interop;
using AutoTest;
using Ivi.Driver;
using System.IO;
using Ivi.Visa;
using static System.Net.Mime.MediaTypeNames;

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
        public AgAPSRegulationModeEnum RegulationMode { get; set; }
        public Agilent.AgAPS.Interop.AgAPS Driver { get; set; }

        public RP795A(string lID)
        {
            ResourceName = "USB0::0x2A8D::0x2802::MY6300" + lID + "::0::INSTR";
            ID = lID;
            RegulationMode = AgAPSRegulationModeEnum.AgAPSRegulationModeVoltageSource;
            VoltageLevel = 5;
            CurrentLevel = 0.1;
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
                RP.Driver.Initialize(RP.ResourceName, pIdQuery, pReset, pOptionString);

                RP.Driver.Output.RegulationMode = RP.RegulationMode;

                RP.Driver.Output.Voltage.Level = RP.VoltageLevel;
                RP.Driver.Output.Current.Level = RP.CurrentLevel;

                RP.Driver.Output.Enabled = true;
                RP.Driver.System.WaitForOperationComplete(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }


            return new Error(true, "No error");
        }

        // Gets the Voltage and Current that is supplied
        static public void GetMeasurements(RP795A RP, out double Voltage, out double Current)
        {
            Voltage = RP.Driver.Measurement.Measure(AgAPSMeasurementTypeEnum.AgAPSMeasurementTypeVoltage, 1000);
            Current = RP.Driver.Measurement.Fetch(AgAPSFetchTypeEnum.AgAPSFetchTypeCurrent, 1000);
        }

        // Gets the value of the set voltage and current, not the one that is provided
        static public void GetSetValues(RP795A RP, out double Voltage, out double Current)
        {
            Voltage = RP.Driver.Output.Voltage.Level;
            Current = RP.Driver.Output.Current.Level;
        }

        // Increments the voltage from V0, to V1, in an interval of Time milliseconds
        static public void SetRampVoltage(RP795A RP, double V0, double V1, double Time)
        {
            if(RP.Driver.Output.Voltage.Level - V0 > 0.01)
            {
                // Make a smooth transition from V0 
                SetRampVoltage(RP, RP.Driver.Output.Voltage.Level, V0, 200);
            }
            RP.Driver.Output.Voltage.Level = V0;
            int IterationNumber = (int)(Time / 100); // Number of iteration ( time / how long 1 iteration should last)
            double IterationValue = (V1 - V0) / IterationNumber;
            int Iterations = 0;

            var Timer = new System.Timers.Timer();
            Timer.Elapsed += (sender, e) =>
            {
                RP.Driver.Output.Voltage.Level += IterationValue;
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, "test.txt");
               var errors =  CheckForErrors(RP);
                foreach(var error in errors ) { File.WriteAllText(filePath, error); }
                
                Iterations++;

                if (Iterations == IterationNumber)
                {
                    Timer.Dispose();

                }
            };
            Timer.Interval = 100;
            Timer.Start();
        }

        static public void SetRampCurrent(RP795A RP, double V0, double V1, double Time)
        {
            if (RP.Driver.Output.Current.Level - V0 > 0.01)
            {
                // Make a smooth transition from V0 
                SetRampCurrent(RP, RP.Driver.Output.Current.Level, V0, 200);
            }
            RP.Driver.Output.Current.Level = V0;
            int IterationNumber = (int)(Time / 100); // Number of iteration ( time / how long 1 iteration should last)
            double IterationValue = (V1 - V0) / IterationNumber;
            int Iterations = 0;

            var Timer = new System.Timers.Timer();
            Timer.Elapsed += (sender, e) =>
            {
                RP.Driver.Output.Current.Level += IterationValue;
                Iterations++;

                if (Iterations == IterationNumber)
                {
                    Timer.Dispose();
                }
            };
            Timer.Interval = 100;
            Timer.Start();
        }

        static public  List<string> CheckForErrors(RP795A RP)
        {
            int errorNum = -1;
            string errorMsg = null;
            List<string> ret = new List<string>();
            while (errorNum != 0)
            {
                RP.Driver.Utility.ErrorQuery(ref errorNum, ref errorMsg);
                ret.Add(errorMsg);
            }
            return ret;
        }

        static public void SetVoltage(RP795A RP, double Voltage)
        {
            RP.Driver.Output.Voltage.Level = Voltage;
        }

        static public void SetCurrent(RP795A RP, double Current)
        {
            RP.Driver.Output.Current.Level = Current;
        }

        static public void CloseDriver(RP795A RP)
        {
            RP.Driver.Close();
            RP.Driver = null;
        }

        public static void SetVoltageLimit(RP795A RP, double Voltage)
        {
            RP.Driver.Output.Voltage.PositiveLimit = Voltage;
        }
        public static void SetCurrentLimitMax(RP795A RP, double Current)
        {
            RP.Driver.Output.Current.PositiveLimit = Current;
        }

        public static void SetCurrentLimitMin(RP795A RP, double Current)
        {
            RP.Driver.Output.Current.NegativeLimit = Current;
        }

        public static void GetConnectedDevices(out string[] Devices)
        {
          //  IEnumerable<string> enums = GlobalResourceManager.Find("USB0::10893::10242?*INSTR");
            var enums = GlobalResourceManager.Find("USB0::10893::10242?*INSTR");
            Devices = new string[enums.Count()];
            int i = 0;
            foreach (var item in enums)
            {
               
                Devices[i] = item.Substring(item.IndexOf("::0") - 4, 4);
                i++;
            }
        }
    }
}
