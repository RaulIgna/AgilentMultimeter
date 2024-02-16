using Agilent.Agilent34410.Interop;
using AutoTest;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AgilentMultimeter;

namespace Agilent_34411A_LIB
{

    public class DMM34410A : DMMInterface
    {
        public string ResourceName { get; set; }
        public string ID { get; set; }
        public double Range { get; set; }
        public double NPLC { get; set; }
        public bool InputImpedance { get; set; }
        public Agilent34410AutoZeroEnum AutoZero;
        public bool NullState { get; set; }
        public double NullValue { get; set; }
        public double Constant { get; set; }
        public double RawValue { get; set; }
        public bool Run { get; set; }

        public Thread Work { get; set; }

        public Agilent34410 Driver = new Agilent34410();

        public DMM34410A(string lID)
        {
            ResourceName = "USB0::0x0957::0x0607::MY5301" + lID + "::0::INSTR";
            ID = lID;
            Range = 100;
            NPLC = 10;
        }

        public void DMM_SetID(string text)
        {
            ID = text;
            ResourceName = "USB0::0x0957::0x0607::MY5301" + ID + "::0::INSTR";
        }

        public void DMM_SelectAutoRange(int AutoZeroID)
        {
            switch (AutoZeroID)
            {
                case 0:
                    AutoZero = Agilent34410AutoZeroEnum.Agilent34410AutoZeroOff;
                    break;
                case 1:
                    AutoZero = Agilent34410AutoZeroEnum.Agilent34410AutoZeroOnce;
                    break;
                case 2:
                    AutoZero = Agilent34410AutoZeroEnum.Agilent34410AutoZeroOn;
                    break;
                default:
                    AutoZero = Agilent34410AutoZeroEnum.Agilent34410AutoZeroOnce;
                    break;

            }

        }


        public void DMM_SelectRange(int SelectedIndex)
        {
            switch (SelectedIndex)
            {
                case 0:
                    Range = 0.1;
                    break;
                case 1:
                    Range = 1;
                    break;
                case 2:
                    Range = 10;
                    break;
                case 3:
                    Range = 100;
                    break;
                case 4:
                    Range = 1000;
                    break;
                default:
                    Range = 100;
                    break;
            }
        }

        public void DMM_SelectNPLC(int SelectedIndex)
        {
            switch (SelectedIndex)
            {
                case 0:
                    NPLC = 0.006;
                    break;
                case 1:
                    NPLC = 0.02;
                    break;
                case 2:
                    NPLC = 0.06;
                    break;
                case 3:
                    NPLC = 0.2;
                    break;
                case 4:
                    NPLC = 1;
                    break;
                case 5:
                    NPLC = 2;
                    break;
                case 6:
                    NPLC = 10;
                    break;
                case 7:
                    NPLC = 100;
                    break;
                default:
                    NPLC = 10;
                    break;
            }
        }
    }
    public class Agilent_34411A_LIB 
    {
        static Error err_rx = new Error(false, "Default Rx thread error");//will be set by Rx thread if OK

        public static Error Open(DMMInterface InterfaceDMM)
        {
            var DMM = (DMM34410A)InterfaceDMM;
            // Setup IVI-defined initialization options
            //OptionString
            //The OptionString allows the user to pass optional settings to the driver. These settings override any settings that are specified in the IVI Configuration Store. If the IVI Configuration Store is not used (a resource descriptor is passed as the ResourceName instead of a logical name) then any setting that is not specified has a default value as specified by IVI. 

            //Console.WriteLine("Test");





            string pOptionString =
                "Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=true, RecordCoercions=false, Simulate=false";

            //IdQuery
            //If this is enabled, the driver will query the instrument model and compare it with a list of instrument models that is supported by the driver. If the model is not supported, Initialize will fail with the E_IVI_ID_QUERY_FAILED error code. 
            bool pIdQuery = true;

            //Reset
            //If this is enabled, the driver will perform a reset of the instrument. If the reset fails, Initialize will fail with the E_IVI_RESET_FAILED error code.
            bool pReset = true;


            try// to close previous oppened
            {
                DMM.Driver.Close();
                DMM.Driver = null;
            }
            catch
            {

            }

            DMM.Driver = new Agilent34410();//driver class Interface




            // Open the I/O session with the driver
            //Dmm.Initialize(addr, true, true, "");*/
            Error er = new Error();
            try
            {
                DMM.Driver.Initialize(DMM.ResourceName, pIdQuery, pReset, pOptionString);
                er = getError("Initialize", DMM);
                if (!er.OK)
                {
                    return er;
                }

                // Configure the 34410A/11A for voltage measurement, using 10 V range


                // double range = 100;//10V// The expected signal value. If the range value specified is negative, autoranging will be enabled. 

                //Least (4.5 digits), for faster speed// Measurement resolution. Select from Least (4.5 digits), Default (5.5 digits) and Best (6.5 digits). Higher resolutions (6.5 digits) result in slower measurement speeds. 
                Agilent34410ResolutionEnum resolution = Agilent34410ResolutionEnum.Agilent34410ResolutionLeast;
                DMM.Driver.Voltage.DCVoltage.Configure(DMM.Range, resolution);

                er = getError("Configure(range, resolution)", DMM);
                if (!er.OK)
                {
                    return er;
                }

                /* Null State */
                DMM.Driver.Voltage.DCVoltage.NullState = DMM.NullState;
                er = getError("Null State", DMM);
                if (!er.OK)
                {
                    return er;
                }

                /* Null Value */
                DMM.Driver.Voltage.DCVoltage.NullValue = DMM.NullValue;
                er = getError("Null Value", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.Voltage.DCVoltage.AutoImpedance = DMM.InputImpedance;
                er = getError("Input Impedance", DMM);
                if (!er.OK)
                {
                    return er;
                }

                //Enables or disables the autozero mode for DC voltage measurements
                //Agilent34410AutoZeroOnce=Issues an immediate zero measurement, and then turns autozero off.
                DMM.Driver.Voltage.DCVoltage.AutoZero = DMM.AutoZero;
                er = getError("AutoZero", DMM);
                if (!er.OK)
                {
                    return er;
                }


                //Sets or gets the integration time in seconds (called aperture time) for DC voltage measurements.
                // Set aperture to 100usec=100E-6 (34411A can be set to 20usec)
                DMM.Driver.Voltage.DCVoltage.NPLC = DMM.NPLC;//(100E-6) / 2;//10ms=10E-3
                er = getError("DCVoltage.Aperture", DMM);
                if (!er.OK)
                {
                    return er;
                }


                // Set up triggering for 1000 samples from a single trigger event
                //Sets or gets the trigger source for measurements.
                //may be set on bus or external
                DMM.Driver.Trigger.TriggerSource = Agilent34410TriggerSourceEnum.Agilent34410TriggerSourceImmediate;
                er = getError(".Trigger.TriggerSource", DMM);
                if (!er.OK)
                {
                    return er;
                }
                //Gets or sets the number of triggers that will be accepted by the meter before returning to the 'idle' trigger state.
                DMM.Driver.Trigger.TriggerCount = 1;//one trig for aq
                er = getError("", DMM);
                if (!er.OK)
                {
                    return er;
                }


                //Gets or sets the delay between the trigger signal and the first measurement.
                //This may be useful in applications where you want to allow the input to settle before taking a reading or for pacing a burst of readings. The programmed trigger delay overrides the default trigger delay that the instrument automatically adds.
                DMM.Driver.Trigger.TriggerDelay = 0;
                er = getError("Trigger.TriggerDelay ", DMM);
                if (!er.OK)
                {
                    return er;
                }
                //Gets or sets the number of readings (samples) the meter will take per trigger.
                DMM.Driver.Trigger.SampleCount = 1;// 1000;
                er = getError("Trigger.SampleCount", DMM);
                if (!er.OK)
                {
                    return er;
                }
                //Gets or sets a sample interval for timed sampling when the sample count is greater than one.
                DMM.Driver.Trigger.SampleInterval = 0.00010;//doesn't meter for Count=1 sample
                er = getError("Trigger.SampleInterval", DMM);
                if (!er.OK)
                {
                    return er;
                }

                // Set up data format for binary transfer 64-bit transfer
                //The data format can be either ASCII or REAL. If ASCII is specified, numeric data is transferred as ASCII characters. The numbers are separated by commas as specified in IEEE 488.2. If REAL is specified, numeric data is transferred as REAL binary data in IEEE 488.2 definite-length block format.
                DMM.Driver.DataFormat.DataFormat = Agilent34410DataFormatEnum.Agilent34410DataFormatReal64;
                er = getError("DataFormat.DataFormat", DMM);
                if (!er.OK)
                {
                    return er;
                }







            }
            catch (Exception ex)
            {
                DMM.Run = false;
                return new Error(ex.Message);
            }





            return new Error(true, "No Error");

        }


        public static void Close(DMMInterface InterfaceDMM)
        {
            var DMM = (DMM34410A)InterfaceDMM;
            try// to close previous oppened
            {
                DMM.Driver.Close();
                DMM.Driver = null;
            }
            catch
            {

            }

        }

        public static void GetData(DMMInterface InterfaceDMM, out double[] Data)
        {
            var DMM = (DMM34410A)InterfaceDMM;
            int dataPts;

            // Initiate the measurement
            //Changes the state of the triggering system from the 'idle' state to the 'wait-for-trigger' state.
            //Measurements will begin when the specified trigger conditions are satisfied following execution of this method. Note that this method also clears the previous set of readings from memory.
            DMM.Driver.Measurement.Initiate();

            Thread.Sleep((int)(DMM.NPLC * 100));//this will slow down
                                                //Indicates whether readings should be taken from reading memory or non-volatile memory.
                                                //Gets the total number of readings currently stored in reading memory (aka volatile memo) or non-volatile memory
                                                //Agilent34410MemoryTypeReadingMemory= Reading from volatile memory. 
            dataPts = DMM.Driver.Measurement.get_ReadingCount(Agilent34410MemoryTypeEnum.Agilent34410MemoryTypeReadingMemory);

            if (dataPts > 0)//if We have data points, lets read them out
            {
                //The readings are erased from memory starting with the oldest reading first. The purpose of this method is to allow you to periodically remove readings from memory during a series of measurements to avoid a reading memory overflow. If the specified number of readings are not yet present in memory when this routine is called, an instrument error will result and the data fetch will fail.
                Data = DMM.Driver.Measurement.RemoveReadings(dataPts);
            }
            else
            {
                Data = null;
            }

        }

        public static Error getError(string title, DMMInterface InterfaceDriver)
        {
            var Driver = ((DMM34410A)InterfaceDriver).Driver;

            int err_code = 0;
            bool checkstat = false;
            string err_msg = "";

            try
            {
                // check for initial error
                Driver.Utility.ErrorQuery(ref err_code, ref err_msg);

                while (err_code != 0)
                {
                    checkstat = true;

                    title += "Error in: " + title + "\n";
                    title += title + "Error Number: " + err_code.ToString() + "\nError Message: " + err_msg;


                    Driver.Utility.ErrorQuery(ref err_code, ref err_msg);
                }

                // Exit program if error is determined
                if (checkstat)
                {
                    title += Environment.NewLine + "Idriver Closed in internal error";
                    Driver.Close();
                    Driver = null;
                    return new Error(title);

                    //Environment.Exit(0); //after message box
                }
            }
            catch (Exception e)
            {
                title += Environment.NewLine + e.Message + "Idriver Closed in exception";
                Driver.Close();
                Driver = null;
                return new Error(title);

                //Environment.Exit(0); //after message box
            }

            return new Error(true, "No error");
        }

    }
}
