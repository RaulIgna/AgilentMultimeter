using Agilent.Ag3446x.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutoTest;
using AgilentMultimeter;
using Agilent.Agilent34410.Interop;
using Ivi.Driver;

namespace Agilent_34465A_LIB
{
    public class DMM34465A : DMMInterface
    {
        public string ResourceName { get; set; }
        public string ID { get; set; }
        public double Range { get; set; }
        public double NPLC { get; set; }
        public bool InputImpedance { get; set; }
        public Ag3446xAutoZeroEnum AutoZero { get; set; }
        public bool NullState { get; set; }
        public double NullValue { get; set; }
        public double Constant { get; set; }
        public double RawValue { get; set; }
        public bool Run { get; set; }
        public string RawId { get; set; }
        public Thread Work { get; set; }
        public Agilent.Ag3446x.Interop.Ag3446x Driver { get; set; }

        public DMM34465A(string lID)
        {
            ResourceName = "USB0::0x2A8D::0x0101::MY6009" + lID + "::INSTR";
            RawId = lID;
                       //ResourceName = "USB0::0x2A8D::0x0101::MY6009" + lID + "::0::INSTR";
            ID = lID;
            Range = 100;
            NPLC = 10;
        }

        public void DMM_SetID(string text)
        {
            ID = text;
            ResourceName = "USB0::0x2A8D::0x0101::MY6009" + ID + "::INSTR";
        }

        public void DMM_SelectAutoRange(int AutoZeroID)
        {
            switch (AutoZeroID)
            {
                case 0:
                    AutoZero = Ag3446xAutoZeroEnum.Ag3446xAutoZeroOff;
                    break;
                case 1:
                    AutoZero = Ag3446xAutoZeroEnum.Ag3446xAutoZeroOnce;
                    break;
                case 2:
                    AutoZero = Ag3446xAutoZeroEnum.Ag3446xAutoZeroOn;
                    break;
                default:
                    AutoZero = Ag3446xAutoZeroEnum.Ag3446xAutoZeroOnce;
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




    public class Agilent_34465A_LIB
    {
        private static Error err_rx = new Error(false, "Default Rx thread error");
        public static Error Open(DMMInterface DMM1)
        {
            var DMM = (DMM34465A)DMM1;
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
            catch { }

            Error er = new Error();
            try
            {
                DMM.Driver = new Agilent.Ag3446x.Interop.Ag3446x();//(DMM.ResourceName, pIdQuery, pReset, pOptionString);
                try
                {
                    DMM.Driver.Initialize(DMM.ResourceName, pIdQuery, pReset, pOptionString);
                }
                catch {
                    Console.WriteLine("Am ajuns aici");
                }
                er = getError("Initialize", DMM);
                if (!er.OK)
                {
                    // Try the other one
                    DMM.ResourceName = "USB0::0x2A8D::0x0101::MY6069" + DMM.RawId + "::INSTR";
                    DMM.Driver.Initialize(DMM.ResourceName, pIdQuery, pReset, pOptionString);
                    er = getError("Initialize", DMM);
                    if(!er.OK)
                    {
                        return er;
                    }
                    //return er;

                }

                //Resolution resolution = Resolution.Max;
                DMM.Driver.DCVoltage.Configure(DMM.Range, 4.5);

                er = getError("Configure(range, resolution)", DMM);
                if (!er.OK)
                {
                    return er;
                }


                DMM.Driver.DCVoltage.NullEnabled = DMM.NullState;

                er = getError("Null State", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.DCVoltage.NullValue = DMM.NullValue;

                er = getError("Null Value", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.DCVoltage.ImpedanceAutoEnabled = DMM.InputImpedance;

                er = getError("Input Impedance", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.DCVoltage.AutoZero = DMM.AutoZero;

                er = getError("AutoZero", DMM);
                if (!er.OK)
                {
                    return er;
                }


                DMM.Driver.DCVoltage.NPLC = DMM.NPLC;
                er = getError("DCVoltage.Aperture", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.Trigger.Source = Ag3446xTriggerSourceEnum.Ag3446xTriggerSourceImmediate;
                er = getError("Trigger.TriggerSource", DMM);
                if (!er.OK)
                {
                    return er;
                }



                DMM.Driver.Trigger.Count = 2;
                er = getError("", DMM);

                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.Trigger.Delay = 0; //PrecisionTimeSpan.Zero;
                er = getError("Trigger.TriggerDelay ", DMM);
                if (!er.OK)
                {
                    return er;
                }

                DMM.Driver.Trigger.SampleCount = 1;
                er = getError("Trigger.SampleCount", DMM);
                if (!er.OK)
                {
                    return er;
                }


                DMM.Driver.Display.Mode = Ag3446xDisplayModeEnum.Ag3446xDisplayModeNumeric;
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

            return new Error(true, "No error");
        }
        public static void Close(DMMInterface DMM1)
        {
            var DMM = (DMM34465A)DMM1;
            try
            {
                DMM.Driver.Close();
                DMM.Driver = null;
            }
            catch
            {

            }
        }


        public static void GetData(DMMInterface DMM1, out double? Data)
        {
            var DMM = (DMM34465A)DMM1;
            int dataPts;

            // Initiate the measurement
            //Changes the state of the triggering system from the 'idle' state to the 'wait-for-trigger' state.
            //Measurements will begin when the specified trigger conditions are satisfied following execution of this method. Note that this method also clears the previous set of readings from memory.
           //DMM.Driver.Measurement.Initiate();


            DMM.Driver.Measurement.Read((int)DMM.NPLC * 1000);

            // Slow down
            Thread.Sleep((int)(DMM.NPLC * 1000));



           // dataPts = DMM.Driver.Measurement.get_ReadingCount(Agilent34410MemoryTypeEnum.Agilent34410MemoryTypeReadingMemory);

            // Gets the total number of reading currently stored in reading memory
            dataPts = DMM.Driver.Measurement.ReadingCount;

            // If there is any data, read and remove the data
            // Otherwise, set Data to null
            if(dataPts > 0)
            {
                Data = DMM.Driver.Measurement.RemoveReadings(dataPts)[0];
            }
            else
            {
                Data = null;
            }
             

        }


        public static Error getError(string title, DMMInterface driver1)
        {
            bool checkstat = false;
            var driver = (DMM34465A)driver1;
            try
            {
                // check for initial error
                int error = 0;
                string errorString = "";
                driver.Driver.Utility.ErrorQuery(error, errorString);

                //driver.Driver.Utility.ErrorQuery();

                while (error != 0)
                {
                    checkstat = true;

                    title += "Error in: " + title + "\n";
                    title += title + "Error Number: " + error + "\nError Message: " + errorString;

                    driver.Driver.Utility.ErrorQuery(error, errorString);
                    // res = driver.Driver.Utility.ErrorQuery();
                }

                // Exit program if error is determined
                if (checkstat)
                {
                    title += Environment.NewLine + "Idriver Closed in internal error";
                    driver.Driver.Close();
                    driver = null;
                    return new Error(title);

                    //Environment.Exit(0); //after message box
                }
            }
            catch (Exception e)
            {
                title += Environment.NewLine + e.Message + "Idriver Closed in exception";
                driver.Driver.Close();
                driver = null;
                return new Error(title);

                //Environment.Exit(0); //after message box
            }

            return new Error(true, "No error");
        }
    }



}
