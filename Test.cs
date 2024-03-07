using Agilent_34411A_LIB;
using Agilent_34465A_LIB;
using AutoTest;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AgilentMultimeter
{
    public static class HP34410A_test
    {
        private static Error err_rx = new Error(false, "Default Rx thread error");
        public static Error MeasureStart(DMMInterface DMM, bool NewMultimeter)
        {

            Error error = new Error("TransactKline: Nothing Received from ECU");
            DMM.Run = true;
            error = AgilentInterface.Open(DMM, NewMultimeter);
            if(!error.OK)
            {
                return error;
            }
            DMM.Work = new Thread((ThreadStart)delegate
            {
                DoWork_Rx(DMM,NewMultimeter);
            });
            DMM.Work.Start();
            return error;
        }

        public static Error Stop(DMMInterface DMM)
        {
            try
            {
                DMM.Run = false;
                DMM.Work.Join();
                DMM.Work = null;
            }
            catch (Exception ex)
            {
                return new Error("Cannot stop: " + ex.Message);
            }

            return new Error(true, "no error");
        }

        private static void DoWork_Rx(DMMInterface DMM, bool NewMultimeter)
        {
            Error error = new Error(true, "no error");
            bool flag = false;
            err_rx = new Error(true, "no error");
            try
            {
                while (DMM.Run)
                {
                    double? Data = null;
                    //if (NewMultimeter)
                    //{
                    //    global::Agilent_34465A_LIB.Agilent_34465A_LIB.GetData(DMM,out Data);
                    //}
                    //else
                    //{
                    //    global::Agilent_34411A_LIB.Agilent_34411A_LIB.GetData(DMM,out Data);
                    //}
                    AgilentInterface.GetData(DMM, NewMultimeter, out Data );
                    try
                    {
                        DMM.RawValue = (double)Data;
                    }
                    catch(Exception ex)
                    {
                        // TODO: Add error handling
                        Console.WriteLine(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                flag = true;
                err_rx = new Error("Receiver thread error :" + ex.Message + getError(" ", DMM,NewMultimeter));
            }

            if (!flag)
            {
                err_rx = new Error(true, "No error");
            }
        }

        private static Error getError(string title, DMMInterface InterfaceDMM, bool NewMultimeter)
        {
            if (NewMultimeter)
            {
                var DMM = (DMM34465A)InterfaceDMM;
                int num = 0;
                bool flag = false;
                string text = "";
                try
                {
                    DMM.Driver.Utility.ErrorQuery(ref num, ref text);
                    while (num != 0)
                    {
                        flag = true;
                        title = title + "Error in: " + title + "\n";
                        title = title + title + "Error Number: " + num + "\nError Message: " + text;
                        DMM.Driver.Utility.ErrorQuery(ref num, ref text);
                    }

                    if (flag)
                    {
                        title = title + Environment.NewLine + "Idriver Closed in internal error";
                        DMM.Driver.Close();
                        DMM.Driver = null;
                        return new Error(title);
                    }
                }
                catch (Exception ex)
                {
                    title = title + Environment.NewLine + ex.Message + "Idriver Closed in exception";
                    if (DMM.Driver != null)
                    {
                        DMM.Driver.Close();
                        DMM.Driver = null;
                        return new Error(title);
                    }

                    return new Error("Driver not opened");
                }

                return new Error(true, "No error");
            }
            else
            {
                var DMM = (DMM34410A)InterfaceDMM;
                int num = 0;
                bool flag = false;
                string text = "";
                try
                {
                    DMM.Driver.Utility.ErrorQuery(ref num, ref text);
                    while (num != 0)
                    {
                        flag = true;
                        title = title + "Error in: " + title + "\n";
                        title = title + title + "Error Number: " + num + "\nError Message: " + text;
                        DMM.Driver.Utility.ErrorQuery(ref num, ref text);
                    }

                    if (flag)
                    {
                        title = title + Environment.NewLine + "Idriver Closed in internal error";
                        DMM.Driver.Close();
                        DMM.Driver = null;
                        return new Error(title);
                    }
                }
                catch (Exception ex)
                {
                    title = title + Environment.NewLine + ex.Message + "Idriver Closed in exception";
                    if (DMM.Driver != null)
                    {
                        DMM.Driver.Close();
                        DMM.Driver = null;
                        return new Error(title);
                    }

                    return new Error("Driver not opened");
                }

                return new Error(true, "No error");
            }
        }
    };
}
