using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTest
{
    public class Error
    {
        private bool oK = false;
        private string serror = (string)null;

        public Error()
        {
            this.oK = false;
            this.serror = "Default error";
        }
        public Error(string sError)
        {
            oK = false;
            serror = sError;
        }
        public Error(bool bOk, string sError)
        {
            this.oK = bOk;
            this.serror = sError;
        }

        public bool OK
        {
            get => this.oK;
            set => this.oK = value;
        }

        public string Str_Error
        {
            set => this.Set_Error_String(value);
            get => this.serror;
        }

        public void Set_Error_String(string sError)
        {
            lock (this)
            {
                this.oK = false;
                this.serror = sError;
            }
        }
    }
}
