using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestExcelComRtd
{
    public class SpotUpdateEventArgs : EventArgs
    {
        public double Value { get; private set; }

        public string ISIN { get; private set; }

        public SpotUpdateEventArgs(string isin, double value)
        {
            this.ISIN = isin;
            this.Value = value;
        }
    }
}
