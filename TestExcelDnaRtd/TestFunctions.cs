using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using System.Threading;

namespace TestExcelDnaRtd
{
    public static class TestFunctions
    {
        [ExcelFunction(Description="Get the real-time value of the spot", HelpTopic="Need help?", Category="TEST")]
        public static object GetSpot([ExcelArgument("ISIN code")] string isin)
        {
            //ExcelDna.Logging.LogDisplay.WriteLine("Function :{0} {1}", typeof(TestRtdServer).FullName, isin);
            return XlCall.RTD(typeof(TestRtdServer).FullName, null, isin);
        }
    }
}
