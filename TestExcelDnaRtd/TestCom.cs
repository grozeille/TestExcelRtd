using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace TestExcelDnaRtd
{
    [Guid("b02ada99-00ce-4a73-a57c-fba4d66aac99")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITestCom))]
    public class TestCom : ITestCom
    {
        public double GetSpot(string isin)
        {
            return 42;
        }
    }
}
