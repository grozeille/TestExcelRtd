using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace TestExcelComRtd
{
    [Guid("21cb87e1-b85a-49b6-906c-94e4d234a28d")]
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
