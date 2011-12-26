using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace TestExcelComRtd
{
    [Guid("805de000-acb6-482f-8de8-bfbe78e1e7b2")]
    [ComVisible(true)]
    public interface ITestCom
    {
        double GetSpot(string isin);
    }
}
