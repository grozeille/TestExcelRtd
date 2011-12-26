using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace TestExcelComRtd
{
    [ComVisible(true)]
    [Guid("11edd7bb-3334-4461-8876-fff72f573a8d")]
    public interface ITestFunctions
    {
        object GetSpot(string isin);
    }
}
