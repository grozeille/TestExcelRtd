using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;

namespace TestExcelComRtd
{
    [Guid("d8ec8db8-51d1-4ed5-8ccb-ae8a815a4e69")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITestFunctions))]
    public class TestFunctions : ITestFunctions 
    {
        [ComRegisterFunctionAttribute] //required to tag a C# class a UDF
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type));
        }

        [ComUnregisterFunctionAttribute] //required to tag a C# class a UDF
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type), false);
        }

        private static string GetSubKeyName(Type type) //required to tag a C# class a UDF
        {
            string s = @"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable";
            return s;
        }

        public object GetSpot(string isin)
        {
            Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");

            try
            {
                return excelApp.WorksheetFunction.RTD(typeof(TestRtdServer).FullName, null, isin);
            }
            finally
            {
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }
        }
    }
}
