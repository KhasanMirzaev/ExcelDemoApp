using System;
using OfficeOpenXml;

namespace ExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
    }    
}
