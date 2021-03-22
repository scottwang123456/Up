using OfficeOpenXml;
using System;
using System.Windows.Forms;

namespace Up
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Application.Run(new Form1());
        }
    }
}
