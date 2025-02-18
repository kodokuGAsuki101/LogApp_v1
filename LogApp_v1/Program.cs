using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.XlsIO;
using System.Data;
using System.IO;
using System.Reflection;
using System.Globalization;

namespace LogApp_v1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            CultureInfo ci = CultureInfo.CreateSpecificCulture("en-US"); // or any other culture
            ci.DateTimeFormat.ShortDatePattern = "MM/dd/yyyy";
            ci.DateTimeFormat.LongDatePattern = "MM/dd/yyyy HH:mm:ss tt";
            CultureInfo.DefaultThreadCurrentCulture = ci;
            CultureInfo.DefaultThreadCurrentUICulture = ci;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
