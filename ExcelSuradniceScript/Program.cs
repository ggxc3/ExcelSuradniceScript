using System;
using System.Runtime.InteropServices;
using OfficeOpenXml;

namespace ExcelSuradniceScript
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            AllocConsole();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var core = new EssCore();
            try
            {
                core.Start();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.WriteLine("Pre zatvorenie konzoly stlačte ľubovoľné tlačidlo.");
            Console.ReadKey();
        }
        
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}