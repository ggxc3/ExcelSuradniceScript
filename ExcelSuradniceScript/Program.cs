using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ExcelSuradniceScript
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            AllocConsole();

            var dialog = new OpenFileDialog();
            dialog.ShowDialog();
        }
        
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}