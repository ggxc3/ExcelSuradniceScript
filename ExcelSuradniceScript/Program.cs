using System;
using System.IO;
using System.Windows.Forms;
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

            Console.ReadKey();
        }
        
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }

    class EssFileManager
    {
        public static string LoadFile()
        {
            var dialog = new OpenFileDialog
            {
                RestoreDirectory = true,
                Title = "Vyber Excel súbor",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                
                CheckFileExists = true,
                CheckPathExists = true
                
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;
        }
    }

    class EssCore
    {
        private ExcelPackage _excelPackage;
        private ExcelWorksheet _mainSheet;

        public EssCore()
        {
            String pathToFile = EssFileManager.LoadFile();
            if (pathToFile == null)
            {
                Console.WriteLine("Chyba pri načítavaní súboru.");
                return;
            }

            var file = new FileInfo(pathToFile);

            _excelPackage = new ExcelPackage(file);
            _mainSheet = _excelPackage.Workbook.Worksheets[SelectSheet() - 1];
        }

        private int SelectSheet()
        {
            int number = 1;
            foreach (var sheet in _excelPackage.Workbook.Worksheets)
            {
                Console.WriteLine($"{number}. {sheet}");
                number++;
            }
            
            int selected;
            do
            {
                if (!int.TryParse(Console.ReadLine(), out selected))
                {
                    Console.WriteLine("Zadaj cislo: ");
                }
                else
                {
                    if (selected >= 1 && selected <= _excelPackage.Workbook.Worksheets.Count)
                    {
                        break;
                    }
                    Console.WriteLine("Zadaj spravne cislo z rozsahu: ");
                }
            } while (true);

            return selected;
        }

        public void Start()
        {
        }
    }
}