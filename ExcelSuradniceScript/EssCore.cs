using System;
using System.IO;
using OfficeOpenXml;

namespace ExcelSuradniceScript
{
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

            
            Console.WriteLine("Načítavám súbor.");
            _excelPackage = new ExcelPackage(file);
            Console.WriteLine("Súbor úspešne načítaný.");
            Console.WriteLine("Počkaj na ďalší pokyn.");
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

        private string[] SelectCols()
        {
            string stringCols;
            do
            {
                Console.WriteLine("Zadaj stlpce na fomatovanie (oddelene ciarkou): ");
                stringCols = Console.ReadLine(); 
            } while (stringCols == null);
            stringCols = stringCols.Replace(" ", "").ToUpper();
            
            return stringCols.Split(",");
        }

        private int SelectStartNumber()
        {
            int number;
            string input;
            do
            {
                Console.WriteLine("Zadaj startovacie cislo riadka: ");
                input = Console.ReadLine();
            } while (!int.TryParse(input, out number));
            
            return number;
        }
        
        private int SelectEndNumber()
        {
            int number;
            string input;
            do
            {
                Console.WriteLine("Zadaj konciace cislo riadka: ");
                input = Console.ReadLine();
            } while (!int.TryParse(input, out number));
            
            return number;
        }

        private void SaveAs()
        {
            String pathToFile = EssFileManager.SaveFile();
            if (pathToFile == null)
            {
                Console.WriteLine("Chyba pri ulkadaní súboru.");
                return;
            }

            var file = new FileInfo(pathToFile);
            _excelPackage.SaveAs(file);
        }

        public void Start()
        {
            var cols = SelectCols();
            var start = SelectStartNumber();
            var end = SelectEndNumber();
            
            var fm = new EssFormatManager(_mainSheet, start, end);
            fm.FormatColumns(cols);

            Console.WriteLine("Ukladám súbor.");
            SaveAs();
        }
    }
}