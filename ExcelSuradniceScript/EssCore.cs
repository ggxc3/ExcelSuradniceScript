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