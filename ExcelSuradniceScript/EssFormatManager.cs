using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace ExcelSuradniceScript
{
    class EssFormatManager
    {
        private ExcelWorksheet _sheet;
        private int _start;
        private int _end;

        public EssFormatManager(ExcelWorksheet sheet, int start, int end)
        {
            _sheet = sheet;
            _start = start;
            _end = end;
        }

        public void FormatColumns(string[] columns)
        {
            foreach (var col in columns)
            {
                FormatColumn(col);
            }
        }
        
        private void FormatColumn(string column)
        {
            var deleteRows = new List<int[]>();
            bool falsing = false;
            for (int i = _start; i <= _end; i++)
            {
                Console.WriteLine($"Riadok: {i}");
                if (!FormatCell(column, i))
                {
                    if (!falsing)
                    {
                        deleteRows.Add(new []{i, 1});
                        falsing = true;
                    }
                    else
                    {
                        deleteRows[deleteRows.Count - 1][1]++;
                    }
                }
                else
                {
                    if (falsing)
                    {
                        falsing = false;
                    }
                }
            }

            foreach (var arr in deleteRows)
            {
                _sheet.DeleteRow(arr[0], arr[1]);
                _end = _end - arr[1];
            }
        }

        private bool FormatCell(string col, int row)
        {
            var c = _sheet.Cells[col + row];
            if (c == null || c.Value == null || c.Value.ToString() == "" || c.Text.Length < 5 || c.Text[1] == '.' || c.Text[1] == ',')
            {
                return false;
            }

            var text = c.Text.Replace(" ", "");
            
            if (text[0] == '0')
            {
                text = text.Substring(1);
            }

            if (int.Parse(text[0].ToString() + text[1]) < 30)
            {
                if (text.Contains("N"))
                {
                    text = text.Replace("N", "");
                }
                if (!text.Contains("E"))
                {
                    text = text.Substring(0, 2) + "E" + text.Substring(2);
                }
            }
            else
            {
                if (text.Contains("E"))
                {
                    text = text.Replace("E", "");
                }
                if (!text.Contains("N"))
                {
                    text = text.Substring(0, 2) + "N" + text.Substring(2);
                }
            }

            c.Value = text;
            return true;
        }
    }
}