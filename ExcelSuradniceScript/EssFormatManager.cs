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

        public void FormatColumns(string[][] columns)
        {
            foreach (var col in columns)
            {
                FormatColumn(col[0], col[1]);
            }
        }
        
        private void FormatColumn(string column, string coorType)
        {
            var deleteRows = new List<int[]>();
            bool falsing = false;
            for (int i = _start; i <= _end; i++)
            {
                Console.WriteLine($"Riadok: {i}");
                if (!FormatCell(column, i, coorType))
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

            var index = 1;
            for (int i = deleteRows.Count - 1; i > -1; i--)
            {
                Console.WriteLine($"Mažem nespravne riadky {index}/{deleteRows.Count}");
                _sheet.DeleteRow(deleteRows[i][0], deleteRows[i][1]);
                _end = _end - deleteRows[i][1];
                index++;
            }
        }

        private bool FormatCell(string col, int row, string coorType)
        {
            var c = _sheet.Cells[col + row];
            if (c == null || c.Value == null || c.Value.ToString() == "" || c.Text.Length < 5 || c.Text[1] == '.' || c.Text[1] == ',')
            {
                return false;
            }

            var text = c.Text.Replace(" ", "");
            text = text.Replace(",", ".");
            
            if (text[0] == '0')
            {
                text = text.Substring(1);
            }
            
            if (int.TryParse(text[0].ToString() + text[1], out _) 
                || (text[0].ToString() == "E" 
                    || text[0].ToString() == "N") 
                    && int.TryParse(text[1].ToString(), out _)
                )
            {
                if (coorType == "E")
                {
                    if (text.Contains("N"))
                    {
                        text = text.Replace("N", "");
                    }

                    if (!text.Contains("E"))
                    {
                        text = text.Substring(0, 2) + "E" + text.Substring(2);
                    }
                } else if (coorType == "N")
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
            }
            else
            {
                return false;
            }

            c.Value = text;
            return true;
        }
    }
}