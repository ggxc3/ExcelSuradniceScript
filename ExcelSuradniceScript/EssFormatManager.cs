using System;
using OfficeOpenXml;

namespace ExcelSuradniceScript
{
    class EssFormatManager
    {
        public static void FormatColumns(ExcelWorksheet sheet, string[] columns, int start, int end)
        {
            foreach (var col in columns)
            {
                FormatColumn(sheet, col, start, end);
            }
        }
        
        private static void FormatColumn(ExcelWorksheet sheet, string column, int start, int end)
        {
            for (int i = start; i <= end; i++)
            {
                FormatCell(sheet, column, i);
            }
        }

        private static void FormatCell(ExcelWorksheet sheet, string col, int row)
        {
            var c = sheet.Cells[col + row];
            if (c == null || c.Value == null || c.Value.ToString() == "" || c.Text.Length < 5 || c.Text[1] == '.' || c.Text[1] == ',')
            {
                sheet.DeleteRow(row);
                return;
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
        }
    }
}