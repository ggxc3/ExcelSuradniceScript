using System.Windows.Forms;

namespace ExcelSuradniceScript
{
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

        public static string SaveFile()
        {
            var dialog = new SaveFileDialog
            {
                Title = "Uložiť ako",
                DefaultExt = "xlsx",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                RestoreDirectory = true
            };
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;
        }
    }
}