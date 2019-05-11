using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Windows.Forms;
using FathersApp.Properties;

namespace FathersApp
{
    public class DataTransport
    {
        //to finish
        public static string DTinit()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (Settings.Default.PathDocumentation.ToString() == string.Empty)
                openFileDialog.InitialDirectory = @"C:\";
            else
                openFileDialog.InitialDirectory = Settings.Default.PathDocumentation.ToString();

            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel Files|*.xls; *.xlsx; *.xlsm";
            openFileDialog.Title = "Wybierz plik excel bazy danych";

            DialogResult dialogResult = new DialogResult();
            dialogResult = openFileDialog.ShowDialog();
            switch(dialogResult)
            {
                case DialogResult.OK:
                    Settings.Default.PathDocumentation = openFileDialog.FileName;
                    return openFileDialog.FileName;
                default:
                    break;
            }
            return string.Empty;
        }
        public static List<string> CheckIfFileExist(string file)
        {
            var pathList = new List<string>();

            return pathList;
        }
        public static string DTAdd(string file, DataGridView dataGridView)
        {
            Excel.Application ExApp;
            Excel.Workbook ExWorkbook;
            Excel.Worksheet ExWorksheet;

            ExApp = new Excel.Application();
            ExApp.Visible = false;

            //check how many rows worksheet has itself & check if worksheet is protected
            if (File.Exists(file))
            {
                ExWorkbook = ExApp.Workbooks.Open(file);
                ExWorksheet = ExWorkbook.Worksheets["Arkusz1"];
                ExWorksheet.Activate();
                if (!ExWorksheet.ProtectContents)
                {
                    var emptyRow = ExWorksheet.UsedRange.Rows.Count;
                    var col_max = dataGridView.Columns.Count;
                    var row_max = dataGridView.Rows.Count;

                    for (var row = 0; row < row_max; row++)
                    {
                        for (var col = 0; col < col_max; col++)
                        {
                            ExWorksheet.Cells[emptyRow + row + 1, col + 1].value = dataGridView.Rows[row].Cells[col].Value;
                        }
                    }

                    ExWorkbook.Save();
                    ExWorkbook.Close();
                    ExApp.Quit();
                    return "Dodanie zakonczone pomyslnie" + file;
                }
                ExWorkbook.Close();
                ExApp.Quit();
                return "Plik jest tylko do odczytu" + file;
            }
            else
            {
                ExApp.Quit();
                return "Nie znalazlem pliku: " + file;
            }   
        }


        public static string ImportFromExcel(string file, DataGridView dataGridView)
        {
            Excel.Application ExApp;
            Excel.Workbook ExWorkbook;
            Excel.Worksheet ExWorksheet;

            ExApp = new Excel.Application();
            ExApp.Visible = false;

            if (File.Exists(file))
            {
                ExWorkbook = ExApp.Workbooks.Open(file);
                ExWorksheet = ExWorkbook.Worksheets["Arkusz1"];
                ExWorksheet = ExWorkbook.ActiveSheet;

                /*Kod do exportu DataGridView*/
                if (!ExWorksheet.ProtectContents)
                {
                    var col_max = dataGridView.Columns.Count-1;
                    var row_max = dataGridView.Rows.Count-1;
                    var exRows = ExWorksheet.UsedRange.Rows.Count;


                    for (var row = 0; row < exRows; row++)
                    {
                        if (row > row_max)
                            dataGridView.Rows.Add(); // problem nie da sie programistycznie dodać do datagridview

                        for (var col = 0; col < col_max; col++)
                        {
                            dataGridView.Rows[row].Cells[col].Value = ExWorksheet.Cells[row + 1, col + 1].value;
                        }
                    }
                    ExWorkbook.Save();
                    ExWorkbook.Close();
                    ExApp.Quit();
                    return "Import udany";
                }
                ExWorkbook.Close();
                ExApp.Quit();
                return "Plik jest tylko do odczytu: " + file;
            }
            else
            {
                ExApp.Quit();
                return "Nie znaleziono pliku: " + file;
            }
                
        }
        public static string DTExport(string file, DataGridView dataGridView)
        {
            Excel.Application ExApp;
            Excel.Workbook ExWorkbook;
            Excel.Worksheet ExWorksheet;

            ExApp = new Excel.Application();
            ExApp.Visible = false;

            if (File.Exists(file))
            {
                ExWorkbook = ExApp.Workbooks.Open(file);
                ExWorksheet = ExWorkbook.Worksheets["Arkusz1"];
                if (!ExWorksheet.ProtectContents)
                {
                    ExWorksheet = ExWorkbook.ActiveSheet;

                    /*Kod do exportu DataGridView*/
                    var col_max = dataGridView.Columns.Count-1;
                    var row_max = dataGridView.Rows.Count-1;

                    for (var row = 0; row < row_max; row++)
                    {
                        for (var col = 0; col < col_max; col++)
                        {
                            ExWorksheet.Cells[row + 1, col + 1].value = dataGridView.Rows[row].Cells[col].Value.ToString();
                        }
                    }
                    ExWorkbook.Save();
                    ExWorkbook.Close();
                    ExApp.Quit();
                    return "Export udany: " + file;
                }
                ExWorkbook.Close();
                ExApp.Quit();
                return "Plik jest tylko do odczytu: " + file;
            }
            else
            {
                ExApp.Quit();
                return "Plik nie istnieje: " + file;
            }

        }
    }
}
