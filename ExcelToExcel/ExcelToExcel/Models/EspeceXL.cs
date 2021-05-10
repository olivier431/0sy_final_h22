using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace ExcelToExcel.Models
{
    public class EspeceXL
    {
        private string sheetname = "especes";

        private IXLWorkbook wb;

        public string Filename { get; set; }
        public bool IsReadOnly { get; private set; } = false;


        public EspeceXL(string filename)
        {
            Filename = filename;
        }

        public void LoadFile()
        {
            wb = new XLWorkbook(Filename);
            IsReadOnly = false;
        }

        public void LoadFileReadOnly()
        {
            using (var fs = new FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wb = new XLWorkbook(fs);
                IsReadOnly = true;
            }
        }

        public object GetCell(int sheet, int row, int column)
        {
            return wb.Worksheet(sheet).Cell(row, column).Value;
        }

        public object GetCell(string sheet, int row, int column)
        {
            return wb.Worksheet(sheet).Cell(row, column).Value;
        }

        public string GetCSV()
        {
            if (!validFileContent())
            {
                throw new ArgumentException("Mauvais format de fichier!");
            }

            string result = string.Empty;

            var worksheet = wb.Worksheet(sheetname);
            var lastCellAddress = worksheet.RangeUsed().LastCell().Address;

            var res = worksheet.Rows(1, lastCellAddress.RowNumber)
                .Select(r => string.Join(",", r.Cells(1, lastCellAddress.ColumnNumber)
                        .Select(cell =>
                        {
                            var cellValue = cell.GetValue<string>();
                            return cellValue.Contains(",") ? $"\"{cellValue}\"" : cellValue;
                        })));

            foreach (var r in res)
            {
                result += r + Environment.NewLine;
            }

            return result;
        }

        private bool validFileContent()
        {
            
            var valid = wb.Worksheets.Contains(sheetname);

            if (valid)
            {
                var ws = wb.Worksheet(sheetname);
                var nom = ws.Cell(1, 1).Value.ToString().Trim().ToLower();
                var nomLatin = ws.Cell(1, 2).Value.ToString().Trim().ToLower();
                var habitat = ws.Cell(1, 3).Value.ToString().Trim().ToLower();

                valid &= nom == "nom";
                valid &= nomLatin == "nom latin";
                valid &= habitat == "habitat";
            }

            return valid;
        }

    }
}
