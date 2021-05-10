using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace ExcelToExcel.Models
{
    public class EspeceXL
    {
        private IXLWorkbook wb;
        private FileStream fs;
        

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
            fs = new FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            wb = new XLWorkbook(fs);
            IsReadOnly = true;
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
            string result = string.Empty;

            var worksheet = wb.Worksheet("especes");
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

    }
}
