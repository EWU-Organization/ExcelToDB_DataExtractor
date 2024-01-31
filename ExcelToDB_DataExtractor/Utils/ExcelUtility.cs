using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace ExcelToDB_DataExtractor.Utils
{
    internal class ExcelUtility
    {
        private static WorkbookPart? _WorkbookPart;
        private static WorksheetPart? _WorksheetPart;
        private static SpreadsheetDocument? _SpreadsheetDocument;

        /// <summary>
        /// Initialize the ExcelUtility class with an excel file and sheet name
        /// </summary>
        /// <param name="file">File path to .xlsx file</param>
        /// <param name="sheetName">An excel sheet that is in the .xlsx file</param>
        public static void Initialize(string file, string sheetName)
        {
            _SpreadsheetDocument = SpreadsheetDocument.Open(file, false);

            // Access the workbook part
            _WorkbookPart = _SpreadsheetDocument.WorkbookPart;
            Sheet? sheet = _WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            if (sheet != null) _WorksheetPart = (WorksheetPart)_WorkbookPart.GetPartById(sheet.Id);
        }
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            SharedStringTablePart? stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && stringTablePart != null)
            {
                // If the cell contains a shared string, get the value from the shared string table
                int index = int.Parse(cell.InnerText);
                return stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
            }
            else
            {
                // Otherwise, get the cell value directly
                return cell.InnerText;
            }
        }

        private static string GetColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static string? CellValue(int row, int column)
        {
            string cellReference = GetColumnName(column) + row;

            // Get the cell with the specified reference
            Cell? cell = _WorksheetPart?.Worksheet.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.CurrentCultureIgnoreCase));

            if (cell != null)
            {
                // Get the cell value
                return GetCellValue(cell, _WorkbookPart);
            }

            return null;
        }

        /// <summary>
        /// Returns the number of rows in the excel sheet
        /// </summary>
        /// <returns>int</returns>
        public static int GetNumberOfRows()
        {
            if (_WorksheetPart != null) return (int)(_WorksheetPart?.Worksheet.Descendants<Row>().Count());
            else return 0;
        }

        public static void Dispose()
        {
            _SpreadsheetDocument?.Dispose();
            _SpreadsheetDocument = null;
            _WorkbookPart = null;
            _WorksheetPart = null;
        }

        public static int ParseInteger(string input)
        {
            if (input == null) return 0;
            MatchCollection matches = new Regex(@"\d+").Matches(input);

            int number = 0;
            foreach (Match match in matches)
            {
                number = int.Parse(match.Value);
            }
            return number;
        }
        public static double ParseDouble(string input)
        {
            if (input == null) return 0;
            MatchCollection matches = new Regex(@"\d+(\.\d+)?").Matches(input);

            // Extract and print the integers
            double number = 0;
            foreach (Match match in matches)
            {
                number = double.Parse(match.Value);
            }
            return number;
        }
    }
}
