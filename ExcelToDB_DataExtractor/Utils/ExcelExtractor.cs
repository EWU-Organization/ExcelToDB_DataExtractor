using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelToDB_DataExtractor.Utils
{
    internal static class ExcelExtractor
    {
        public static void ProcessExcelFiles(string[] directoryPath)
        {
            try
            {
                foreach (var file in directoryPath)
                {
                    ProcessExcelFile(file);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }

        static void ProcessExcelFile(string file)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(file, false))
            {
                // Access the workbook part
                WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet? sheet = workbookPart?.Workbook?.Descendants<Sheet>().FirstOrDefault(); ;

                foreach (var item in workbookPart.Workbook.Descendants<Sheet>())
                {
                    if (item.Name == "PO") sheet = workbookPart?.Workbook?.Descendants<Sheet>().FirstOrDefault(s => s.Name == item.Name);
                }

                if (sheet != null)
                {
                    // Access the worksheet part
                    WorksheetPart? worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                    // Get the sheet data
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Iterate through rows and cells
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            // Access the cell value
                            string cellValue = GetCellValue(cell, workbookPart);
                            Console.Out.WriteLine($"{cellValue}\t");
                        }

                        Console.WriteLine(); // Move to the next row
                    }
                }
            }
        }

        static string GetCellValue(Cell cell, WorkbookPart workbookPart)
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
    }
}
