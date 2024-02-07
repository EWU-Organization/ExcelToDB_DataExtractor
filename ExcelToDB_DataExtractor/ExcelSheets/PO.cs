using ExcelToDB_DataExtractor.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB_DataExtractor.ExcelSheets
{
    internal class PO
    {
        public static void ExtractInfo(string file)
        {
            ExcelUtility.Initialize(file, "PO");
            ExtractCalculatedCOInfo();
        }

        static void ExtractCalculatedCOInfo()
        {
            int rows = ExcelUtility.GetNumberOfRows();
            double[] coValues = [];
            for (int column = 9; column < 21; column++)
            {
                for (int row = 1; row < rows; row++)
                {
                    string value = ExcelUtility.CellValue(row, column);
                    if(value != "")
                    {
                        coValues[column] = ExcelUtility.ParseDouble(value);
                    }
                }
            }
        }
    }
}
