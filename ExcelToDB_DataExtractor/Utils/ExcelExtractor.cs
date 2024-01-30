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
                    ExcelUtility.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }

        static void ProcessExcelFile(string file)
        {
            ExcelUtility.Initialize(file);

            var test = ExcelUtility.CellValue(25, 3);
            Console.WriteLine(test);
        }
    }
}
