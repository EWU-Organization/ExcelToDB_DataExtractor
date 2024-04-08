using ExcelToDB_ConsoleApp.Utils;

namespace ExcelToDB_ConsoleApp
{
    internal class Program
    {
        private static string[]? _excelFilesPath;
        private static string _Username = null;
        private static string _Password = null;
        private static string _Directory = null;

        static void Main(string[] args)
        {
            // Check if any arguments are provided
            if (args.Length == 0)
            {
                Console.WriteLine("No arguments provided.");
                return;
            }

            // Parsing arguments
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-u":
                    case "--user":
                    case "--username":
                        if (i + 1 < args.Length)
                            _Username = args[i + 1];
                        break;
                    case "-p":
                    case "--pass":
                    case "--password":
                        if (i + 1 < args.Length)
                            _Password = args[i + 1];
                        break;
                    case "-d":
                    case "--dir":
                    case "--directory":
                        if (i + 1 < args.Length)
                            _Directory = args[i + 1];
                        break;
                }
            }

            Work();
        }

        static async void Work()
        {
            try
            {
                // Check if the argument starts with "-d"
                if (_Username != null && _Password != null && _Directory != null)
                {
                    // Extract the directory path
                    _excelFilesPath = Directory.GetFiles(_Directory, "*.xlsx");
                    if (_excelFilesPath.Count() < 1) Console.WriteLine("Excel file(s) not found");
                    else await ExcelExtractor.ProcessExcelFiles(_excelFilesPath);
                }
                else
                {
                    // Handle unrecognized arguments
                    Console.WriteLine("Unrecognized argument");
                    Console.WriteLine("Usage: ExcelToDB -u myusername -p mypassword -d \"/path/to/directory\"");
                    Console.WriteLine("Usage: ExcelToDB --user myusername --pass mypassword --dir \"/path/to/directory\"");
                    Console.WriteLine("Usage: ExcelToDB --username myusername --password mypassword --directory \"/path/to/directory\"");
                }
            }
            catch (DirectoryNotFoundException)
            {
                Console.Error.WriteLine("Directory not found");
            }
        }
    }
}
