using ExcelToDB_DataExtractor.Utils;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace ExcelToDB_DataExtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string[]? _excelFilesPath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectExcelFilesFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenFolderDialog folderBrowserDialog = new OpenFolderDialog
            {
                Title = "Select the folder containing the Excel files",
            };

            if (folderBrowserDialog.ShowDialog() == true)
            {
                folderPathTextBlock.Text = folderBrowserDialog.FolderName;

                // Read Excel files
                _excelFilesPath = Directory.GetFiles(folderBrowserDialog.FolderName, "*.xlsx");
                folderPathTextBlock.Visibility = Visibility.Visible;
                if (_excelFilesPath.Count() < 1)
                {
                    MessageBox.Show("Excel file(s) not found");
                    folderPathTextBlock.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void Process_GenerateReport_Click(object sender, RoutedEventArgs e)
        {
            if (_excelFilesPath?.Length >= 1)
            {
                ExcelExtractor.ProcessExcelFiles(_excelFilesPath);
            }
        }
    }
}