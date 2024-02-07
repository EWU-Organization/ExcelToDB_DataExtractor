using ExcelToDB_DataExtractor.Utils;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ExcelToDB_DataExtractor.Views.UC
{
    /// <summary>
    /// Interaction logic for StartPage.xaml
    /// </summary>
    public partial class StartPage : UserControl
    {
        private string[]? _excelFilesPath;

        public StartPage()
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
                // Read Excel files
                _excelFilesPath = Directory.GetFiles(folderBrowserDialog.FolderName, "*.xlsx");
                folderPathTextBlock.Text = $"Number of files found: {_excelFilesPath.Length}";
                folderPathTextBlock.Visibility = Visibility.Visible;
                if (_excelFilesPath.Count() < 1)
                {
                    MessageBox.Show("Excel file(s) not found");
                    folderPathTextBlock.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            if (_excelFilesPath?.Length >= 1)
            {
                ExcelExtractor.ProcessExcelFiles(_excelFilesPath);
            }
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.Instance.startPage.Visibility = Visibility.Collapsed;
            MainWindow.Instance.settings.Visibility = Visibility.Visible;

            MainWindow.Instance.WindowAnimation(285, 350);
        }
    }
}
