using System.Configuration;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ExcelToDB_DataExtractor.Views.UC
{
    /// <summary>
    /// Interaction logic for Settings.xaml
    /// </summary>
    public partial class Settings : UserControl
    {
        public Settings()
        {
            InitializeComponent();
        }

        private string EncryptPassword(string password)
        {
            using (Aes aes = Aes.Create())
            {
                byte[] key = Encoding.UTF8.GetBytes("h{K`V4ubq|sKZ~K@"); // Replace with your key
                byte[] iv = Encoding.UTF8.GetBytes("j}+Ms+wj0/.ODn<f"); // Replace with your IV
                aes.Key = key;
                aes.IV = iv;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(cs))
                        {
                            sw.Write(password);
                        }
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
        }

        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(username.Text) || string.IsNullOrEmpty(password.Password) || string.IsNullOrEmpty(port.Text) || string.IsNullOrEmpty(host.Text))
            {
                MessageBox.Show("Please fill in all fields", "Excel To Database");
                return;
            }

            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).AppSettings.Settings;

            config.Remove("Username");
            config.Add("Username", username.Text);

            config.Remove("Password");
            config.Add("Password", EncryptPassword(password.Password));

            config.Remove("Port");
            config.Add("Port", port.Text);

            config.Remove("ConnectionString");
            config.Add("ConnectionString", host.Text);

            ConfigurationManager.RefreshSection("appSettings");
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.Instance.settings.Visibility = Visibility.Collapsed;
            MainWindow.Instance.startPage.Visibility = Visibility.Visible;

            MainWindow.Instance.WindowAnimation(350, 285);
        }
    }
}
