using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Aspose.Email.Clients;
using Aspose.Email.Storage.Mbox;
using System.Windows.Media;
using System.Diagnostics;

namespace MboxReader
{
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Mbox files (*.mbox)|*.mbox|All files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                txtFileName.Text = openFileDialog.FileName;
            }
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text))
            {
                MessageBox.Show("Please select an mbox file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string fileName = txtFileName.Text;
            string outputDirectory = "output"; // Change this to your desired output directory

            try
            {
                Directory.CreateDirectory(outputDirectory);

                int totalMessages = CountTotalMessages(fileName);
                progressBar.Maximum = totalMessages;
                progressBar.Value = 0;

                var mboxLoadOptions = new MboxLoadOptions();
                using (var mbox = MboxStorageReader.CreateReader(fileName, mboxLoadOptions))
                {
                    foreach (var eml in mbox.EnumerateMessages())
                    {

                        string emlFileName = Path.Combine(outputDirectory, $"{eml.Subject}.eml");
                        eml.Save(emlFileName);
                        progressBar.Value++;

                    }
                }


                MessageBox.Show("All emails saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                // Open the output directory
                Process.Start(outputDirectory);
                // Clear input values and reset progress bar
                txtFileName.Text = string.Empty;
                progressBar.Value = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private int CountTotalMessages(string fileName)
        {
            int count = 0;
            var mboxLoadOptions = new MboxLoadOptions();
            using (var mbox = MboxStorageReader.CreateReader(fileName, mboxLoadOptions))
            {
                foreach (var eml in mbox.EnumerateMessages())
                {
                    count++;
                }
            }
            return count;
        }




    }
}
