using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;

namespace TestApp
{
    public partial class MainWindow : Window
    {
        private string selectedFile = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv"
            };

            if (dialog.ShowDialog() == true)
            {
                selectedFile = dialog.FileName;
                FilePathBox.Text = selectedFile;
            }
        }

        private void FileDropped(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    selectedFile = files[0];
                    FilePathBox.Text = selectedFile;
                }
            }
        }

        private void RunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedFile))
            {
                MessageBox.Show("Please select or drop a CSV file first.");
                return;
            }

            try
            {
                var output = Path.ChangeExtension(selectedFile, ".xlsx");
                ResponseTimeConverter.Convert(selectedFile, output);
                MessageBox.Show($"Excel created:\n{output}", "Done", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}