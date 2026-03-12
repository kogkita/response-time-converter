using System;
using System.Windows;
using Microsoft.Win32;

namespace ResponseTimeConverter
{
    public partial class MainWindow : Window
    {
        // Variable to store the selected file
        private string selectedFile = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        // Browse button click
        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            if (dialog.ShowDialog() == true)
            {
                selectedFile = dialog.FileName;
                FilePathBox.Text = selectedFile;
            }
        }

        // Drag & Drop file
        private void FileDropped(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length > 0)
                {
                    selectedFile = files[0];
                    FilePathBox.Text = selectedFile;
                }
            }
        }

        // Run Processing button
        private void RunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFile))
            {
                MessageBox.Show("Please select or drop a file first.");
                return;
            }

            try
            {
                // Your processing logic will go here
                MessageBox.Show("Processing file:\n\n" + selectedFile);

                // Example placeholder
                // ProcessFile(selectedFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing file:\n" + ex.Message);
            }
        }
    }
}