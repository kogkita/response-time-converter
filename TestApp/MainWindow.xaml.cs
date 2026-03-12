using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace TestApp
{
    public partial class MainWindow : Window
    {
        private string selectedFile = "";
        private Button? activeNavButton;

        public MainWindow()
        {
            InitializeComponent();
            activeNavButton = NavConvert;
        }

        // ── Custom title bar ─────────────────────────────────

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
                ToggleMaximize();
            else
                DragMove();
        }

        private void MinimizeWindow_Click(object sender, RoutedEventArgs e)
            => WindowState = WindowState.Minimized;

        private void MaxRestoreWindow_Click(object sender, RoutedEventArgs e)
            => ToggleMaximize();

        private void CloseWindow_Click(object sender, RoutedEventArgs e)
            => Close();

        private void ToggleMaximize()
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
                MaxRestoreBtn.Content = "\uE922"; // restore icon
            }
            else
            {
                WindowState = WindowState.Maximized;
                MaxRestoreBtn.Content = "\uE923"; // maximize icon
            }
        }

        // ── Navigation ───────────────────────────────────────

        private void SetActivePage(Button clicked, UIElement page)
        {
            if (activeNavButton != null)
                activeNavButton.Style = (Style)Resources["NavButtonStyle"];

            clicked.Style = (Style)Resources["NavButtonActiveStyle"];
            activeNavButton = clicked;

            PageConvert.Visibility = Visibility.Collapsed;
            page.Visibility = Visibility.Visible;
        }

        private void NavConvert_Click(object sender, RoutedEventArgs e)
            => SetActivePage(NavConvert, PageConvert);

        // ── Convert Response Times page ──────────────────────

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv" };
            if (dialog.ShowDialog() == true)
            {
                selectedFile = dialog.FileName;
                FilePathBox.Text = selectedFile;
            }
        }

        private void FileDropped(object sender, DragEventArgs e)
        {
            ResetDropZone();
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

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                DropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2563EB"));
                DropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827"));
            }
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
            => ResetDropZone();

        private void ResetDropZone()
        {
            DropZone.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E2640"));
            DropZone.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D1020"));
        }

        private void RunProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedFile))
            {
                MessageBox.Show("Please select or drop a CSV file first.", "No File", MessageBoxButton.OK, MessageBoxImage.Warning);
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
