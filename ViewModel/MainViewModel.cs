using FirstProject.Helper;    // includes LogConverter + TechStudioMfgLayoutConverter
using FirstProject.Models;    // TechStudioRunRow
using Microsoft.Win32;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace FirstProject.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        // ===== Existing fields/ctor =====
        private ObservableCollection<string> logLines;

        public MainViewModel()
        {
            // Existing commands (your current parser/exporter)
            BrowseCommand = new RelayCommand(Browse);
            DownloadCommand = new RelayCommand(Download);

            // NEW: Tech Studio tab commands (MFG-style 33-col sheet)
           // TechStudioBrowseCommand = new RelayCommand(TechStudioBrowse);
            TechStudioDownloadMfgCommand = new RelayCommand(TechStudioDownloadMfg);
        }

        // ===== Existing properties =====
        private string selectedFilePath;
        private string generatedExcelPath = "ExcelFileName";
        private string excelPath;

        public string SelectedFilePath
        {
            get => selectedFilePath;
            set { selectedFilePath = value; OnPropertyChanged(nameof(SelectedFilePath)); }
        }

        public string ExcelPath
        {
            get => excelPath;
            set { excelPath = value; OnPropertyChanged(nameof(ExcelPath)); }
        }

        public string GeneratedExcelPath
        {
            get => generatedExcelPath;
            set { generatedExcelPath = value; OnPropertyChanged(nameof(GeneratedExcelPath)); }
        }

        public ICommand BrowseCommand { get; }
        public ICommand DownloadCommand { get; }

        // ===== Existing methods (Manufacturing API flow) =====
        private void Browse()
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "(*.txt)|*.txt",
                Title = "Select the log file."
            };

            bool? success = ofd.ShowDialog();
            if (success == true)
            {
                SelectedFilePath = ofd.FileName;
                logLines = new ObservableCollection<string>(File.ReadAllLines(SelectedFilePath));
            }
            else
            {
                MessageBox.Show("File not Found", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Download()
        {
            if (string.IsNullOrEmpty(SelectedFilePath) || !File.Exists(SelectedFilePath))
            {
                MessageBox.Show("Please select a valid log File.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Title = "Save Excel File",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = Path.GetFileNameWithoutExtension(SelectedFilePath) + ".xlsx"
            };

            bool? saveResult = saveDialog.ShowDialog();
            if (saveResult != true) return;

            string outputPath = saveDialog.FileName;
            try
            {
                // Your existing converter call
                LogConverter.ConvertLogToExcel(SelectedFilePath, outputPath);

                ExcelPath = Path.GetFileName(outputPath);
                GeneratedExcelPath = outputPath;

                MessageBox.Show($"File saved to:\n{outputPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = outputPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during conversion or saving:\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // =====================================================================
        // ====================  NEW: Tech Studio (MFG layout) ==================
        // =====================================================================

        private string techStudioFilePath = string.Empty;
        public string TechStudioFilePath
        {
            get => techStudioFilePath;
            set
            {
                techStudioFilePath = value;
                OnPropertyChanged(nameof(TechStudioFilePath));
                OnPropertyChanged(nameof(CanTechStudioRun));
            }
        }

        // Optional preview collection if you later bind a grid
        public ObservableCollection<TechStudioRunRow> TechStudioRows { get; } = new();

        public bool CanTechStudioRun =>
            !string.IsNullOrEmpty(TechStudioFilePath) && File.Exists(TechStudioFilePath);

    
        public ICommand TechStudioDownloadMfgCommand { get; }

        private void TechStudioDownloadMfg()
        {
            if (!CanTechStudioRun)
            {
                MessageBox.Show("Please select a valid Tech Studio log.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Title = "Save Excel (MFG layout)",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = Path.GetFileNameWithoutExtension(TechStudioFilePath) + "_MfgLayout.xlsx"
            };
            if (sfd.ShowDialog() != true) return;

            string outputPath = sfd.FileName;
            try
            {
                // Build the 33-column sheet identical to your MFG API sample
                Techstudioconverter.ConvertToMfgExcel(TechStudioFilePath, outputPath);

                MessageBox.Show($"File saved to:\n{outputPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = outputPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during conversion or saving:\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ===== INotifyPropertyChanged =====
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
