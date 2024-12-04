//using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using WpfHelpers;
using WpfHelpers.Controls;
//using Formatting = Newtonsoft.Json.Formatting;


namespace BaseApp.ViewModels
{

    public class Settings : ViewModelBase
    {

        private string _ipAddress;
        private int _port;
        private string _filePath;

        public ICommand OpenFileDialog { get; set; }
        public ICommand SaveCommand { get; }

        public string IpAddress
        {
            get => _ipAddress;
            set
            {
                _ipAddress = value;
                OnPropertyChanged(nameof(IpAddress));
            }
        }

        public int Port
        {
            get => _port;
            set
            {
                _port = value;
                OnPropertyChanged(nameof(Port));
            }
        }

        public string FilePath
        {
            get => _filePath;
            set
            {
                _filePath = value;
                OnPropertyChanged(nameof(FilePath));
            }
        }
        

        private string _ExcelPath;
        public string ExcelPath
        {
            get
            {
                return this._ExcelPath;
            }
            set
            {
                this._ExcelPath = value;
                this.OnPropertyChanged("ExcelPath");
            }
        }
        public Settings()
        {
          
            OpenFileDialog = new DelegateCommand((param) =>
            {
                using (var folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Select a Folder";
                    folderBrowserDialog.ShowNewFolderButton = true; // Allow creating new folders if necessary

                    var result = folderBrowserDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        ExcelPath = folderBrowserDialog.SelectedPath; // Store the selected folder path
                        ReadExcel(ExcelPath);                                            
                    }
                }
            });

                SaveCommand = new RelayCommand(SaveSettings);
        }

        private void ReadExcel(string excelPath)
        {

        }


        private void SaveSettings()
        {
            // Logic to save IP address, port, and file path to a file or database
            SaveToSettingsFile(IpAddress, Port, FilePath, ExcelPath);


        }

        private void SaveToSettingsFile(string ip, int port, string filePath, string excelPath)
        {
            if (string.IsNullOrWhiteSpace(ExcelPath))
            {
                // Prompt user to select an Excel file
                System.Windows.MessageBox.Show("Please select an Excel file to save the data.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                SaveToExcel(Bootstrap.settingPath, IpAddress, Port, excelPath);
                System.Windows.MessageBox.Show("Settings saved successfully to Excel file.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"An error occurred while saving settings to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }


        private void SaveToExcel(string filePath, string ipAddress, int port, string excelPath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                // Check if the workbook exists, otherwise create it
                var worksheet = workbook.Worksheets.Count > 0 ? workbook.Worksheet(1) : workbook.Worksheets.Add("Settings");

                // Add headers if worksheet is new
                if (worksheet.Cell(1, 1).IsEmpty())
                {
                    worksheet.Cell(1, 1).Value = "IP Address";
                    worksheet.Cell(1, 2).Value = "Port";
                    worksheet.Cell(1, 3).Value = "File Path";
                    worksheet.Cell(1, 4).Value = "Folder Path";
                }

                // Find the next available row
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                var newRow = lastRow + 1;

                // Add data to the worksheet
                worksheet.Cell(newRow, 1).Value = ipAddress;
                worksheet.Cell(newRow, 2).Value = port;
                worksheet.Cell(newRow, 3).Value = filePath;
                worksheet.Cell(newRow, 4).Value = excelPath;

                // Save the workbook
                workbook.SaveAs(filePath);
            }
        }


    }
    
}
