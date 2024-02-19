using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using XLSX_toXML_toDoc.ViewModel;

namespace XLSX_toXML_toDoc.View
{
    /// <summary>
    /// Логика взаимодействия для MainView.xaml
    /// </summary>
    public partial class MainView : UserControl
    {
        private string ImportXlsxInitialStatus = "Выберите Excel-файл";
        public MainView()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
            XlsxImportStatus.Text = ImportXlsxInitialStatus;
            BtnImportXlsx.IsEnabled = true;
            BtnFormReport.IsEnabled = false;
            BtnSaveAsDoc.IsEnabled = false;
        }

        private void BtnImportXlsx_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.ImportAndSetDirectory();
                
            }
            if (MainViewModel.IsFileImported == true) {
                XlsxImportStatus.Text = "Excel-файл выбран";
                XlsxImportStatus.Foreground = Brushes.LimeGreen;
                BtnImportXlsx.IsEnabled = false;
                BtnFormReport.IsEnabled = true;
                ReportStatus.Text = "Сформируйте отчет";
                ReportStatus.Foreground = Brushes.Black;
            } 
            
        }

        private void FormReport_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.GenerateReport();
            }
            if (MainViewModel.IsReportDone == true)
            {
                    ReportStatus.Text = "Отчет сформирован";
                    ReportStatus.Foreground = Brushes.LimeGreen;
                    BtnFormReport.IsEnabled = false;
                    BtnImportXlsx.IsEnabled= false;
                    BtnSaveAsDoc.IsEnabled= true;
                    SaveToDocStatus.Text = "Файл с отчетом можно сохранить";
            }

        }

        private void BtnSaveAsDoc_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.GenerateAndSaveDoc();
            }
            if (MainViewModel.IsReportSaved == true) {
                BtnImportXlsx.IsEnabled = true;
                BtnSaveAsDoc.IsEnabled = false;

                XlsxImportStatus.Text = ImportXlsxInitialStatus;
                XlsxImportStatus.Foreground = Brushes.Black;
                ReportStatus.Text = "";
                SaveToDocStatus.Text = "Отчет сохранен";
                SaveToDocStatus.Foreground = Brushes.LimeGreen;

                DispatcherTimer timer = new DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(0.5);
                timer.Tick += (sender, args) =>
                {
                    SaveToDocStatus.Foreground = Brushes.Black;
                    SaveToDocStatus.Text = ""; 
                    timer.Stop();             
                };
                timer.Start();
            }
        }
    }
}
