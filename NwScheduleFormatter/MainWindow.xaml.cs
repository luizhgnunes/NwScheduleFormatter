using Microsoft.Win32;
using NwScheduleFormatter.Configuration;
using NwScheduleFormatter.Services;
using NwScheduleFormatter.Services.Interfaces;
using System.Reflection;
using System.Windows;

namespace NwScheduleFormatter;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly IPdfGenerationService _pdfGenerationService;
    private readonly ICsvParsingService _csvParsingService;
    private readonly IWhatsAppMessagingService _whatsAppMessagingService;

    public MainWindow(IPdfGenerationService pdfGenerationService, ICsvParsingService csvParsingService, IWhatsAppMessagingService whatsAppMessagingService)
    {
        _pdfGenerationService = pdfGenerationService;
        _csvParsingService = csvParsingService;
        _whatsAppMessagingService = whatsAppMessagingService;

        InitializeComponent();

        SetWindowTitle();

        PopulateYearComboBox();
        PopulateMonthComboBox();
        PopulateDefaultOutputDirectory();
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        await UpdateService.CheckForUpdatesAsync();
    }

    private void SetWindowTitle()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        this.Title = $"NwScheduleFormatter - v{version.Major}.{version.Minor}.{version.Build}";
    }

    private void PopulateYearComboBox()
    {
        var currentYear = DateTime.Now.Year;

        for (int i = 0; i <= 1; i++)
        {
            YearComboBox.Items.Add(currentYear + i);
        }
        YearComboBox.SelectedIndex = 0;
    }

    private void PopulateMonthComboBox()
    {
        for (int i = 1; i <= 12; i++)
        {
            MonthComboBox.Items.Add(i);
        }
        var currentMonth = DateTime.Now.Month;
        MonthComboBox.SelectedIndex = currentMonth - 1;
    }

    private void PopulateDefaultOutputDirectory()
    {
        TxtOutputDirectory.Text = $@"{ApplicationSettings.DEFAULT_OUTPUT_DIRECTORY}";
    }

    private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
    {
        BtnGenerate.IsEnabled = false;
        var progress = new Progress<int>(value =>
        {
            MainProgressBar.Value = value;
            StatusLabel.Text = $"Processando: {value}%";
        });

        var meetings = await _csvParsingService.GetMeetingScheduleFromCsvAsync(
            TxtInputFile.Text, 
            (int)YearComboBox.SelectedItem, 
            (int)MonthComboBox.SelectedItem, 
            progress);

        _pdfGenerationService.SetContext((int)YearComboBox.SelectedItem, (int)MonthComboBox.SelectedItem);

        if (ChkGenarateMeetingProgram.IsChecked == true)
            await _pdfGenerationService.GenerateMeetingSchedulePdfAsync(TxtOutputDirectory.Text, meetings, progress, ChkGenarateStudentDesignations.IsChecked == true ? 10 : 20);

        if (ChkGenarateStudentDesignations.IsChecked == true)
            await _pdfGenerationService.GenerateStudentDesignationsPdfAsync(TxtOutputDirectory.Text, meetings, progress, ChkGenarateMeetingProgram.IsChecked == true ? 10 : 20);

        if (ChkSendWhatsappMessage.IsChecked == true)
            await _whatsAppMessagingService.SendConfirmationRequestsAsync(meetings);

        BtnGenerate.IsEnabled = true;
    }

    private void BtnBrowseInputFile_Click(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog();

        fileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
        fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        bool? result = fileDialog.ShowDialog();

        if (result == true)
            TxtInputFile.Text = fileDialog.FileName;
    }

    private void BtnBrowseOutputDirectory_Click(object sender, RoutedEventArgs e)
    {
        var folderDialog = new OpenFolderDialog();

        folderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        bool? result = folderDialog.ShowDialog();

        if (result == true)
            TxtOutputDirectory.Text = folderDialog.FolderName;
    }
}
