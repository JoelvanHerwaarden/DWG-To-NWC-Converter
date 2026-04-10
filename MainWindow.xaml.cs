using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;

namespace DWGToNWCConverter;

public partial class MainWindow : Window
{
    private readonly ConversionService _conversionService = new();
    private readonly TaskSchedulerService _taskSchedulerService = new();
    private bool _isBusy;
    private CancellationTokenSource? _runCancellationTokenSource;
    private int _currentBatchTotal;
    private int _currentBatchCompleted;
    private int _currentBatchFailed;

    public MainWindow()
    {
        InitializeComponent();
        SourceInitialized += (_, _) => WindowThemeHelper.ApplyTheme(this);
        LoadSettingsIntoUi();
        Loaded += async (_, _) => await RefreshSchedulerStatusAsync();
    }

    private void LoadSettingsIntoUi()
    {
        var settings = SettingsService.Load();

        InputFolderTextBox.Text = settings.InputFolder;
        OutputFolderTextBox.Text = settings.OutputFolder;
        ToolPathTextBox.Text = string.IsNullOrWhiteSpace(settings.ToolPath)
            ? NavisworksPathHelper.GetPreferredToolPath()
            : settings.ToolPath;
        SameAsInputCheckBox.IsChecked = settings.UseInputFolderAsOutput;
        TaskNameTextBox.Text = settings.ScheduledTaskName;
        ScheduledTimeTextBox.Text = settings.ScheduledTime;
        SelectComboBoxItem(ScheduledDayComboBox, settings.ScheduledDay);
        ApplyOutputFolderState();
    }

    private async void RunNowButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        try
        {
            _isBusy = true;
            ToggleBusyState();

            var settings = ReadSettingsFromUi();
            SettingsService.Save(settings);

            _runCancellationTokenSource = new CancellationTokenSource();
            ResetProgressState();
            Log($"Starting conversion for {settings.InputFolder}");
            var progress = new Progress<string>(HandleConversionProgress);
            var result = await _conversionService.RunBatchAsync(settings, _runCancellationTokenSource.Token, progress);

            SetProgressState(result.TotalFiles, result.ConvertedFiles + _currentBatchFailed);
            ProgressTextBlock.Text =
                result.TotalFiles == 0
                    ? "No DWG files were found."
                    : $"Finished. {result.ConvertedFiles} converted, {_currentBatchFailed} failed.";
            Log($"Batch complete. {result.ConvertedFiles}/{result.TotalFiles} converted.");
        }
        catch (OperationCanceledException)
        {
            ProgressTextBlock.Text = "Conversion cancelled.";
            Log("Conversion cancelled.");
        }
        catch (Exception ex)
        {
            ProgressTextBlock.Text = "Conversion failed.";
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Conversion error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _runCancellationTokenSource?.Dispose();
            _runCancellationTokenSource = null;
            _isBusy = false;
            ToggleBusyState();
        }
    }

    private void SaveSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = ReadSettingsFromUi();
            SettingsService.Save(settings);
            Log("Settings saved.");
        }
        catch (Exception ex)
        {
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Settings error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void CreateScheduleButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = ReadSettingsFromUi();
            SettingsService.Save(settings);

            var executablePath = Path.Combine(
                AppContext.BaseDirectory,
                $"{typeof(MainWindow).Assembly.GetName().Name}.exe");

            if (!File.Exists(executablePath))
            {
                throw new InvalidOperationException($"Unable to find the application executable at '{executablePath}'.");
            }

            await _taskSchedulerService.CreateOrUpdateBiWeeklyTaskAsync(
                settings.ScheduledTaskName,
                executablePath,
                settings.ScheduledDay,
                settings.ScheduledTime);

            Log($"Scheduled task '{settings.ScheduledTaskName}' created or updated.");
            await RefreshSchedulerStatusAsync();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Schedule error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void RunScheduledTaskButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = ReadSettingsFromUi();
            await _taskSchedulerService.RunTaskNowAsync(settings.ScheduledTaskName);
            Log($"Scheduled task '{settings.ScheduledTaskName}' started.");
            await RefreshSchedulerStatusAsync();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Task Scheduler error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void DeleteScheduleButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = ReadSettingsFromUi();
            await _taskSchedulerService.DeleteTaskAsync(settings.ScheduledTaskName);
            Log($"Scheduled task '{settings.ScheduledTaskName}' deleted.");
            await RefreshSchedulerStatusAsync();
        }
        catch (Exception ex)
        {
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Task Scheduler error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void RefreshSchedulerButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshSchedulerStatusAsync();
    }

    private async Task RefreshSchedulerStatusAsync()
    {
        var taskName = string.IsNullOrWhiteSpace(TaskNameTextBox.Text)
            ? SettingsService.Load().ScheduledTaskName
            : TaskNameTextBox.Text.Trim();

        var info = await _taskSchedulerService.GetTaskInfoAsync(taskName);
        SchedulerStatusText.Text = info.Exists ? info.StatusText : "Not configured";
        SchedulerStatusText.Foreground = !info.Exists
            ? (System.Windows.Media.Brush)FindResource("DangerBrush")
            : IsReadyStatus(info.StatusText)
                ? (System.Windows.Media.Brush)FindResource("ReadyBrush")
                : (System.Windows.Media.Brush)FindResource("AccentBrush");
    }

    private void BrowseInputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = BrowseForFolder(InputFolderTextBox.Text);
        if (folder is null)
        {
            return;
        }

        InputFolderTextBox.Text = folder;
        if (SameAsInputCheckBox.IsChecked == true)
        {
            OutputFolderTextBox.Text = folder;
        }
    }

    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = BrowseForFolder(OutputFolderTextBox.Text);
        if (folder is not null)
        {
            OutputFolderTextBox.Text = folder;
        }
    }

    private void OpenInputFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolderInExplorer(InputFolderTextBox.Text);
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolderInExplorer(OutputFolderTextBox.Text);
    }

    private void BrowseToolPath_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            FileName = ToolPathTextBox.Text
        };

        if (dialog.ShowDialog(this) == true)
        {
            ToolPathTextBox.Text = dialog.FileName;
        }
    }

    private void SameAsInputCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        ApplyOutputFolderState();
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        LogListBox.Items.Clear();
    }

    private void CopyAllLogButton_Click(object sender, RoutedEventArgs e)
    {
        var logText = GetAllLogText();
        if (string.IsNullOrWhiteSpace(logText))
        {
            Log("Copy log skipped. No log entries available.");
            return;
        }

        System.Windows.Clipboard.SetText(logText);
        Log("Copied all activity log entries to the clipboard.");
    }

    private void CloseAppButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ApplyOutputFolderState()
    {
        var useInputFolder = SameAsInputCheckBox.IsChecked == true;
        OutputFolderTextBox.IsEnabled = !useInputFolder;
        BrowseOutputButton.IsEnabled = !useInputFolder;
        OpenOutputButton.IsEnabled = true;

        if (useInputFolder && !string.IsNullOrWhiteSpace(InputFolderTextBox.Text))
        {
            OutputFolderTextBox.Text = InputFolderTextBox.Text;
        }
    }

    private AppSettings ReadSettingsFromUi()
    {
        var scheduledDay = (ScheduledDayComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "FRI";

        if (!TimeOnly.TryParse(ScheduledTimeTextBox.Text, out var scheduledTime))
        {
            throw new InvalidOperationException("Enter the schedule time in 24-hour format, for example 08:00.");
        }

        return new AppSettings
        {
            InputFolder = InputFolderTextBox.Text.Trim(),
            OutputFolder = OutputFolderTextBox.Text.Trim(),
            UseInputFolderAsOutput = SameAsInputCheckBox.IsChecked == true,
            ToolPath = ToolPathTextBox.Text.Trim(),
            MaxParallelism = 1,
            ScheduledTaskName = string.IsNullOrWhiteSpace(TaskNameTextBox.Text)
                ? "DWGToNWCConverterBiWeekly"
                : TaskNameTextBox.Text.Trim(),
            ScheduledDay = scheduledDay,
            ScheduledTime = scheduledTime.ToString("HH:mm")
        };
    }

    private static void SelectComboBoxItem(System.Windows.Controls.ComboBox comboBox, string value)
    {
        foreach (var item in comboBox.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Content?.ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }

        if (comboBox.Items.Count > 0)
        {
            comboBox.SelectedIndex = 0;
        }
    }

    private void ToggleBusyState()
    {
        RunNowButton.IsEnabled = !_isBusy;
        RunNowButton.Content = _isBusy ? "Running..." : "Run Conversion Now";
        CancelRunButton.IsEnabled = _isBusy;
    }

    private void CancelRunButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isBusy || _runCancellationTokenSource is null || _runCancellationTokenSource.IsCancellationRequested)
        {
            return;
        }

        _runCancellationTokenSource.Cancel();
        ProgressTextBlock.Text = "Cancelling conversion...";
        Log("Cancellation requested.");
    }

    private void ResetProgressState()
    {
        _currentBatchTotal = 0;
        _currentBatchCompleted = 0;
        _currentBatchFailed = 0;
        ConversionProgressBar.Maximum = 1;
        ConversionProgressBar.Value = 0;
        ProgressTextBlock.Text = "Preparing conversion batch...";
    }

    private void SetProgressState(int totalFiles, int processedFiles)
    {
        _currentBatchTotal = Math.Max(0, totalFiles);
        ConversionProgressBar.Maximum = Math.Max(1, _currentBatchTotal);
        ConversionProgressBar.Value = Math.Min(processedFiles, ConversionProgressBar.Maximum);
    }

    private void HandleConversionProgress(string message)
    {
        if (message.StartsWith("Found ", StringComparison.OrdinalIgnoreCase))
        {
            _currentBatchTotal = ExtractFirstInteger(message);
            SetProgressState(_currentBatchTotal, 0);
            ProgressTextBlock.Text = _currentBatchTotal == 1
                ? "Found 1 DWG file. Starting conversion..."
                : $"Found {_currentBatchTotal} DWG files. Starting conversion...";
            return;
        }

        if (message.StartsWith("Done:", StringComparison.OrdinalIgnoreCase))
        {
            Log(message);
            return;
        }

        if (message.StartsWith("Failed", StringComparison.OrdinalIgnoreCase))
        {
            _currentBatchFailed++;
            SetProgressState(_currentBatchTotal, _currentBatchCompleted + _currentBatchFailed);
            ProgressTextBlock.Text = BuildProgressText();
            Log(message);
            return;
        }

        if (message.StartsWith("Detected:", StringComparison.OrdinalIgnoreCase))
        {
            _currentBatchCompleted++;
            SetProgressState(_currentBatchTotal, _currentBatchCompleted + _currentBatchFailed);
            ProgressTextBlock.Text = BuildProgressText();
            Log(message);
            return;
        }

        if (message.StartsWith("Workspace:", StringComparison.OrdinalIgnoreCase) ||
            message.StartsWith("Starting:", StringComparison.OrdinalIgnoreCase))
        {
            Log(message);
            return;
        }

        if (message.StartsWith("No DWG files", StringComparison.OrdinalIgnoreCase) ||
            message.StartsWith("Conversion finished with missing outputs", StringComparison.OrdinalIgnoreCase))
        {
            ProgressTextBlock.Text = message;
            Log(message);
            return;
        }

        if (message.StartsWith("Temp conversion check:", StringComparison.OrdinalIgnoreCase) ||
            message.StartsWith("Copied ", StringComparison.OrdinalIgnoreCase))
        {
            Log(message);
        }
    }

    private string BuildProgressText()
    {
        if (_currentBatchTotal <= 0)
        {
            return "Conversion in progress...";
        }

        return $"Detected {_currentBatchCompleted} of {_currentBatchTotal} output files. " +
               $"{_currentBatchFailed} missing or failed.";
    }

    private static int ExtractFirstInteger(string message)
    {
        var digits = new string(message.SkipWhile(c => !char.IsDigit(c))
            .TakeWhile(char.IsDigit)
            .ToArray());

        return int.TryParse(digits, out var value) ? value : 0;
    }

    private void Log(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        LogListBox.Items.Insert(0, line);
    }

    private string GetAllLogText() =>
        string.Join(Environment.NewLine, LogListBox.Items.Cast<object>().Select(item => item?.ToString()));

    private static string? BrowseForFolder(string currentPath)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            InitialDirectory = Directory.Exists(currentPath) ? currentPath : Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            ShowNewFolderButton = true
        };

        return dialog.ShowDialog() == Forms.DialogResult.OK ? dialog.SelectedPath : null;
    }

    private void OpenFolderInExplorer(string folderPath)
    {
        var trimmedPath = folderPath.Trim();
        if (!Directory.Exists(trimmedPath))
        {
            System.Windows.MessageBox.Show(
                this,
                "The selected folder does not exist yet.",
                "Open folder",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{trimmedPath}\"",
            UseShellExecute = true
        });
    }

    private static bool IsReadyStatus(string statusText) =>
        string.Equals(statusText?.Trim(), "Ready", StringComparison.OrdinalIgnoreCase);
}
