using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using Forms = System.Windows.Forms;

namespace DWGToNWCConverter;

public partial class MainWindow : Window
{
    private readonly ConversionService _conversionService = new();
    private readonly TaskSchedulerService _taskSchedulerService = new();
    private bool _isBusy;
    private bool _isApplyingCenteredMaximize;

    public MainWindow()
    {
        InitializeComponent();
        SourceInitialized += (_, _) => WindowThemeHelper.ApplyTheme(this);
        LoadSettingsIntoUi();
        Loaded += MainWindow_Loaded;
        StateChanged += MainWindow_StateChanged;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        CenterWindowOnCurrentScreen();
        await RefreshSchedulerStatusAsync();
    }

    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        if (_isApplyingCenteredMaximize)
        {
            return;
        }

        if (WindowState == WindowState.Maximized)
        {
            ApplyCenteredMaximize();
            return;
        }

        if (WindowState == WindowState.Normal && IsLoaded)
        {
            CenterWindowOnCurrentScreen();
        }
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

        SelectComboBoxItem(ParallelismComboBox, settings.MaxParallelism.ToString());
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

            Log($"Starting conversion for {settings.InputFolder}");
            var progress = new Progress<string>(Log);
            var result = await _conversionService.RunBatchAsync(settings, CancellationToken.None, progress);

            SummaryTextBlock.Text =
                $"Converted {result.ConvertedFiles} of {result.TotalFiles} DWG files to NWC.";
            Log($"Batch complete. {result.ConvertedFiles}/{result.TotalFiles} converted.");
        }
        catch (Exception ex)
        {
            SummaryTextBlock.Text = "Conversion failed.";
            Log(ex.Message);
            System.Windows.MessageBox.Show(this, ex.Message, "Conversion error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
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
            SummaryTextBlock.Text = "Settings saved successfully.";
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
        var parallelText = (ParallelismComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "4";
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
            MaxParallelism = int.Parse(parallelText),
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
    }

    private void Log(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        LogListBox.Items.Insert(0, line);
    }

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

    private void ApplyCenteredMaximize()
    {
        try
        {
            _isApplyingCenteredMaximize = true;

            var screen = GetCurrentScreen();
            var workArea = screen.WorkingArea;
            var targetWidth = Math.Min(MaxWidth, workArea.Width);
            var targetHeight = Math.Min(Math.Max(Height, MinHeight), workArea.Height);

            WindowState = WindowState.Normal;
            Width = targetWidth;
            Height = targetHeight;
            Left = workArea.Left + ((workArea.Width - targetWidth) / 2d);
            Top = workArea.Top + ((workArea.Height - targetHeight) / 2d);
        }
        finally
        {
            _isApplyingCenteredMaximize = false;
        }
    }

    private void CenterWindowOnCurrentScreen()
    {
        var screen = GetCurrentScreen();
        var workArea = screen.WorkingArea;
        var targetWidth = double.IsNaN(Width) ? ActualWidth : Width;
        var targetHeight = double.IsNaN(Height) ? ActualHeight : Height;

        targetWidth = Math.Min(targetWidth, workArea.Width);
        targetHeight = Math.Min(targetHeight, workArea.Height);

        Left = workArea.Left + ((workArea.Width - targetWidth) / 2d);
        Top = workArea.Top + ((workArea.Height - targetHeight) / 2d);
    }

    private Forms.Screen GetCurrentScreen()
    {
        var handle = new WindowInteropHelper(this).Handle;
        return handle != IntPtr.Zero ? Forms.Screen.FromHandle(handle) : Forms.Screen.PrimaryScreen;
    }
}
