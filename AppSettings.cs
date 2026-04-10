namespace DWGToNWCConverter;

public sealed class AppSettings
{
    public string InputFolder { get; set; } = string.Empty;

    public string OutputFolder { get; set; } = string.Empty;

    public bool UseInputFolderAsOutput { get; set; }

    public string ToolPath { get; set; } = NavisworksPathHelper.GetPreferredToolPath();

    public int MaxParallelism { get; set; } = 4;

    public string ScheduledTaskName { get; set; } = "DWGToNWCConverterBiWeekly";

    public string ScheduledDay { get; set; } = "FRI";

    public string ScheduledTime { get; set; } = "08:00";
}
