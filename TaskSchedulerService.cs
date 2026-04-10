using System.Diagnostics;

namespace DWGToNWCConverter;

public sealed class TaskSchedulerService
{
    public async Task<ScheduledTaskInfo> GetTaskInfoAsync(string taskName)
    {
        var result = await RunSchtasksAsync($"/Query /TN \"{taskName}\" /FO LIST");
        if (result.ExitCode != 0)
        {
            return new ScheduledTaskInfo
            {
                Exists = false,
                StatusText = "Not configured"
            };
        }

        var statusLine = result.Output
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault(line => line.StartsWith("Status:", StringComparison.OrdinalIgnoreCase));

        return new ScheduledTaskInfo
        {
            Exists = true,
            StatusText = statusLine?.Split(':', 2).LastOrDefault()?.Trim() ?? "Configured"
        };
    }

    public async Task CreateOrUpdateBiWeeklyTaskAsync(
        string taskName,
        string executablePath,
        string scheduledDay,
        string scheduledTime)
    {
        var arguments =
            $"/Create /F /SC WEEKLY /MO 2 /D {scheduledDay} /ST {scheduledTime} /TN \"{taskName}\" /TR \"\\\"{executablePath}\\\" --run-scheduled\"\"";

        var result = await RunSchtasksAsync(arguments);
        EnsureSuccess(result, "create or update the scheduled task");
    }

    public async Task DeleteTaskAsync(string taskName)
    {
        var result = await RunSchtasksAsync($"/Delete /F /TN \"{taskName}\"");
        if (result.ExitCode != 0 && !result.Output.Contains("cannot find", StringComparison.OrdinalIgnoreCase))
        {
            EnsureSuccess(result, "delete the scheduled task");
        }
    }

    public async Task RunTaskNowAsync(string taskName)
    {
        var result = await RunSchtasksAsync($"/Run /TN \"{taskName}\"");
        EnsureSuccess(result, "start the scheduled task");
    }

    private static void EnsureSuccess((int ExitCode, string Output) result, string action)
    {
        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException($"Unable to {action}.{Environment.NewLine}{result.Output}");
        }
    }

    private static async Task<(int ExitCode, string Output)> RunSchtasksAsync(string arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "schtasks.exe",
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start schtasks.exe.");

        var standardOutput = await process.StandardOutput.ReadToEndAsync();
        var standardError = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        return (process.ExitCode, $"{standardOutput}{Environment.NewLine}{standardError}".Trim());
    }
}
