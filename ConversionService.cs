using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;

namespace DWGToNWCConverter;

public sealed class ConversionService
{
    public async Task<ConversionResult> RunBatchAsync(
        AppSettings settings,
        CancellationToken cancellationToken,
        IProgress<string>? progress = null)
    {
        ValidateSettings(settings);

        var outputFolder = settings.UseInputFolderAsOutput
            ? settings.InputFolder
            : settings.OutputFolder;

        Directory.CreateDirectory(outputFolder);

        var dwgFiles = Directory.GetFiles(settings.InputFolder, "*.dwg", SearchOption.TopDirectoryOnly);
        var messages = new ConcurrentQueue<string>();
        var convertedCount = 0;

        await Parallel.ForEachAsync(
            dwgFiles,
            new ParallelOptions
            {
                MaxDegreeOfParallelism = Math.Max(1, settings.MaxParallelism),
                CancellationToken = cancellationToken
            },
            async (dwg, ct) =>
            {
                var outputFile = Path.Combine(
                    outputFolder,
                    $"{Path.GetFileNameWithoutExtension(dwg)}.nwc");

                var startInfo = new ProcessStartInfo
                {
                    FileName = settings.ToolPath,
                    Arguments = $"-in \"{dwg}\" -out \"{outputFile}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(startInfo)
                    ?? throw new InvalidOperationException($"Could not start {settings.ToolPath}.");

                await process.WaitForExitAsync(ct);

                var message = process.ExitCode == 0
                    ? $"Done: {Path.GetFileName(outputFile)}"
                    : $"Failed ({process.ExitCode}): {Path.GetFileName(dwg)}";

                if (process.ExitCode == 0)
                {
                    Interlocked.Increment(ref convertedCount);
                }

                messages.Enqueue(message);
                progress?.Report(message);
            });

        return new ConversionResult
        {
            TotalFiles = dwgFiles.Length,
            ConvertedFiles = convertedCount,
            Messages = messages.ToArray()
        };
    }

    private static void ValidateSettings(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.InputFolder) || !Directory.Exists(settings.InputFolder))
        {
            throw new InvalidOperationException("Choose a valid DWG input folder.");
        }

        if (!settings.UseInputFolderAsOutput && string.IsNullOrWhiteSpace(settings.OutputFolder))
        {
            throw new InvalidOperationException("Choose a valid NWC output folder or enable same-folder output.");
        }

        if (string.IsNullOrWhiteSpace(settings.ToolPath) || !File.Exists(settings.ToolPath))
        {
            throw new InvalidOperationException("Choose a valid FiletoolsTaskRunner.exe path.");
        }
    }
}
