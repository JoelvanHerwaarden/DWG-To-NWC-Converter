using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;

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

        if (dwgFiles.Length == 0)
        {
            var noFilesMessage = "No DWG files were found in the selected input folder.";
            messages.Enqueue(noFilesMessage);
            progress?.Report(noFilesMessage);

            return new ConversionResult
            {
                TotalFiles = 0,
                ConvertedFiles = 0,
                Messages = messages.ToArray()
            };
        }

        var tempWorkspacePath = CreateTempWorkspacePath();
        var tempInputFolder = Path.Combine(tempWorkspacePath, "input");
        var tempOutputFolder = Path.Combine(tempWorkspacePath, "output");

        try
        {
            Directory.CreateDirectory(tempInputFolder);
            Directory.CreateDirectory(tempOutputFolder);

            Report($"Workspace: {tempWorkspacePath}", messages, progress);
            Report(
                $"Found {dwgFiles.Length} DWG files. Copying them to temporary workspace: {tempWorkspacePath}",
                messages,
                progress);

            var stagedFiles = await StageInputFilesAsync(dwgFiles, tempInputFolder, cancellationToken);

            Report(
                $"Starting one Navisworks batch conversion in temp workspace for {stagedFiles.Length} files.",
                messages,
                progress);

            await ConvertStagedFilesAsync(
                settings.ToolPath,
                stagedFiles,
                tempOutputFolder,
                cancellationToken,
                messages,
                progress);

            var stagedNwcFiles = GetConvertedNwcFiles(stagedFiles, tempOutputFolder);
            Report(
                $"Temp conversion check: created {stagedNwcFiles.Length} NWC files for {stagedFiles.Length} DWG files.",
                messages,
                progress);

            var copiedCount = await CopyConvertedFilesToOutputAsync(stagedNwcFiles, outputFolder, cancellationToken);
            Report(
                $"Copied {copiedCount} NWC files from temp workspace to output folder.",
                messages,
                progress);

            if (stagedNwcFiles.Length != stagedFiles.Length)
            {
                Report(
                    "Conversion finished with missing outputs. Review the log entries for the DWG files that did not produce an NWC file.",
                    messages,
                    progress);
            }

            return new ConversionResult
            {
                TotalFiles = stagedFiles.Length,
                ConvertedFiles = copiedCount,
                Messages = messages.ToArray()
            };
        }
        finally
        {
            TryDeleteDirectory(tempWorkspacePath);
        }
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

    private static async Task<string[]> StageInputFilesAsync(
        IReadOnlyList<string> sourceFiles,
        string tempInputFolder,
        CancellationToken cancellationToken)
    {
        var stagedFiles = new List<string>(sourceFiles.Count);

        foreach (var sourceFile in sourceFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var stagedPath = Path.Combine(tempInputFolder, Path.GetFileName(sourceFile));
            File.Copy(sourceFile, stagedPath, overwrite: true);
            stagedFiles.Add(stagedPath);

            await Task.Yield();
        }

        return stagedFiles.ToArray();
    }

    private static async Task ConvertStagedFilesAsync(
        string toolPath,
        IReadOnlyList<string> stagedDwgFiles,
        string tempOutputFolder,
        CancellationToken cancellationToken,
        ConcurrentQueue<string> messages,
        IProgress<string>? progress)
    {
        var inputListPath = Path.Combine(tempOutputFolder, "batch-input.txt");
        var tempInputFolder = Path.GetDirectoryName(stagedDwgFiles[0]) ?? tempOutputFolder;

        try
        {
            await File.WriteAllLinesAsync(
                inputListPath,
                stagedDwgFiles,
                cancellationToken);

            Report(
                $"Starting: batch conversion for {stagedDwgFiles.Count} files.",
                messages,
                progress);

            var startInfo = CreateStartInfo(toolPath, inputListPath, tempOutputFolder);
            using var process = StartConversionProcess(startInfo);
            using var watcherCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var watcherTask = MonitorOutputFilesAsync(
                stagedDwgFiles,
                tempInputFolder,
                tempOutputFolder,
                messages,
                progress,
                watcherCancellation.Token);
            var execution = await WaitForCompletionAsync(process, cancellationToken);
            watcherCancellation.Cancel();
            await watcherTask;

            foreach (var message in BuildBatchExecutionMessages(stagedDwgFiles, tempOutputFolder, execution))
            {
                Report(message, messages, progress);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
        {
            Report($"Failed: Batch conversion | {ex.Message}", messages, progress);
        }
        finally
        {
            TryDeleteFile(inputListPath);
        }
    }

    private static async Task MonitorOutputFilesAsync(
        IReadOnlyList<string> stagedDwgFiles,
        string tempInputFolder,
        string tempOutputFolder,
        ConcurrentQueue<string> messages,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var knownOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                ReportNewOutputs(stagedDwgFiles, tempInputFolder, tempOutputFolder, knownOutputs, messages, progress);
                await Task.Delay(500, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            ReportNewOutputs(stagedDwgFiles, tempInputFolder, tempOutputFolder, knownOutputs, messages, progress);
        }
    }

    private static void ReportNewOutputs(
        IReadOnlyList<string> stagedDwgFiles,
        string tempInputFolder,
        string tempOutputFolder,
        HashSet<string> knownOutputs,
        ConcurrentQueue<string> messages,
        IProgress<string>? progress)
    {
        foreach (var stagedDwgFilePath in stagedDwgFiles)
        {
            var outputFilePath = ResolveActualOutputFilePath(
                BuildOutputFilePath(stagedDwgFilePath, tempInputFolder),
                BuildOutputFilePath(stagedDwgFilePath, tempOutputFolder));

            if (outputFilePath is null || !knownOutputs.Add(outputFilePath))
            {
                continue;
            }

            Report($"Detected: {Path.GetFileName(outputFilePath)}", messages, progress);
        }
    }

    private static string BuildOutputFilePath(string dwgFilePath, string outputFolder) =>
        Path.Combine(outputFolder, $"{Path.GetFileNameWithoutExtension(dwgFilePath)}.nwc");

    private static string[] GetConvertedNwcFiles(
        IReadOnlyList<string> stagedDwgFiles,
        string tempOutputFolder)
    {
        var convertedFiles = new List<string>(stagedDwgFiles.Count);

        foreach (var stagedDwgFilePath in stagedDwgFiles)
        {
            var stagedInputNwcFilePath = BuildOutputFilePath(
                stagedDwgFilePath,
                Path.GetDirectoryName(stagedDwgFilePath) ?? tempOutputFolder);
            var stagedOutputNwcFilePath = BuildOutputFilePath(stagedDwgFilePath, tempOutputFolder);
            var actualOutputFilePath = ResolveActualOutputFilePath(stagedInputNwcFilePath, stagedOutputNwcFilePath);

            if (actualOutputFilePath is not null)
            {
                convertedFiles.Add(actualOutputFilePath);
            }
        }

        return convertedFiles.ToArray();
    }

    private static async Task<int> CopyConvertedFilesToOutputAsync(
        IReadOnlyList<string> stagedNwcFiles,
        string outputFolder,
        CancellationToken cancellationToken)
    {
        var copiedCount = 0;

        foreach (var stagedNwcFilePath in stagedNwcFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var destinationFilePath = Path.Combine(outputFolder, Path.GetFileName(stagedNwcFilePath));
            File.Copy(stagedNwcFilePath, destinationFilePath, overwrite: true);
            copiedCount++;

            await Task.Yield();
        }

        return copiedCount;
    }

    private static ProcessStartInfo CreateStartInfo(string toolPath, string inputListPath, string outputFolder)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = toolPath,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = Path.GetDirectoryName(toolPath) ?? AppContext.BaseDirectory
        };

        startInfo.ArgumentList.Add("/i");
        startInfo.ArgumentList.Add(inputListPath);
        startInfo.ArgumentList.Add("/od");
        startInfo.ArgumentList.Add(outputFolder);
        startInfo.ArgumentList.Add("/over");
        startInfo.ArgumentList.Add("/lang");
        startInfo.ArgumentList.Add("en-US");

        return startInfo;
    }

    private static Process StartConversionProcess(ProcessStartInfo startInfo) =>
        Process.Start(startInfo)
        ?? throw new InvalidOperationException($"Could not start '{startInfo.FileName}'.");

    private static async Task<ProcessExecutionResult> WaitForCompletionAsync(
        Process process,
        CancellationToken cancellationToken)
    {
        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        await process.WaitForExitAsync(cancellationToken);
        var standardOutput = await standardOutputTask;
        var standardError = await standardErrorTask;

        return new ProcessExecutionResult(
            process.ExitCode,
            standardOutput.Trim(),
            standardError.Trim());
    }

    private static List<string> BuildBatchExecutionMessages(
        IReadOnlyList<string> stagedDwgFiles,
        string tempOutputFolder,
        ProcessExecutionResult execution)
    {
        var messages = new List<string>();

        if (!string.IsNullOrWhiteSpace(execution.StandardOutput))
        {
            messages.Add($"Output [batch]: {execution.StandardOutput}");
        }

        if (!string.IsNullOrWhiteSpace(execution.StandardError))
        {
            messages.Add($"Error [batch]: {execution.StandardError}");
        }

        foreach (var stagedDwgFilePath in stagedDwgFiles)
        {
            var stagedOutputFilePath = BuildOutputFilePath(stagedDwgFilePath, tempOutputFolder);
            var stagedInputNwcFilePath = BuildOutputFilePath(
                stagedDwgFilePath,
                Path.GetDirectoryName(stagedDwgFilePath) ?? tempOutputFolder);
            var actualOutputFilePath = ResolveActualOutputFilePath(stagedInputNwcFilePath, stagedOutputFilePath);
            var sourceFileName = Path.GetFileName(stagedDwgFilePath);

            if (actualOutputFilePath is not null)
            {
                messages.Add($"Done: {Path.GetFileName(actualOutputFilePath)}");
                continue;
            }

            var failureReason = new StringBuilder()
                .Append($"Failed ({execution.ExitCode}): {sourceFileName}")
                .Append(" | Output file was not created.")
                .Append(ToolPrintedUsage(execution) ? " | FileToolsTaskRunner rejected the command-line arguments." : string.Empty)
                .ToString();

            messages.Add(failureReason);
        }

        return messages;
    }

    private static bool ToolPrintedUsage(ProcessExecutionResult execution) =>
        execution.StandardOutput.Contains("Usage:", StringComparison.OrdinalIgnoreCase) ||
        execution.StandardError.Contains("Usage:", StringComparison.OrdinalIgnoreCase);

    private static string? ResolveActualOutputFilePath(
        string stagedInputNwcFilePath,
        string stagedOutputNwcFilePath)
    {
        if (File.Exists(stagedInputNwcFilePath))
        {
            return stagedInputNwcFilePath;
        }

        if (File.Exists(stagedOutputNwcFilePath))
        {
            return stagedOutputNwcFilePath;
        }

        return null;
    }

    private static string CreateTempWorkspacePath()
    {
        var commonAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        var tempRoot = Path.Combine(commonAppData, "DWGToNWCConverter", "ConversionTemp");

        Directory.CreateDirectory(tempRoot);
        return Path.Combine(tempRoot, Guid.NewGuid().ToString("N"));
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
            // Cleanup should not fail the conversion after the result was already determined.
        }
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Cleanup should not fail the conversion after the result was already determined.
        }
    }

    private static void Report(
        string message,
        ConcurrentQueue<string> messages,
        IProgress<string>? progress)
    {
        messages.Enqueue(message);
        progress?.Report(message);
    }

    private sealed record ProcessExecutionResult(
        int ExitCode,
        string StandardOutput,
        string StandardError);
}
