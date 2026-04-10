# DWG to NWC Converter

<p align="center">
  <img src="favicon.ico" alt="DWG to NWC Converter icon" width="96" />
</p>

<p align="center">
  A Windows desktop tool for converting <code>.dwg</code> files to <code>.nwc</code> with Autodesk Navisworks, either on demand or through a bi-weekly Task Scheduler automation.
</p>

## <img src="docs/icons/overview.svg" alt="" width="22" valign="middle" /> Overview

This project is a WPF desktop app built on `.NET 8` for teams that need a repeatable DWG-to-NWC workflow.

The app lets you:

- choose a DWG input folder
- choose an NWC output folder, or reuse the input folder
- point the app to `FiletoolsTaskRunner.exe`
- run batch conversions immediately
- save settings locally for future runs
- create, refresh, run, and delete a Windows Scheduled Task that launches the app automatically every 2 weeks

## <img src="docs/icons/requirements.svg" alt="" width="22" valign="middle" /> Requirements

Before using the app, make sure the machine has:

- Windows
- `.NET 8` SDK or runtime for building/running the project
- Autodesk Navisworks Manage installed
- access to `FiletoolsTaskRunner.exe`

The app automatically looks for Navisworks in the default install path pattern below and prefers the newest installed version it can find between 2022 and 2030:

```text
C:\Program Files\Autodesk\Navisworks Manage <YEAR>\FiletoolsTaskRunner.exe
```

If nothing is detected, the default fallback path is:

```text
C:\Program Files\Autodesk\Navisworks Manage 2024\FiletoolsTaskRunner.exe
```

## <img src="docs/icons/start.svg" alt="" width="22" valign="middle" /> Getting Started

The main way to use this app is by downloading the packaged executable from the GitHub Releases page.

Recommended for end users:

1. Open the repository's `Releases` page on GitHub.
2. Download the latest packaged app or `.exe`.
3. Start the application on a Windows machine that has Autodesk Navisworks Manage installed.
4. Configure the folders and Navisworks tool path in the app.

If someone wants to run the project from source instead, they can:

```powershell
dotnet run
```

The project target is:

```text
net8.0-windows
```

## <img src="docs/icons/usage.svg" alt="" width="22" valign="middle" /> How To Use The App

When the app opens, the main window is split into two areas:

- `Conversion Setup`: manual batch conversion settings and run controls
- `Automation`: Task Scheduler setup for unattended runs

### Conversion Setup

Use the left-hand panel to configure the conversion:

1. Set `Folder A: DWG input` to the folder containing the `.dwg` files you want to process.
2. Decide where `.nwc` files should go:
   `Save NWC files into Folder A` writes results into the input folder.
   If that is unchecked, set `Folder B: NWC output` to a separate destination.
3. Confirm or browse to `Navisworks FiletoolsTaskRunner.exe`.
4. Choose `Parallel conversions`.
5. Click `Run Conversion Now`.

What happens during a manual run:

- the app validates the configured folders and Navisworks tool path
- it scans the input folder for `*.dwg` files in the top-level directory only
- it starts one Navisworks conversion process per DWG file
- it limits concurrency using the selected parallelism value
- it logs success or failure per file in the activity log
- it updates the summary box with the final converted count

### Helpful UI behavior

- `Save Settings` stores your current configuration locally so the next launch restores it.
- If the output folder does not exist yet, the app creates it when conversion starts.
- `Open` buttons open the selected input or output folder in Windows Explorer.
- `Clear` removes entries from the activity log in the current session.

## <img src="docs/icons/automation.svg" alt="" width="22" valign="middle" /> How The Automation Works

The automation is built on top of Windows Task Scheduler.

When you click `Create or Update Schedule`, the app:

1. saves the settings currently shown in the UI
2. finds the built application executable
3. creates or replaces a scheduled task using `schtasks.exe`
4. configures it as a `WEEKLY` task with `MO 2`, which means every 2 weeks
5. sets the selected day and time
6. launches this app with the command-line argument `--run-scheduled`

The scheduled command effectively looks like this:

```text
DWGToNWCConverter.exe --run-scheduled
```

When the app starts with `--run-scheduled`:

- it does not open the normal UI
- it loads the last saved settings from disk
- it runs the same batch conversion logic used by the manual workflow
- it shuts itself down when the batch is finished

### Automation controls

- `Create or Update Schedule`: creates the task or overwrites an existing task with the same name
- `Run Scheduled Task Now`: triggers the configured task immediately through Task Scheduler
- `Refresh Status`: queries Task Scheduler and updates the status badge in the app
- `Delete Schedule`: removes the scheduled task

### Where scheduled settings come from

Scheduled runs always use the saved settings file, not unsaved UI changes. If you change any folder, tool path, task name, day, time, or parallelism, save or recreate the schedule so the automation uses the latest values.

## <img src="docs/icons/settings.svg" alt="" width="22" valign="middle" /> Settings And Storage

The app stores settings as JSON in the current user's local app data folder:

```text
%LOCALAPPDATA%\DWGToNWCConverter\settings.json
```

The saved settings include:

- input folder
- output folder
- whether the input folder is also the output folder
- Navisworks tool path
- max parallelism
- scheduled task name
- scheduled day
- scheduled time

## <img src="docs/icons/architecture.svg" alt="" width="22" valign="middle" /> Codebase Guide

The main files are:

- [App.xaml.cs](App.xaml.cs): application startup and headless `--run-scheduled` execution path
- [MainWindow.xaml](MainWindow.xaml): WPF layout, styling, and UI structure
- [MainWindow.xaml.cs](MainWindow.xaml.cs): UI event handlers, settings binding, logging, scheduling actions, and window behavior
- [ConversionService.cs](ConversionService.cs): batch conversion engine and parallel process execution
- [TaskSchedulerService.cs](TaskSchedulerService.cs): wrapper around `schtasks.exe` for task creation, status lookup, execution, and deletion
- [SettingsService.cs](SettingsService.cs): load/save of the JSON settings file
- [AppSettings.cs](AppSettings.cs): persisted configuration model
- [NavisworksPathHelper.cs](NavisworksPathHelper.cs): Navisworks install path detection
- [WindowThemeHelper.cs](WindowThemeHelper.cs): native DWM caption and border theming for the window

### Runtime flow

Manual run:

1. user edits settings in the window
2. `MainWindow` reads the form values into `AppSettings`
3. `SettingsService.Save()` persists them
4. `ConversionService.RunBatchAsync()` converts the DWG files
5. the UI log and summary are updated

Scheduled run:

1. Windows Task Scheduler launches the app with `--run-scheduled`
2. `App.OnStartup()` detects the argument
3. `SettingsService.Load()` restores the saved settings
4. `ConversionService.RunBatchAsync()` performs the conversion without showing the window
5. the process exits

## <img src="docs/icons/troubleshooting.svg" alt="" width="22" valign="middle" /> Troubleshooting

- If conversion fails immediately, verify that the selected `FiletoolsTaskRunner.exe` path exists.
- If no files are converted, make sure the input folder contains `.dwg` files directly in that folder; subfolders are not scanned.
- If the schedule cannot be created, try running the app as Administrator. Some Windows policies restrict Task Scheduler changes.
- If the scheduler status shows `Not configured`, the task name in the app may not match the actual task name in Windows Task Scheduler.
- If output files are missing, confirm whether `Save NWC files into Folder A` is enabled; that setting changes the destination folder.

## <img src="docs/icons/notes.svg" alt="" width="22" valign="middle" /> Notes For Future Improvements

Areas that would be good next steps for the codebase:

- recursive DWG discovery for nested project folders
- richer conversion failure details by capturing standard output/error from Navisworks runs
- cancellation support for long manual batches
- stronger validation and user feedback around schedule permissions
- automated tests around settings persistence and scheduler command generation
