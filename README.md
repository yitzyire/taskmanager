# New Windows Task Manager

`New Windows Task Manager` is a WinForms-based process inspection tool built as a darker, more technical alternative to the default Windows Task Manager.

It focuses on:

- process visibility in a nested parent/child view
- quick triage of running binaries
- startup item visibility
- service visibility and control
- binary-oriented details that are useful to a security-minded user

## Current Features

- `Processes` tab with a Task Manager-style table
- nested process display in the `Name` column
- process columns for:
  - `Name`
  - `Process name`
  - `PID`
  - `CPU`
  - `Memory`
  - `GPU`
  - `Disk`
  - `Network`
  - `Status`
  - `Command line`
- fuzzy search for process names and paths
- right-click process actions:
  - `End Task`
  - `Open File Location`
- selected-process detail panel with:
  - name
  - PID
  - parent PID
  - status
  - memory
  - path
  - security information
  - command line
  - signer / certificate details when available
  - SHA-256 hash
- process state coloring:
  - new processes shown in green
  - closed processes shown in red for a short period
  - root processes shown with a distinct color
- `Startup` tab with common startup registry and folder locations
- `Services` tab with service status and start/stop actions
- `Calls` tab for internal activity logging
- dark UI styling and technical-looking layout

## Platform Notes

The project is written in C# using WinForms and targets `net10.0-windows`.

The UI itself is Windows-only because WinForms is Windows-only.

Some metadata collection is abstracted behind provider classes, but process metadata and several low-level features currently use Windows-native APIs for best results on Windows.

## Project Layout

- `NewWindowsTaskManager/Form1.cs`
  Main UI logic, refresh flow, process rendering, detail loading, and actions
- `NewWindowsTaskManager/Form1.Designer.cs`
  WinForms layout and control initialization
- `NewWindowsTaskManager/NativeProcessMethods.cs`
  Native Windows interop helpers for process and service metadata
- `NewWindowsTaskManager/ProcessMetadataProvider.cs`
  Metadata provider abstraction
- `NewWindowsTaskManager/Program.cs`
  Application entry point
- `NewWindowsTaskManager/NewWindowsTaskManager.csproj`
  Project file

## Build

From the repository root:

```powershell
dotnet build NewWindowsTaskManager\NewWindowsTaskManager.csproj -c Release
```

If you want to publish the output to the shared deploy folder used in this repo:

```powershell
dotnet build NewWindowsTaskManager\NewWindowsTaskManager.csproj -c Release -p:OutDir=artifacts\deploy\
```

## Run

You can launch the built executable from:

```text
artifacts\deploy\NewWindowsTaskManager.exe
```

Or from the standard build output:

```text
NewWindowsTaskManager\bin\Release\net10.0-windows\NewWindowsTaskManager.exe
```

## How It Works

### Process Refresh

- the app periodically reads the process list
- parent/child relationships are built into a nested display
- rows are rendered through a virtualized backing model to reduce visible refresh churn
- selected process details are loaded separately from the main refresh loop

### Command Line Collection

- command lines are not collected by PowerShell
- command lines are retrieved through the Windows native process metadata path
- the selected process is resolved immediately when needed
- visible rows are also hydrated in the background so the `Command line` column fills in without requiring exact clicks on every row

### Security Details

When a process is inspected, the tool attempts to collect:

- binary path
- file metadata
- SHA-256 hash
- signature status
- signer subject
- signer issuer
- certificate thumbprint
- parent PID
- file size
- last write time

## Known Limitations

- `GPU`, `Disk`, and `Network` columns are currently placeholders and do not yet show real per-process metrics
- command line hydration is asynchronous, so it may populate progressively instead of appearing instantly for every row
- some protected system processes may not expose full metadata
- because the app uses WinForms, it is not cross-platform as a UI application
- if the deployed executable is still running, rebuilding directly into `artifacts\deploy\` can fail because Windows locks the `.exe`

## Development Notes

- avoid rebuilding into `artifacts\deploy\` while the deployed executable is open
- if the app is running from the deploy folder, close it before overwriting that executable
- the codebase has been tuned to reduce obvious refresh flicker, but future improvements could include more diff-based row updates and richer live metrics

## Next Improvements

- real `GPU`, `Disk`, and `Network` usage collection
- deeper process ancestry and child expansion controls
- richer service details
- export / snapshot support
- more security-enrichment data for binaries

## Name

Application name:

`New Windows Task Manager`
