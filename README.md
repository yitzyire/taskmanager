# TaskManager

`TaskManager` is a WinForms-based process inspection tool built as a darker, more technical alternative to the default Windows Task Manager.

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

- `TaskManager/Form1.cs`
  Main UI logic, refresh flow, process rendering, detail loading, and actions
- `TaskManager/Form1.Designer.cs`
  WinForms layout and control initialization
- `TaskManager/NativeProcessMethods.cs`
  Native Windows interop helpers for process and service metadata
- `TaskManager/ProcessMetadataProvider.cs`
  Metadata provider abstraction
- `TaskManager/Program.cs`
  Application entry point
- `TaskManager/TaskManager.csproj`
  Project file

## Build

From the repository root:

```powershell
dotnet build TaskManager\TaskManager.csproj -c Release
```

## Run

You can launch the built executable from the standard build output:

```text
TaskManager\bin\Release\net10.0-windows\TaskManager.exe
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

- `Disk` now shows per-process throughput derived from Windows I/O counters
- `GPU` uses the Windows `GPU Engine` performance counters, so some systems or drivers may report little or no usable data
- `Network` currently shows live per-process connection activity when byte-accurate per-process throughput is not available through the lightweight collection path used by this app
- command line hydration is asynchronous, so it may populate progressively instead of appearing instantly for every row
- some protected system processes may not expose full metadata
- because the app uses WinForms, it is not cross-platform as a UI application

## Development Notes

- the codebase has been tuned to reduce obvious refresh flicker, but future improvements could include more diff-based row updates and richer live metrics

## Next Improvements

- more accurate per-process network throughput via ETW or a similar deeper telemetry path
- deeper process ancestry and child expansion controls
- richer service details
- export / snapshot support
- more security-enrichment data for binaries

## Name

Application name:

`TaskManager`
