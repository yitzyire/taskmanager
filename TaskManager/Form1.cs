using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;

namespace TaskManager;

public partial class Form1 : Form
{
    private readonly Dictionary<int, ProcessNodeState> processStates = [];
    private readonly Dictionary<string, BinarySecurityDetails> binaryDetailsCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, string> processCommandLineCache = [];
    private readonly List<ProcessRowEntry> processRows = [];
    private readonly HashSet<int> collapsedProcessIds = [];
    private readonly ProcessRuntimeMetricsProvider processRuntimeMetricsProvider = new();
    private readonly SemaphoreSlim commandLineLoadLimiter = new(2, 2);
    private readonly Dictionary<int, Task<string>> commandLineLoadTasks = [];
    private readonly object commandLineLoadSync = new();
    private readonly System.Windows.Forms.Timer refreshTimer = new() { Interval = 4000 };
    private readonly DateTime appStartedUtc = DateTime.UtcNow;
    private readonly string processNameColumnName = "ProcessNameColumn";
    private readonly string processCommandLineColumnName = "CommandLineColumn";
    private bool isRefreshingProcesses;
    private bool isRefreshingStartup;
    private bool isRefreshingServices;
    private bool isRebuildingProcessGrid;
    private int? lastRenderedDetailsPid;

    public Form1()
    {
        InitializeComponent();
        Text = "TaskManager";
        ApplyWindowIcon();
        ApplyWindowChrome();
        Load += (_, _) => ConfigureProcessSplit();
        Resize += (_, _) => ConfigureProcessSplit();
        mainTabControl.SelectedIndexChanged += mainTabControl_SelectedIndexChanged;
        processGrid.SelectionChanged += processGrid_SelectionChanged;
        processGrid.CellMouseDown += processGrid_CellMouseDown;
        processGrid.CellMouseClick += processGrid_CellMouseClick;
        processGrid.CellFormatting += processGrid_CellFormatting;
        processGrid.CellValueNeeded += processGrid_CellValueNeeded;
        processGrid.Scroll += processGrid_Scroll;
        processGrid.VirtualMode = true;
        var commandLineColumn = processGrid.Columns[processCommandLineColumnName];
        var processNameColumn = processGrid.Columns[processNameColumnName];
        if (commandLineColumn is not null)
        {
            commandLineColumn.DefaultCellStyle.Font = new Font("Consolas", 8.75F);
            commandLineColumn.DefaultCellStyle.ForeColor = Color.FromArgb(148, 163, 184);
            commandLineColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            commandLineColumn.MinimumWidth = 360;
            commandLineColumn.FillWeight = 100F;
        }

        if (processNameColumn is not null)
        {
            processNameColumn.DefaultCellStyle.ForeColor = Color.FromArgb(148, 163, 184);
        }

        EnableDoubleBuffering(processGrid);
        refreshTimer.Tick += async (_, _) => await RefreshProcessesAsync();
        refreshTimer.Start();
        Shown += async (_, _) =>
        {
            await BeginInitialLoadAsync();
        };
        UpdateTabButtons();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            components?.Dispose();
            refreshTimer.Dispose();
            commandLineLoadLimiter.Dispose();
            processRuntimeMetricsProvider.Dispose();
        }

        base.Dispose(disposing);
    }

    private void ApplyWindowIcon()
    {
        try
        {
            var taskManagerPath = Path.Combine(Environment.SystemDirectory, "Taskmgr.exe");
            Icon = File.Exists(taskManagerPath)
                ? Icon.ExtractAssociatedIcon(taskManagerPath)
                : SystemIcons.Application;
        }
        catch
        {
            Icon = SystemIcons.Application;
        }
    }

    private void ApplyWindowChrome()
    {
        mainMenuStrip.Renderer = new DarkMenuRenderer();

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return;
        }

        HandleCreated += (_, _) =>
        {
            try
            {
                const int immersiveDarkModeAttribute = 20;
                const int immersiveDarkModeAttributeLegacy = 19;
                var enabled = 1;

                var result = DwmSetWindowAttribute(Handle, immersiveDarkModeAttribute, ref enabled, sizeof(int));
                if (result != 0)
                {
                    DwmSetWindowAttribute(Handle, immersiveDarkModeAttributeLegacy, ref enabled, sizeof(int));
                }
            }
            catch
            {
            }
        };
    }

    private void ConfigureProcessSplit()
    {
        if (!IsHandleCreated || processesSplitContainer.Width <= 0)
        {
            return;
        }

        const int desiredRightPanelWidth = 420;
        const int minLeftPanelWidth = 620;
        const int minRightPanelWidth = 360;

        var availableWidth = processesSplitContainer.Width - processesSplitContainer.SplitterWidth;
        if (availableWidth <= 0)
        {
            return;
        }

        var desiredRightWidth = Math.Min(desiredRightPanelWidth, Math.Max(220, availableWidth / 3));
        var maxRightWidth = Math.Max(220, availableWidth - 220);
        var safeRightWidth = Math.Min(desiredRightWidth, maxRightWidth);
        var safeSplitterDistance = Math.Max(220, availableWidth - safeRightWidth);

        var safeLeftMin = Math.Min(minLeftPanelWidth, Math.Max(0, safeSplitterDistance));
        var safeRightMin = Math.Min(minRightPanelWidth, Math.Max(0, availableWidth - safeSplitterDistance));

        var currentDistance = processesSplitContainer.SplitterDistance;
        var clampedCurrentDistance = Math.Max(safeLeftMin, Math.Min(currentDistance, availableWidth - safeRightMin));
        if (clampedCurrentDistance > 0 && clampedCurrentDistance < availableWidth)
        {
            processesSplitContainer.SplitterDistance = clampedCurrentDistance;
        }

        processesSplitContainer.Panel1MinSize = safeLeftMin;
        processesSplitContainer.Panel2MinSize = safeRightMin;

        var preferredDistance = Math.Max(safeLeftMin, Math.Min(safeSplitterDistance, availableWidth - safeRightMin));
        if (preferredDistance > 0 && preferredDistance < availableWidth && processesSplitContainer.SplitterDistance != preferredDistance)
        {
            processesSplitContainer.SplitterDistance = preferredDistance;
        }
    }

    private async void refreshProcessesButton_Click(object? sender, EventArgs e) => await RefreshProcessesAsync();

    private void endTaskButton_Click(object? sender, EventArgs e) => EndSelectedProcess();

    private async void refreshStartupButton_Click(object? sender, EventArgs e) => await RefreshStartupItemsAsync();

    private async void refreshServicesButton_Click(object? sender, EventArgs e) => await RefreshServicesAsync();

    private async void openFileLocationToolStripMenuItem_Click(object? sender, EventArgs e) => await OpenSelectedProcessLocationAsync();

    private void processSearchTextBox_TextChanged(object? sender, EventArgs e) => RenderProcessTree();

    private void processesTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = processesTabPage;

    private void startupTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = startupTabPage;

    private void servicesTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = servicesTabPage;

    private void callsTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = callsTabPage;

    private void mainTabControl_SelectedIndexChanged(object? sender, EventArgs e) => UpdateTabButtons();

    private void clearCallsButton_Click(object? sender, EventArgs e)
    {
        callsGrid.Rows.Clear();
        UpdateCallsStatus();
    }

    private void startServiceButton_Click(object? sender, EventArgs e) => ChangeSelectedService(ServiceAction.Start);

    private void stopServiceButton_Click(object? sender, EventArgs e) => ChangeSelectedService(ServiceAction.Stop);

    private async Task BeginInitialLoadAsync()
    {
        processStatusLabel.Text = "Loading processes...";
        startupStatusLabel.Text = "Loading startup items...";
        servicesStatusLabel.Text = "Loading services...";
        LogCall("Application", "Launch", "Initialized dark-mode task manager shell.");

        var processTask = RefreshProcessesAsync();
        var startupTask = RefreshStartupItemsAsync();
        var servicesTask = RefreshServicesAsync();
        await Task.WhenAll(processTask, startupTask, servicesTask);
    }

    private async Task RefreshProcessesAsync()
    {
        if (isRefreshingProcesses)
        {
            return;
        }

        isRefreshingProcesses = true;
        refreshProcessesButton.Enabled = false;

        try
        {
            var nowUtc = DateTime.UtcNow;
            var currentSnapshot = await Task.Run(() => ProcessSnapshotReader.ReadProcesses(processRuntimeMetricsProvider));
            var livePids = currentSnapshot.Keys.ToHashSet();

            foreach (var snapshot in currentSnapshot.Values)
            {
                if (!processStates.TryGetValue(snapshot.ProcessId, out var state))
                {
                    var createdState = new ProcessNodeState(snapshot, nowUtc, nowUtc > appStartedUtc.AddSeconds(2));
                    if (processCommandLineCache.TryGetValue(snapshot.ProcessId, out var cachedCommandLine))
                    {
                        createdState.CommandLine = cachedCommandLine;
                    }

                    processStates[snapshot.ProcessId] = createdState;
                    continue;
                }

                state.Update(snapshot, nowUtc);
                if (processCommandLineCache.TryGetValue(snapshot.ProcessId, out var cachedCommandLineForState))
                {
                    state.CommandLine = cachedCommandLineForState;
                }
            }

            foreach (var state in processStates.Values.Where(state => !state.IsExited && !livePids.Contains(state.ProcessId)))
            {
                state.MarkExited(nowUtc);
            }

            var stalePids = processStates.Values
                .Where(state => state.IsExited && state.ExitedAtUtc.HasValue && nowUtc - state.ExitedAtUtc.Value > TimeSpan.FromSeconds(10))
                .Select(state => state.ProcessId)
                .ToArray();

            foreach (var pid in stalePids)
            {
                processStates.Remove(pid);
                collapsedProcessIds.Remove(pid);
            }

            RenderProcessTree();
            processStatusLabel.Text = $"Processes: {currentSnapshot.Count} live, {processStates.Values.Count(state => state.IsExited)} recently closed";
        }
        catch (Exception ex)
        {
            processStatusLabel.Text = $"Process refresh failed: {ex.Message}";
            LogCall("Processes", "Refresh Failed", ex.Message);
        }
        finally
        {
            isRefreshingProcesses = false;
            refreshProcessesButton.Enabled = true;
        }
    }

    private void RenderProcessTree()
    {
        var selectedPid = GetSelectedProcessId();
        var firstDisplayedRowIndex = GetFirstDisplayedProcessRowIndex();
        var filter = processSearchTextBox.Text.Trim();
        var visibleStates = processStates.Values
            .Where(state => MatchesSearch(state, filter))
            .ToList();

        var childrenByParent = visibleStates
            .OrderBy(state => state.IsExited)
            .ThenBy(state => state.Name, StringComparer.OrdinalIgnoreCase)
            .GroupBy(state => state.ParentProcessId)
            .ToDictionary(group => group.Key ?? 0, group => group.ToList());

        var visiblePidSet = visibleStates.Select(state => state.ProcessId).ToHashSet();

        var rootStates = visibleStates
            .Where(state => state.ParentProcessId is null
                || state.ParentProcessId == state.ProcessId
                || !visiblePidSet.Contains(state.ParentProcessId.Value))
            .OrderBy(state => state.IsExited)
            .ThenBy(state => state.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (rootStates.Count == 0)
        {
            rootStates = visibleStates
                .OrderBy(state => state.IsExited)
                .ThenBy(state => state.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        SuspendControlDrawing(processGrid);
        isRebuildingProcessGrid = true;

        try
        {
            processGrid.SuspendLayout();
            processRows.Clear();

            foreach (var rootState in rootStates)
            {
                AddProcessRows(rootState, childrenByParent, [], 0, processRows, collapsedProcessIds);
            }

            processGrid.RowCount = processRows.Count;
            processGrid.Refresh();
            RestoreSelectedProcessRow(selectedPid);
            RestoreFirstDisplayedProcessRowIndex(firstDisplayedRowIndex);
            UpdateProcessContextMenuState();

            if (!selectedPid.HasValue && processRows.Count > 0)
            {
                processGrid.CurrentCell = processGrid[0, 0];
            }
        }
        finally
        {
            isRebuildingProcessGrid = false;
            processGrid.ResumeLayout();
            ResumeControlDrawing(processGrid);
        }

        UpdateSelectedProcessDetails();
        QueueVisibleCommandLineLoads();
    }

    private static void AddProcessRows(
        ProcessNodeState state,
        IReadOnlyDictionary<int, List<ProcessNodeState>> childrenByParent,
        HashSet<int> ancestry,
        int depth,
        ICollection<ProcessRowEntry> rows,
        ISet<int> collapsedProcessIds)
    {
        if (!ancestry.Add(state.ProcessId))
        {
            return;
        }

        childrenByParent.TryGetValue(state.ProcessId, out var children);
        var hasChildren = children is { Count: > 0 };
        var isExpanded = hasChildren && !collapsedProcessIds.Contains(state.ProcessId);
        rows.Add(new ProcessRowEntry(state, depth, hasChildren, isExpanded));

        if (hasChildren && isExpanded && children is not null)
        {
            foreach (var child in children)
            {
                if (ancestry.Contains(child.ProcessId))
                {
                    continue;
                }

                AddProcessRows(child, childrenByParent, [.. ancestry], depth + 1, rows, collapsedProcessIds);
            }
        }
    }

    private static string FormatProcessName(ProcessNodeState state, int depth)
    {
        var prefix = depth == 0
            ? string.Empty
            : $"{new string(' ', depth * 3)}└ ";
        return prefix + state.Name;
    }

    private static string FormatMemoryColumn(long workingSetBytes)
    {
        return workingSetBytes > 0
            ? $"{workingSetBytes / (1024d * 1024d):N1} MB"
            : "-";
    }

    private int? GetFirstDisplayedProcessRowIndex()
    {
        if (processRows.Count == 0)
        {
            return null;
        }

        try
        {
            return processGrid.FirstDisplayedScrollingRowIndex >= 0
                ? processGrid.FirstDisplayedScrollingRowIndex
                : null;
        }
        catch
        {
            return null;
        }
    }

    private static Color GetProcessNodeColor(ProcessNodeState state)
    {
        if (state.IsExited)
        {
            return Color.FromArgb(220, 110, 110);
        }

        if (state.WasSeenAfterLaunch)
        {
            return Color.FromArgb(94, 234, 150);
        }

        if (state.ParentProcessId is null || state.ParentProcessId == state.ProcessId)
        {
            return Color.FromArgb(120, 171, 219);
        }

        return Color.FromArgb(208, 216, 224);
    }

    private static bool MatchesSearch(ProcessNodeState state, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return true;
        }

        return IsFuzzyMatch(state.Name, filter)
            || (!string.IsNullOrWhiteSpace(state.ExecutablePath) && IsFuzzyMatch(state.ExecutablePath, filter))
            || state.ProcessId.ToString().Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFuzzyMatch(string text, string pattern)
    {
        if (text.Contains(pattern, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var textIndex = 0;
        var patternIndex = 0;
        while (textIndex < text.Length && patternIndex < pattern.Length)
        {
            if (char.ToUpperInvariant(text[textIndex]) == char.ToUpperInvariant(pattern[patternIndex]))
            {
                patternIndex++;
            }

            textIndex++;
        }

        return patternIndex == pattern.Length;
    }

    private void RestoreSelectedProcessRow(int? selectedPid)
    {
        if (!selectedPid.HasValue)
        {
            return;
        }

        for (var rowIndex = 0; rowIndex < processRows.Count; rowIndex++)
        {
            if (TryGetProcessState(rowIndex, out var state) && state.ProcessId == selectedPid.Value)
            {
                processGrid.CurrentCell = processGrid[0, rowIndex];
                break;
            }
        }
    }

    private void RestoreFirstDisplayedProcessRowIndex(int? rowIndex)
    {
        if (!rowIndex.HasValue || processRows.Count == 0)
        {
            return;
        }

        var safeIndex = Math.Max(0, Math.Min(rowIndex.Value, processRows.Count - 1));
        try
        {
            processGrid.FirstDisplayedScrollingRowIndex = safeIndex;
        }
        catch
        {
        }
    }

    private void processGrid_SelectionChanged(object? sender, EventArgs e)
    {
        if (isRebuildingProcessGrid)
        {
            return;
        }

        UpdateSelectedProcessDetails();
        QueueVisibleCommandLineLoads();
    }

    private void processGrid_Scroll(object? sender, ScrollEventArgs e)
    {
        if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
        {
            QueueVisibleCommandLineLoads();
        }
    }

    private void UpdateSelectedProcessDetails()
    {
        var state = GetSelectedProcessState();
        if (state is null)
        {
            detailsNameValueLabel.Text = "-";
            detailsPidValueLabel.Text = "-";
            detailsParentValueLabel.Text = "-";
            detailsStatusValueLabel.Text = "-";
            detailsMemoryValueLabel.Text = "-";
            detailsPathValueLabel.Text = "-";
            securityDetailsTextBox.Text = "Select a process to inspect binary details.";
            lastRenderedDetailsPid = null;
            UpdateProcessContextMenuState();
            return;
        }

        detailsNameValueLabel.Text = state.Name;
        detailsPidValueLabel.Text = state.ProcessId.ToString();
        detailsParentValueLabel.Text = state.ParentProcessId?.ToString() ?? "-";
        detailsStatusValueLabel.Text = state.IsExited ? $"Closed at {state.ExitedAtUtc?.ToLocalTime():T}" : "Running";
        detailsMemoryValueLabel.Text = FormatMemoryColumn(state.WorkingSetBytes);
        UpdateProcessContextMenuState();

        if (lastRenderedDetailsPid == state.ProcessId)
        {
            return;
        }

        lastRenderedDetailsPid = state.ProcessId;
        detailsPathValueLabel.Text = "Loading path...";
        securityDetailsTextBox.Text = "Loading binary intelligence...";
        _ = LoadSelectedProcessDetailsAsync(state);
    }

    private void processGrid_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.RowIndex < 0)
        {
            return;
        }

        if (e.Button == MouseButtons.Right)
        {
            var columnIndex = e.ColumnIndex >= 0 ? e.ColumnIndex : 0;
            processGrid.CurrentCell = processGrid[columnIndex, e.RowIndex];
            UpdateProcessContextMenuState();
            processContextMenu.Show(Cursor.Position);
        }
    }

    private void processGrid_CellMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left || e.RowIndex < 0 || e.ColumnIndex < 0)
        {
            return;
        }

        var nameColumn = processGrid.Columns["NameColumn"];
        if (nameColumn is null || e.ColumnIndex != nameColumn.Index)
        {
            return;
        }

        if (!TryGetProcessRow(e.RowIndex, out var rowEntry) || !rowEntry.HasChildren)
        {
            return;
        }

        if (!IsExpansionGlyphHit(rowEntry, e.X))
        {
            return;
        }

        ToggleProcessExpansion(rowEntry.State.ProcessId);
    }

    private static bool IsExpansionGlyphHit(ProcessRowEntry rowEntry, int clickX)
    {
        var glyphStart = rowEntry.Depth * 22;
        return clickX >= glyphStart && clickX <= glyphStart + 32;
    }

    private void ToggleProcessExpansion(int processId)
    {
        if (!collapsedProcessIds.Add(processId))
        {
            collapsedProcessIds.Remove(processId);
        }

        RenderProcessTree();
    }

    private void processGrid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || !TryGetProcessState(e.RowIndex, out var state))
        {
            return;
        }

        var statusColumn = processGrid.Columns["StatusColumn"];
        if (statusColumn is not null && e.ColumnIndex == statusColumn.Index)
        {
            e.CellStyle.ForeColor = state.IsExited
                ? Color.FromArgb(220, 110, 110)
                : state.WasSeenAfterLaunch
                    ? Color.FromArgb(94, 234, 150)
                    : Color.FromArgb(203, 213, 225);
        }

        var nameColumn = processGrid.Columns["NameColumn"];
        if (nameColumn is not null
            && e.ColumnIndex == nameColumn.Index
            && (state.ParentProcessId is null || state.ParentProcessId == state.ProcessId))
        {
            e.CellStyle.ForeColor = Color.FromArgb(120, 171, 219);
        }
    }

    private void processGrid_CellValueNeeded(object? sender, DataGridViewCellValueEventArgs e)
    {
        if (!TryGetProcessRow(e.RowIndex, out var rowEntry))
        {
            e.Value = string.Empty;
            return;
        }

        var state = rowEntry.State;
        var columnName = processGrid.Columns[e.ColumnIndex].Name;
        e.Value = columnName switch
        {
            "NameColumn" => FormatProcessNamePlain(rowEntry),
            "ProcessNameColumn" => state.ProcessName,
            "PidColumn" => state.ProcessId.ToString(),
            "CpuColumn" => $"{state.CpuPercent:N1}%",
            "MemoryColumn" => FormatMemoryColumn(state.WorkingSetBytes),
            "GpuColumn" => FormatPercentColumn(state.GpuPercent),
            "DiskColumn" => FormatThroughputColumn(state.DiskBytesPerSecond),
            "NetworkColumn" => FormatNetworkColumn(state.NetworkBytesPerSecond, state.NetworkConnectionCount),
            "StatusColumn" => state.IsExited ? "Closed" : "Running",
            "CommandLineColumn" => string.IsNullOrWhiteSpace(state.CommandLine) ? "-" : state.CommandLine,
            _ => string.Empty
        };
    }

    private static string FormatPercentColumn(double value)
    {
        return value > 0
            ? $"{value:N1}%"
            : "0.0%";
    }

    private static string FormatThroughputColumn(double bytesPerSecond)
    {
        if (bytesPerSecond <= 0)
        {
            return "0 KB/s";
        }

        var kibPerSecond = bytesPerSecond / 1024d;
        if (kibPerSecond >= 1024d)
        {
            return $"{kibPerSecond / 1024d:N1} MB/s";
        }

        return $"{kibPerSecond:N0} KB/s";
    }

    private static string FormatNetworkColumn(double bytesPerSecond, int connectionCount)
    {
        if (bytesPerSecond > 0)
        {
            return FormatThroughputColumn(bytesPerSecond);
        }

        return connectionCount > 0
            ? $"{connectionCount} conn"
            : "0 conn";
    }

    private void UpdateProcessContextMenuState()
    {
        var state = GetSelectedProcessState();
        if (state is null)
        {
            endTaskToolStripMenuItem.Enabled = false;
            endTaskMenuToolStripMenuItem.Enabled = false;
            openFileLocationToolStripMenuItem.Enabled = false;
            endTaskButton.Enabled = false;
            return;
        }

        endTaskToolStripMenuItem.Enabled = !state.IsExited;
        endTaskMenuToolStripMenuItem.Enabled = !state.IsExited;
        openFileLocationToolStripMenuItem.Enabled = !state.IsExited || !string.IsNullOrWhiteSpace(state.ExecutablePath);
        endTaskButton.Enabled = !state.IsExited;
    }

    private int? GetSelectedProcessId()
    {
        return GetSelectedProcessState()?.ProcessId;
    }

    private void EndSelectedProcess()
    {
        var state = GetSelectedProcessState();
        if (state is null || state.IsExited)
        {
            return;
        }

        try
        {
            using var process = Process.GetProcessById(state.ProcessId);
            process.Kill(true);
            processStatusLabel.Text = $"Ended {state.Name} (PID {state.ProcessId})";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to end process {state.Name}.\n\n{ex.Message}", "End Task Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogCall("Processes", "End Task Failed", $"{state.Name} ({state.ProcessId}): {ex.Message}");
        }
        finally
        {
            LogCall("Processes", "End Task", $"{state.Name} ({state.ProcessId})");
            _ = RefreshProcessesAsync();
        }
    }

    private async Task OpenSelectedProcessLocationAsync()
    {
        var state = GetSelectedProcessState();
        if (state is null)
        {
            return;
        }

        var path = state.ExecutablePath;
        if (string.IsNullOrWhiteSpace(path) && !state.IsExited)
        {
            path = await Task.Run(() => NativeProcessMethods.TryGetProcessPath(state.ProcessId));
            state.ExecutablePath = path;
        }

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            MessageBox.Show(this, "Executable path is unavailable for the selected process.", "Open File Location", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LogCall("Binary", "Open Location Failed", $"{state.Name} ({state.ProcessId}): path unavailable");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{path}\"",
                UseShellExecute = true
            });
            LogCall("Binary", "Open Location", path);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to open file location.\n\n{ex.Message}", "Open File Location Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogCall("Binary", "Open Location Failed", ex.Message);
        }
    }

    private async Task LoadSelectedProcessDetailsAsync(ProcessNodeState state)
    {
        var path = state.ExecutablePath;

        if (!state.IsExited && string.IsNullOrWhiteSpace(path))
        {
            path = await Task.Run(() => NativeProcessMethods.TryGetProcessPath(state.ProcessId));
            state.ExecutablePath = path;
        }

        BinarySecurityDetails details;
        if (!string.IsNullOrWhiteSpace(path) && binaryDetailsCache.TryGetValue(path, out var cached))
        {
            details = cached;
        }
        else
        {
            details = await Task.Run(() => BinarySecurityInspector.ReadDetails(path, state));
        }

        var commandLine = await EnsureCommandLineLoadedAsync(state);
        state.CommandLine = commandLine;
        processCommandLineCache[state.ProcessId] = commandLine;

        if (GetSelectedProcessState() is ProcessNodeState current && current.ProcessId == state.ProcessId)
        {
            detailsPathValueLabel.Text = string.IsNullOrWhiteSpace(details.Path) ? "Unavailable" : details.Path;
            securityDetailsTextBox.Text = details.ToDisplayText(commandLine);
        }

        var commandLineColumn = processGrid.Columns[processCommandLineColumnName];
        if (commandLineColumn is not null)
        {
            processGrid.InvalidateColumn(commandLineColumn.Index);
        }

        if (!string.IsNullOrWhiteSpace(details.Path))
        {
            binaryDetailsCache[details.Path] = details;
        }
        LogCall("Binary", "Inspect", $"{state.Name} ({state.ProcessId})");
    }

    private void QueueVisibleCommandLineLoads()
    {
        if (!IsHandleCreated || processRows.Count == 0)
        {
            return;
        }

        int firstRowIndex;
        int displayedRowCount;
        try
        {
            firstRowIndex = Math.Max(0, processGrid.FirstDisplayedScrollingRowIndex);
            displayedRowCount = Math.Max(1, processGrid.DisplayedRowCount(false));
        }
        catch
        {
            return;
        }

        var lastRowIndex = Math.Min(processRows.Count - 1, firstRowIndex + Math.Max(40, displayedRowCount + 40));
        for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
        {
            if (!TryGetProcessState(rowIndex, out var state)
                || state.IsExited
                || !string.IsNullOrWhiteSpace(state.CommandLine))
            {
                continue;
            }

            _ = EnsureCommandLineLoadedAsync(state);
        }
    }

    private async Task<string> EnsureCommandLineLoadedAsync(ProcessNodeState state)
    {
        if (!string.IsNullOrWhiteSpace(state.CommandLine))
        {
            return state.CommandLine;
        }

        if (processCommandLineCache.TryGetValue(state.ProcessId, out var cachedCommandLine))
        {
            state.CommandLine = cachedCommandLine;
            return cachedCommandLine;
        }

        Task<string> loadTask;
        lock (commandLineLoadSync)
        {
            if (!commandLineLoadTasks.TryGetValue(state.ProcessId, out loadTask!))
            {
                loadTask = LoadCommandLineCoreAsync(state);
                commandLineLoadTasks[state.ProcessId] = loadTask;
            }
        }

        return await loadTask;
    }

    private async Task<string> LoadCommandLineCoreAsync(ProcessNodeState state)
    {
        try
        {
            await commandLineLoadLimiter.WaitAsync();

            if (processCommandLineCache.TryGetValue(state.ProcessId, out var cachedCommandLine))
            {
                state.CommandLine = cachedCommandLine;
                return cachedCommandLine;
            }

            var commandLine = await Task.Run(() => ProcessMetadataProvider.Current.TryGetCommandLine(state.ProcessId));
            if (string.IsNullOrWhiteSpace(commandLine))
            {
                commandLine = "Unavailable";
            }

            state.CommandLine = commandLine;
            processCommandLineCache[state.ProcessId] = commandLine;
            InvalidateCommandLineColumn();
            return commandLine;
        }
        finally
        {
            commandLineLoadLimiter.Release();
            lock (commandLineLoadSync)
            {
                commandLineLoadTasks.Remove(state.ProcessId);
            }
        }
    }

    private void InvalidateCommandLineColumn()
    {
        if (!IsHandleCreated)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(InvalidateCommandLineColumn);
            return;
        }

        var commandLineColumn = processGrid.Columns[processCommandLineColumnName];
        if (commandLineColumn is not null)
        {
            processGrid.InvalidateColumn(commandLineColumn.Index);
        }
    }

    private async Task RefreshStartupItemsAsync()
    {
        if (isRefreshingStartup)
        {
            return;
        }

        isRefreshingStartup = true;
        refreshStartupButton.Enabled = false;

        try
        {
            var entries = await Task.Run(() => StartupReader.ReadStartupEntries()
                .OrderBy(entry => entry.Scope)
                .ThenBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                .ToList());

            startupGrid.Rows.Clear();
            foreach (var entry in entries)
            {
                startupGrid.Rows.Add(entry.Name, entry.Scope, entry.Type, entry.Location, entry.Command);
            }

            startupStatusLabel.Text = $"Startup items: {startupGrid.Rows.Count}";
        }
        catch (Exception ex)
        {
            startupStatusLabel.Text = $"Startup refresh failed: {ex.Message}";
            LogCall("Startup", "Refresh Failed", ex.Message);
        }
        finally
        {
            isRefreshingStartup = false;
            refreshStartupButton.Enabled = true;
        }
    }

    private async Task RefreshServicesAsync()
    {
        if (isRefreshingServices)
        {
            return;
        }

        isRefreshingServices = true;
        refreshServicesButton.Enabled = false;

        try
        {
            var entries = await Task.Run(() => ServiceReader.ReadServices().ToList());
            servicesGrid.Rows.Clear();

            foreach (var entry in entries)
            {
                var rowIndex = servicesGrid.Rows.Add(entry.DisplayName, entry.ServiceName, entry.Status, entry.StartType);
                servicesGrid.Rows[rowIndex].Tag = entry;
            }

            servicesStatusLabel.Text = $"Services: {servicesGrid.Rows.Count}";
        }
        catch (Exception ex)
        {
            servicesStatusLabel.Text = $"Services refresh failed: {ex.Message}";
            LogCall("Services", "Refresh Failed", ex.Message);
        }
        finally
        {
            isRefreshingServices = false;
            refreshServicesButton.Enabled = true;
        }
    }

    private void ChangeSelectedService(ServiceAction action)
    {
        if (servicesGrid.CurrentRow?.Tag is not ServiceEntry entry)
        {
            return;
        }

        try
        {
            using var controller = new System.ServiceProcess.ServiceController(entry.ServiceName);

            if (action == ServiceAction.Start && controller.Status == System.ServiceProcess.ServiceControllerStatus.Stopped)
            {
                controller.Start();
                controller.WaitForStatus(System.ServiceProcess.ServiceControllerStatus.Running, TimeSpan.FromSeconds(5));
            }
            else if (action == ServiceAction.Stop && controller.Status == System.ServiceProcess.ServiceControllerStatus.Running)
            {
                controller.Stop();
                controller.WaitForStatus(System.ServiceProcess.ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(5));
            }

            LogCall("Services", action.ToString(), entry.DisplayName);
            _ = RefreshServicesAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to {action.ToString().ToLowerInvariant()} service {entry.DisplayName}.\n\n{ex.Message}", "Service Action Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogCall("Services", $"{action} Failed", $"{entry.DisplayName}: {ex.Message}");
        }
    }

    private void servicesGrid_SelectionChanged(object? sender, EventArgs e)
    {
        if (servicesGrid.CurrentRow?.Tag is not ServiceEntry entry)
        {
            startServiceButton.Enabled = false;
            stopServiceButton.Enabled = false;
            return;
        }

        startServiceButton.Enabled = entry.Status.Equals("Stopped", StringComparison.OrdinalIgnoreCase);
        stopServiceButton.Enabled = entry.Status.Equals("Running", StringComparison.OrdinalIgnoreCase);
    }

    private void LogCall(string area, string action, string details)
    {
        callsGrid.Rows.Insert(0, DateTime.Now.ToString("HH:mm:ss"), area, action, details);

        while (callsGrid.Rows.Count > 250)
        {
            callsGrid.Rows.RemoveAt(callsGrid.Rows.Count - 1);
        }

        UpdateCallsStatus();
    }

    private void UpdateCallsStatus()
    {
        callsStatusLabel.Text = $"Calls logged: {callsGrid.Rows.Count}";
    }

    private void UpdateTabButtons()
    {
        StyleTabButton(processesTabButton, mainTabControl.SelectedTab == processesTabPage);
        StyleTabButton(startupTabButton, mainTabControl.SelectedTab == startupTabPage);
        StyleTabButton(servicesTabButton, mainTabControl.SelectedTab == servicesTabPage);
        StyleTabButton(callsTabButton, mainTabControl.SelectedTab == callsTabPage);
    }

    private static void StyleTabButton(Button button, bool isSelected)
    {
        button.BackColor = isSelected ? Color.FromArgb(30, 41, 59) : Color.FromArgb(15, 22, 30);
        button.ForeColor = isSelected ? Color.FromArgb(241, 245, 249) : Color.FromArgb(148, 163, 184);
    }

    private static void SuspendControlDrawing(Control control)
    {
        SendMessage(control.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
    }

    private static void ResumeControlDrawing(Control control)
    {
        SendMessage(control.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
        control.Invalidate(true);
        control.Update();
    }

    private ProcessNodeState? GetSelectedProcessState()
    {
        return processGrid.CurrentCell is not null && TryGetProcessState(processGrid.CurrentCell.RowIndex, out var state)
            ? state
            : null;
    }

    private bool TryGetProcessState(int rowIndex, out ProcessNodeState state)
    {
        if (TryGetProcessRow(rowIndex, out var rowEntry))
        {
            state = rowEntry.State;
            return true;
        }

        state = null!;
        return false;
    }

    private bool TryGetProcessRow(int rowIndex, out ProcessRowEntry rowEntry)
    {
        if (rowIndex >= 0 && rowIndex < processRows.Count)
        {
            rowEntry = processRows[rowIndex];
            return true;
        }

        rowEntry = default;
        return false;
    }

    private static void EnableDoubleBuffering(DataGridView grid)
    {
        typeof(DataGridView)
            .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            ?.SetValue(grid, true);
    }

    private static string FormatProcessNamePlain(ProcessRowEntry rowEntry)
    {
        var indent = rowEntry.Depth == 0
            ? string.Empty
            : new string(' ', rowEntry.Depth * 3);
        var branch = rowEntry.HasChildren
            ? (rowEntry.IsExpanded ? "[-] " : "[+] ")
            : "    ";
        return indent + branch + rowEntry.State.Name;
    }

    private const int WM_SETREDRAW = 0x000B;

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);
}

internal sealed class DarkMenuRenderer : ToolStripProfessionalRenderer
{
    public DarkMenuRenderer() : base(new DarkColorTable())
    {
    }
}

internal sealed class DarkColorTable : ProfessionalColorTable
{
    public override Color MenuStripGradientBegin => Color.FromArgb(15, 22, 30);
    public override Color MenuStripGradientEnd => Color.FromArgb(15, 22, 30);
    public override Color ToolStripDropDownBackground => Color.FromArgb(20, 27, 35);
    public override Color ImageMarginGradientBegin => Color.FromArgb(20, 27, 35);
    public override Color ImageMarginGradientMiddle => Color.FromArgb(20, 27, 35);
    public override Color ImageMarginGradientEnd => Color.FromArgb(20, 27, 35);
    public override Color MenuItemSelected => Color.FromArgb(30, 41, 59);
    public override Color MenuItemSelectedGradientBegin => Color.FromArgb(30, 41, 59);
    public override Color MenuItemSelectedGradientEnd => Color.FromArgb(30, 41, 59);
    public override Color MenuItemBorder => Color.FromArgb(51, 65, 85);
    public override Color MenuItemPressedGradientBegin => Color.FromArgb(30, 41, 59);
    public override Color MenuItemPressedGradientMiddle => Color.FromArgb(30, 41, 59);
    public override Color MenuItemPressedGradientEnd => Color.FromArgb(30, 41, 59);
    public override Color SeparatorDark => Color.FromArgb(51, 65, 85);
    public override Color SeparatorLight => Color.FromArgb(51, 65, 85);
}

internal readonly record struct ProcessRowEntry(ProcessNodeState State, int Depth, bool HasChildren, bool IsExpanded);

internal sealed class ProcessNodeState
{
    public ProcessNodeState(ProcessSnapshot snapshot, DateTime seenUtc, bool wasSeenAfterLaunch)
    {
        ProcessId = snapshot.ProcessId;
        Name = snapshot.Name;
        ProcessName = snapshot.ProcessName;
        ParentProcessId = snapshot.ParentProcessId;
        ExecutablePath = snapshot.ExecutablePath;
        WorkingSetBytes = snapshot.WorkingSetBytes;
        CommandLine = snapshot.CommandLine;
        LastTotalProcessorTime = snapshot.TotalProcessorTime;
        LastCpuSampleUtc = seenUtc;
        FirstSeenUtc = seenUtc;
        WasSeenAfterLaunch = wasSeenAfterLaunch;
    }

    public int ProcessId { get; }

    public string Name { get; private set; }

    public string ProcessName { get; private set; }

    public int? ParentProcessId { get; private set; }

    public string? ExecutablePath { get; set; }

    public string CommandLine { get; set; }

    public long WorkingSetBytes { get; private set; }

    public double GpuPercent { get; private set; }

    public double DiskBytesPerSecond { get; private set; }

    public double NetworkBytesPerSecond { get; private set; }

    public int NetworkConnectionCount { get; private set; }

    public DateTime FirstSeenUtc { get; }

    public bool WasSeenAfterLaunch { get; }

    public double CpuPercent { get; private set; }

    public bool IsExited { get; private set; }

    public DateTime? ExitedAtUtc { get; private set; }

    private TimeSpan LastTotalProcessorTime { get; set; }

    private DateTime LastCpuSampleUtc { get; set; }

    public void Update(ProcessSnapshot snapshot, DateTime seenUtc)
    {
        Name = snapshot.Name;
        ProcessName = snapshot.ProcessName;
        ParentProcessId = snapshot.ParentProcessId;
        ExecutablePath = snapshot.ExecutablePath;
        if (!string.IsNullOrWhiteSpace(snapshot.CommandLine))
        {
            CommandLine = snapshot.CommandLine;
        }
        WorkingSetBytes = snapshot.WorkingSetBytes;
        GpuPercent = snapshot.GpuPercent;
        DiskBytesPerSecond = snapshot.DiskBytesPerSecond;
        NetworkBytesPerSecond = snapshot.NetworkBytesPerSecond;
        NetworkConnectionCount = snapshot.NetworkConnectionCount;
        CpuPercent = ComputeCpuPercent(snapshot.TotalProcessorTime, seenUtc);
        IsExited = false;
        ExitedAtUtc = null;
    }

    public void MarkExited(DateTime exitedAtUtc)
    {
        IsExited = true;
        ExitedAtUtc = exitedAtUtc;
        WorkingSetBytes = 0;
        CpuPercent = 0;
        GpuPercent = 0;
        DiskBytesPerSecond = 0;
        NetworkBytesPerSecond = 0;
        NetworkConnectionCount = 0;
    }

    private double ComputeCpuPercent(TimeSpan totalProcessorTime, DateTime seenUtc)
    {
        var elapsed = seenUtc - LastCpuSampleUtc;
        var cpuUsed = totalProcessorTime - LastTotalProcessorTime;

        LastTotalProcessorTime = totalProcessorTime;
        LastCpuSampleUtc = seenUtc;

        if (elapsed <= TimeSpan.Zero || cpuUsed < TimeSpan.Zero)
        {
            return CpuPercent;
        }

        var cpu = cpuUsed.TotalMilliseconds / (elapsed.TotalMilliseconds * Environment.ProcessorCount) * 100d;
        return Math.Max(0, Math.Min(cpu, 999.9d));
    }
}

internal sealed record BinarySecurityDetails(
    string? Path,
    string FileName,
    string FileDescription,
    string CompanyName,
    string ProductName,
    string FileVersion,
    string ParentPid,
    string Sha256,
    string SignatureStatus,
    string SignerSubject,
    string SignerIssuer,
    string Thumbprint,
    string LastWriteUtc,
    string FileSize)
{
    public string ToDisplayText(string commandLine)
    {
        return
            $"Binary:         {Path ?? "Unavailable"}{Environment.NewLine}" +
            $"Command Line:   {commandLine}{Environment.NewLine}" +
            $"File Name:      {FileName}{Environment.NewLine}" +
            $"Description:    {FileDescription}{Environment.NewLine}" +
            $"Company:        {CompanyName}{Environment.NewLine}" +
            $"Product:        {ProductName}{Environment.NewLine}" +
            $"Version:        {FileVersion}{Environment.NewLine}" +
            $"Parent PID:     {ParentPid}{Environment.NewLine}" +
            $"SHA256:         {Sha256}{Environment.NewLine}" +
            $"Signature:      {SignatureStatus}{Environment.NewLine}" +
            $"Signer:         {SignerSubject}{Environment.NewLine}" +
            $"Issuer:         {SignerIssuer}{Environment.NewLine}" +
            $"Thumbprint:     {Thumbprint}{Environment.NewLine}" +
            $"File Size:      {FileSize}{Environment.NewLine}" +
            $"Last Write UTC: {LastWriteUtc}";
    }
}

internal static class BinarySecurityInspector
{
    public static BinarySecurityDetails ReadDetails(string? path, ProcessNodeState state)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return new BinarySecurityDetails(
                path,
                "Unavailable",
                "Unavailable",
                "Unavailable",
                "Unavailable",
                "Unavailable",
                state.ParentProcessId?.ToString() ?? "-",
                "Unavailable",
                "Path unavailable",
                "Unsigned",
                "Unsigned",
                "Unsigned",
                "Unavailable",
                "Unavailable");
        }

        var fileInfo = FileVersionInfo.GetVersionInfo(path);
        var file = new FileInfo(path);
        var sha256 = ComputeSha256(path);
        var signature = ReadSignature(path);

        return new BinarySecurityDetails(
            path,
            file.Name,
            fileInfo.FileDescription ?? "Unavailable",
            fileInfo.CompanyName ?? "Unavailable",
            fileInfo.ProductName ?? "Unavailable",
            fileInfo.FileVersion ?? "Unavailable",
            state.ParentProcessId?.ToString() ?? "-",
            sha256,
            signature.Status,
            signature.Subject,
            signature.Issuer,
            signature.Thumbprint,
            file.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss"),
            $"{file.Length / 1024d:N1} KB");
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash);
    }

    private static (string Status, string Subject, string Issuer, string Thumbprint) ReadSignature(string path)
    {
        try
        {
#pragma warning disable SYSLIB0057
            using var certificate = new X509Certificate2(X509Certificate.CreateFromSignedFile(path));
#pragma warning restore SYSLIB0057
            using var chain = new X509Chain();
            chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
            var trusted = chain.Build(certificate);
            return (
                trusted ? "Signed and chain-built" : "Signed, chain not fully trusted",
                certificate.Subject,
                certificate.Issuer,
                certificate.Thumbprint ?? "Unavailable");
        }
        catch
        {
            return ("Unsigned or unreadable signature", "Unsigned", "Unsigned", "Unsigned");
        }
    }
}

internal sealed record ProcessSnapshot(
    int ProcessId,
    int? ParentProcessId,
    string Name,
    string ProcessName,
    string? ExecutablePath,
    string CommandLine,
    long WorkingSetBytes,
    double GpuPercent,
    double DiskBytesPerSecond,
    double NetworkBytesPerSecond,
    int NetworkConnectionCount,
    TimeSpan TotalProcessorTime);

internal static class ProcessSnapshotReader
{
    public static Dictionary<int, ProcessSnapshot> ReadProcesses(ProcessRuntimeMetricsProvider metricsProvider)
    {
        var parents = NativeProcessMethods.ReadParentProcessMap();
        var runtimeMetrics = metricsProvider.ReadCurrentMetrics();
        var snapshots = new Dictionary<int, ProcessSnapshot>();

        foreach (var process in Process.GetProcesses())
        {
            using (process)
            {
                var name = SafeRead(() => process.ProcessName + ".exe", $"pid-{process.Id}");
                var processName = SafeRead(() => process.ProcessName, $"pid-{process.Id}");
                var workingSet = SafeRead(() => process.WorkingSet64, 0L);
                var totalProcessorTime = SafeRead(() => process.TotalProcessorTime, TimeSpan.Zero);
                var metrics = runtimeMetrics.GetValueOrDefault(process.Id);
                int? parentPid = parents.TryGetValue(process.Id, out var value) && value > 0 ? value : null;
                snapshots[process.Id] = new ProcessSnapshot(
                    process.Id,
                    parentPid,
                    name,
                    processName,
                    null,
                    string.Empty,
                    workingSet,
                    metrics.GpuPercent,
                    metrics.DiskBytesPerSecond,
                    metrics.NetworkBytesPerSecond,
                    metrics.ConnectionCount,
                    totalProcessorTime);
            }
        }

        return snapshots;
    }

    private static T SafeRead<T>(Func<T> reader, T fallback)
    {
        try
        {
            return reader();
        }
        catch
        {
            return fallback;
        }
    }
}

internal readonly record struct ProcessRuntimeMetrics(double DiskBytesPerSecond, double NetworkBytesPerSecond, int ConnectionCount, double GpuPercent);

internal sealed class ProcessRuntimeMetricsProvider : IDisposable
{
    private readonly object sync = new();
    private readonly Dictionary<int, (ulong Bytes, DateTime SampleUtc)> lastDiskSamples = [];
    private readonly Dictionary<string, CounterSample> lastGpuSamples = new(StringComparer.OrdinalIgnoreCase);
    private readonly bool gpuCountersAvailable;

    public ProcessRuntimeMetricsProvider()
    {
        try
        {
            gpuCountersAvailable = PerformanceCounterCategory.Exists("GPU Engine");
        }
        catch
        {
            gpuCountersAvailable = false;
        }
    }

    public IReadOnlyDictionary<int, ProcessRuntimeMetrics> ReadCurrentMetrics()
    {
        lock (sync)
        {
            var connectionCounts = NativeProcessMethods.ReadNetworkConnectionCounts();
            var gpuPercents = gpuCountersAvailable ? ReadGpuUsage() : new Dictionary<int, double>();
            var metrics = new Dictionary<int, ProcessRuntimeMetrics>();
            var sampledAtUtc = DateTime.UtcNow;

            foreach (var process in Process.GetProcesses())
            {
                using (process)
                {
                    var pid = process.Id;
                    var diskRate = NativeProcessMethods.TryGetProcessDiskBytes(pid, out var totalDiskBytes)
                        ? ComputeDiskBytesPerSecond(pid, totalDiskBytes, sampledAtUtc)
                        : 0d;

                    connectionCounts.TryGetValue(pid, out var connectionCount);
                    gpuPercents.TryGetValue(pid, out var gpuPercent);

                    metrics[pid] = new ProcessRuntimeMetrics(
                        diskRate,
                        0d,
                        connectionCount,
                        gpuPercent);
                }
            }

            CleanupState(metrics.Keys);
            return metrics;
        }
    }

    public void Dispose()
    {
    }

    private double ComputeDiskBytesPerSecond(int pid, ulong totalDiskBytes, DateTime sampledAtUtc)
    {
        if (!lastDiskSamples.TryGetValue(pid, out var previous))
        {
            lastDiskSamples[pid] = (totalDiskBytes, sampledAtUtc);
            return 0d;
        }

        lastDiskSamples[pid] = (totalDiskBytes, sampledAtUtc);
        var elapsedSeconds = (sampledAtUtc - previous.SampleUtc).TotalSeconds;
        if (elapsedSeconds <= 0 || totalDiskBytes < previous.Bytes)
        {
            return 0d;
        }

        return (totalDiskBytes - previous.Bytes) / elapsedSeconds;
    }

    private Dictionary<int, double> ReadGpuUsage()
    {
        var usageByPid = new Dictionary<int, double>();
        string[] instanceNames;

        try
        {
            instanceNames = new PerformanceCounterCategory("GPU Engine").GetInstanceNames();
        }
        catch
        {
            return usageByPid;
        }

        var activeInstances = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var instanceName in instanceNames)
        {
            var pid = TryParseGpuPid(instanceName);
            if (!pid.HasValue)
            {
                continue;
            }

            activeInstances.Add(instanceName);

            try
            {
                using var counter = new PerformanceCounter("GPU Engine", "Utilization Percentage", instanceName, readOnly: true);
                var currentSample = counter.NextSample();
                double utilization = 0;

                if (lastGpuSamples.TryGetValue(instanceName, out var previousSample))
                {
                    utilization = CounterSample.Calculate(previousSample, currentSample);
                }

                lastGpuSamples[instanceName] = currentSample;
                if (!double.IsNaN(utilization) && !double.IsInfinity(utilization) && utilization > 0)
                {
                    usageByPid[pid.Value] = usageByPid.GetValueOrDefault(pid.Value) + utilization;
                }
            }
            catch
            {
            }
        }

        foreach (var staleInstance in lastGpuSamples.Keys.Except(activeInstances, StringComparer.OrdinalIgnoreCase).ToArray())
        {
            lastGpuSamples.Remove(staleInstance);
        }

        return usageByPid;
    }

    private static int? TryParseGpuPid(string instanceName)
    {
        const string marker = "pid_";
        var start = instanceName.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (start < 0)
        {
            return null;
        }

        start += marker.Length;
        var end = start;
        while (end < instanceName.Length && char.IsDigit(instanceName[end]))
        {
            end++;
        }

        return end > start && int.TryParse(instanceName[start..end], out var pid)
            ? pid
            : null;
    }

    private void CleanupState(IEnumerable<int> livePids)
    {
        var livePidSet = livePids.ToHashSet();
        foreach (var stalePid in lastDiskSamples.Keys.Where(pid => !livePidSet.Contains(pid)).ToArray())
        {
            lastDiskSamples.Remove(stalePid);
        }
    }
}

internal sealed record StartupEntry(string Name, string Scope, string Type, string Location, string Command);

internal static class StartupReader
{
    public static IEnumerable<StartupEntry> ReadStartupEntries()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in ReadStandardRegistryLocations())
        {
            if (seen.Add($"{entry.Type}|{entry.Location}|{entry.Name}|{entry.Command}"))
            {
                yield return entry;
            }
        }

        foreach (var entry in ReadStartupFolder(Environment.SpecialFolder.Startup, "Current user"))
        {
            yield return entry;
        }

        foreach (var entry in ReadStartupFolder(Environment.SpecialFolder.CommonStartup, "All users"))
        {
            yield return entry;
        }
    }

    private static IEnumerable<StartupEntry> ReadStandardRegistryLocations()
    {
        var locations = new (RegistryHive Hive, RegistryView View, string SubKeyPath, string Scope, string Type)[]
        {
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\Run", "Current user", "Registry Run"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "Current user", "Registry RunOnce"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "Current user", "Registry Policy Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\Run", "All users x64", "Registry Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "All users x64", "Registry RunOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "All users x64", "Registry Policy Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\Run", "All users x86", "Registry Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "All users x86", "Registry RunOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "All users x86", "Registry Policy Run")
        };

        foreach (var location in locations)
        {
            foreach (var entry in ReadRegistryEntries(location.Hive, location.View, location.SubKeyPath, location.Scope, location.Type))
            {
                yield return entry;
            }
        }
    }

    private static IEnumerable<StartupEntry> ReadRegistryEntries(RegistryHive hive, RegistryView view, string subKeyPath, string scope, string type)
    {
        using var key = RegistryKey.OpenBaseKey(hive, view).OpenSubKey(subKeyPath);
        if (key is null)
        {
            yield break;
        }

        foreach (var valueName in key.GetValueNames())
        {
            yield return new StartupEntry(
                valueName,
                scope,
                type,
                $@"{hive} [{view}]\{subKeyPath}",
                key.GetValue(valueName)?.ToString() ?? string.Empty);
        }
    }

    private static IEnumerable<StartupEntry> ReadStartupFolder(Environment.SpecialFolder folder, string scope)
    {
        var path = Environment.GetFolderPath(folder);
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
            yield break;
        }

        foreach (var file in Directory.EnumerateFiles(path).OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
        {
            yield return new StartupEntry(Path.GetFileNameWithoutExtension(file), scope, "Startup folder", path, file);
        }
    }
}

internal sealed record ServiceEntry(string ServiceName, string DisplayName, string Status, string StartType);

internal static class ServiceReader
{
    public static IEnumerable<ServiceEntry> ReadServices()
    {
        foreach (var service in System.ServiceProcess.ServiceController.GetServices().OrderBy(service => service.DisplayName, StringComparer.OrdinalIgnoreCase))
        {
            using (service)
            {
                yield return new ServiceEntry(
                    service.ServiceName,
                    service.DisplayName,
                    service.Status.ToString(),
                    NativeProcessMethods.TryReadServiceStartType(service.ServiceName));
            }
        }
    }
}

internal enum ServiceAction
{
    Start,
    Stop
}
