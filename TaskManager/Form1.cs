using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Diagnostics.Tracing.Parsers;
using Microsoft.Diagnostics.Tracing.Session;
using Microsoft.Win32;

namespace TaskManager;

public partial class Form1 : Form
{
    private readonly Dictionary<int, ProcessNodeState> processStates = [];
    private readonly List<ProcessHistoryEntry> processHistoryEntries = [];
    private readonly Dictionary<int, ProcessHistoryEntry> activeProcessHistoryEntries = [];
    private readonly Dictionary<string, BinarySecurityDetails> binaryDetailsCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, string> processCommandLineCache = [];
    private readonly Dictionary<ConnectionRecordKey, ConnectionRecord> connectionRecords = [];
    private readonly HashSet<string> pendingDnsLookups = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ProcessRowEntry> processRows = [];
    private readonly List<ProcessHistoryRowEntry> processHistoryRows = [];
    private readonly HashSet<int> collapsedProcessIds = [];
    private readonly HashSet<long> collapsedProcessHistoryIds = [];
    private readonly ProcessRuntimeMetricsProvider processRuntimeMetricsProvider = new();
    private readonly SemaphoreSlim commandLineLoadLimiter = new(2, 2);
    private readonly Dictionary<int, Task<string>> commandLineLoadTasks = [];
    private readonly object commandLineLoadSync = new();
    private readonly System.Windows.Forms.Timer refreshTimer = new() { Interval = 1000 };
    private readonly DateTime appStartedUtc = DateTime.UtcNow;
    private readonly string processNameColumnName = "ProcessNameColumn";
    private readonly string processCommandLineColumnName = "CommandLineColumn";
    private long nextProcessHistoryId = 1;
    private string? processSortColumnName;
    private SortOrder processSortOrder = SortOrder.None;
    private bool isRefreshingProcesses;
    private bool isRefreshingStartup;
    private bool isRefreshingTasks;
    private bool isRefreshingServices;
    private bool isRefreshingConnections;
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
        processHistoryGrid.CellFormatting += processHistoryGrid_CellFormatting;
        processGrid.CellValueNeeded += processGrid_CellValueNeeded;
        processGrid.ColumnHeaderMouseClick += processGrid_ColumnHeaderMouseClick;
        processGrid.Scroll += processGrid_Scroll;
        processHistoryGrid.CellMouseClick += processHistoryGrid_CellMouseClick;
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
        EnableDoubleBuffering(processHistoryGrid);
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
    private async void refreshTasksButton_Click(object? sender, EventArgs e) => await RefreshScheduledTasksAsync();

    private async void refreshServicesButton_Click(object? sender, EventArgs e) => await RefreshServicesAsync();

    private async void openFileLocationToolStripMenuItem_Click(object? sender, EventArgs e) => await OpenSelectedProcessLocationAsync();

    private void processSearchTextBox_TextChanged(object? sender, EventArgs e) => RenderProcessTree();
    private void processHistorySearchTextBox_TextChanged(object? sender, EventArgs e) => RenderProcessHistoryGrid();

    private void processesTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = processesTabPage;
    private void processHistoryTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = processHistoryTabPage;

    private void startupTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = startupTabPage;
    private void tasksTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = tasksTabPage;

    private void servicesTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = servicesTabPage;

    private void conTabButton_Click(object? sender, EventArgs e) => mainTabControl.SelectedTab = conTabPage;

    private void mainTabControl_SelectedIndexChanged(object? sender, EventArgs e) => UpdateTabButtons();

    private void clearCallsButton_Click(object? sender, EventArgs e)
    {
        connectionRecords.Clear();
        callsGrid.Rows.Clear();
        UpdateConnectionStatus();
    }

    private void startServiceButton_Click(object? sender, EventArgs e) => ChangeSelectedService(ServiceAction.Start);

    private void stopServiceButton_Click(object? sender, EventArgs e) => ChangeSelectedService(ServiceAction.Stop);

    private async Task BeginInitialLoadAsync()
    {
        processStatusLabel.Text = "Loading processes...";
        processHistoryStatusLabel.Text = "Watching for process launches and exits...";
        startupStatusLabel.Text = "Loading startup items...";
        tasksStatusLabel.Text = "Loading scheduled tasks...";
        servicesStatusLabel.Text = "Loading services...";
        callsStatusLabel.Text = "Loading connections...";
        var processTask = RefreshProcessesAsync();
        var startupTask = RefreshStartupItemsAsync();
        var tasksTask = RefreshScheduledTasksAsync();
        var servicesTask = RefreshServicesAsync();
        await Task.WhenAll(processTask, startupTask, tasksTask, servicesTask);
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
                    TrackProcessHistory(createdState, nowUtc);
                    continue;
                }

                state.Update(snapshot, nowUtc);
                if (processCommandLineCache.TryGetValue(snapshot.ProcessId, out var cachedCommandLineForState))
                {
                    state.CommandLine = cachedCommandLineForState;
                }

                UpdateActiveProcessHistory(state, nowUtc);
            }

            foreach (var state in processStates.Values.Where(state => !state.IsExited && !livePids.Contains(state.ProcessId)))
            {
                state.MarkExited(nowUtc);
                MarkProcessHistoryExited(state.ProcessId, nowUtc);
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

            await RefreshConnectionsAsync(currentSnapshot);
            RenderProcessTree();
            RenderProcessHistoryGrid();
            processStatusLabel.Text = $"Processes: {currentSnapshot.Count} live, {processStates.Values.Count(state => state.IsExited)} recently closed";
        }
        catch (Exception ex)
        {
            processStatusLabel.Text = $"Process refresh failed: {ex.Message}";
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
            .ToList();

        if (rootStates.Count == 0)
        {
            rootStates = visibleStates
                .ToList();
        }

        SortProcessStates(rootStates);
        foreach (var childList in childrenByParent.Values)
        {
            SortProcessStates(childList);
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
        if (state.WasSeenAfterLaunch)
        {
            return Color.FromArgb(94, 234, 150);
        }

        if (state.IsExited)
        {
            return Color.FromArgb(220, 110, 110);
        }

        if (state.ParentProcessId is null || state.ParentProcessId == state.ProcessId)
        {
            return Color.FromArgb(96, 165, 250);
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
        if (nameColumn is not null && e.ColumnIndex == nameColumn.Index)
        {
            e.CellStyle.ForeColor = GetProcessNodeColor(state);
        }
    }

    private void processGrid_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.ColumnIndex < 0 || e.ColumnIndex >= processGrid.Columns.Count)
        {
            return;
        }

        var clickedColumn = processGrid.Columns[e.ColumnIndex];
        processSortOrder = processSortColumnName == clickedColumn.Name && processSortOrder == SortOrder.Ascending
            ? SortOrder.Descending
            : SortOrder.Ascending;
        processSortColumnName = clickedColumn.Name;
        UpdateProcessSortGlyphs();
        RenderProcessTree();
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
            return "0 B/s";
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
            UpdateActiveProcessHistory(state, DateTime.UtcNow);
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
                .ThenBy(entry => entry.Type)
                .ThenBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                .ToList());

            startupGrid.Rows.Clear();
            foreach (var entry in entries)
            {
                startupGrid.Rows.Add(entry.Name, entry.Scope, entry.Type, entry.LastRunTime, entry.NextExecutionTime, entry.Location, entry.Command);
            }

            startupStatusLabel.Text = $"Startup items: {startupGrid.Rows.Count}";
        }
        catch (Exception ex)
        {
            startupStatusLabel.Text = $"Startup refresh failed: {ex.Message}";
        }
        finally
        {
            isRefreshingStartup = false;
            refreshStartupButton.Enabled = true;
        }
    }

    private async Task RefreshScheduledTasksAsync()
    {
        if (isRefreshingTasks)
        {
            return;
        }

        isRefreshingTasks = true;
        refreshTasksButton.Enabled = false;

        try
        {
            var entries = await Task.Run(() => StartupReader.ReadScheduledTaskStartupEntries()
                .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                .ToList());

            tasksGrid.Rows.Clear();
            foreach (var entry in entries)
            {
                tasksGrid.Rows.Add(entry.Name, entry.LastRunTime, entry.NextExecutionTime, entry.Command);
            }

            tasksStatusLabel.Text = $"Scheduled startup tasks: {tasksGrid.Rows.Count}";
        }
        catch (Exception ex)
        {
            tasksStatusLabel.Text = $"Scheduled task refresh failed: {ex.Message}";
            LogCall("Tasks", "Refresh Failed", ex.Message);
        }
        finally
        {
            isRefreshingTasks = false;
            refreshTasksButton.Enabled = true;
        }
    }

    private async Task RefreshConnectionsAsync(IReadOnlyDictionary<int, ProcessSnapshot> currentSnapshot)
    {
        if (isRefreshingConnections)
        {
            return;
        }

        isRefreshingConnections = true;

        try
        {
            var snapshots = await Task.Run(() => NativeProcessMethods.ReadActiveTcpConnections());
            var seenAt = DateTime.Now;

            foreach (var snapshot in snapshots)
            {
                var key = new ConnectionRecordKey(snapshot.ProcessId, snapshot.Protocol, snapshot.RemoteAddress, snapshot.RemotePort);
                var binaryPath = currentSnapshot.TryGetValue(snapshot.ProcessId, out var processSnapshot) && !string.IsNullOrWhiteSpace(processSnapshot.ExecutablePath)
                    ? processSnapshot.ExecutablePath!
                    : NativeProcessMethods.TryGetProcessPath(snapshot.ProcessId) ?? $"pid-{snapshot.ProcessId}";
                var binaryName = Path.GetFileName(binaryPath);
                var uploadValue = currentSnapshot.TryGetValue(snapshot.ProcessId, out var networkSnapshot)
                    ? FormatThroughputColumn(networkSnapshot.NetworkUploadBytesPerSecond)
                    : "0 B/s";
                var downloadValue = currentSnapshot.TryGetValue(snapshot.ProcessId, out networkSnapshot)
                    ? FormatThroughputColumn(networkSnapshot.NetworkDownloadBytesPerSecond)
                    : "0 B/s";
                var dnsName = snapshot.RemoteAddress;

                if (!connectionRecords.TryGetValue(key, out var record))
                {
                    record = new ConnectionRecord(
                        key,
                        snapshot.ProcessId,
                        binaryName,
                        binaryPath,
                        dnsName,
                        snapshot.RemoteAddress,
                        uploadValue,
                        downloadValue,
                        seenAt,
                        seenAt);
                    connectionRecords[key] = record;
                    QueueDnsLookup(snapshot.RemoteAddress, key);
                }
                else
                {
                    connectionRecords[key] = record with
                    {
                        ProcessId = snapshot.ProcessId,
                        BinaryName = binaryName,
                        BinaryPath = binaryPath,
                        Upload = uploadValue,
                        Download = downloadValue,
                        LastSeenLocal = seenAt
                    };
                }
            }

            UpdateConnectionGrid();
        }
        catch (Exception ex)
        {
            callsStatusLabel.Text = $"Network refresh failed: {ex.Message}";
        }
        finally
        {
            isRefreshingConnections = false;
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
            var entries = await Task.Run(() => ServiceReader.ReadServices()
                .OrderBy(entry => entry.StartType, StringComparer.OrdinalIgnoreCase)
                .ThenBy(entry => entry.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList());
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

            _ = RefreshServicesAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Unable to {action.ToString().ToLowerInvariant()} service {entry.DisplayName}.\n\n{ex.Message}", "Service Action Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

    private void UpdateConnectionGrid()
    {
        callsGrid.Rows.Clear();

        foreach (var record in connectionRecords.Values.OrderByDescending(record => record.LastSeenLocal).ThenBy(record => record.BinaryName, StringComparer.OrdinalIgnoreCase))
        {
            callsGrid.Rows.Add(
                record.BinaryName,
                record.DnsName,
                record.RemoteAddress,
                record.Upload,
                record.Download);
        }

        UpdateConnectionStatus();
    }

    private void UpdateConnectionStatus()
    {
        callsStatusLabel.Text = $"Network records: {connectionRecords.Count} | Upload/Download counters currently unavailable";
    }

    private void TrackProcessHistory(ProcessNodeState state, DateTime seenUtc)
    {
        if (activeProcessHistoryEntries.ContainsKey(state.ProcessId))
        {
            UpdateActiveProcessHistory(state, seenUtc);
            return;
        }

        var entry = new ProcessHistoryEntry(
            nextProcessHistoryId++,
            state.ProcessId,
            state.ParentProcessId,
            state.ParentProcessId.HasValue && activeProcessHistoryEntries.TryGetValue(state.ParentProcessId.Value, out var parentHistoryEntry)
                ? parentHistoryEntry.Id
                : null,
            state.Name,
            state.ProcessName,
            state.ExecutablePath ?? "Unavailable",
            string.IsNullOrWhiteSpace(state.CommandLine) ? "Unavailable" : state.CommandLine,
            seenUtc.ToLocalTime(),
            null,
            state.WasSeenAfterLaunch,
            "Running");
        processHistoryEntries.Add(entry);
        activeProcessHistoryEntries[state.ProcessId] = entry;
        ExpandProcessHistoryAncestors(entry.ParentHistoryId);
    }

    private void UpdateActiveProcessHistory(ProcessNodeState state, DateTime seenUtc)
    {
        if (!activeProcessHistoryEntries.TryGetValue(state.ProcessId, out var entry))
        {
            TrackProcessHistory(state, seenUtc);
            return;
        }

        entry.Name = state.Name;
        entry.ProcessName = state.ProcessName;
        entry.ParentProcessId = state.ParentProcessId;
        if (state.ParentProcessId.HasValue && activeProcessHistoryEntries.TryGetValue(state.ParentProcessId.Value, out var parentHistoryEntry))
        {
            entry.ParentHistoryId = parentHistoryEntry.Id;
        }
        entry.BinaryPath = string.IsNullOrWhiteSpace(state.ExecutablePath) ? entry.BinaryPath : state.ExecutablePath!;
        if (!string.IsNullOrWhiteSpace(state.CommandLine))
        {
            entry.CommandLine = state.CommandLine;
        }
    }

    private void MarkProcessHistoryExited(int processId, DateTime exitedUtc)
    {
        if (!activeProcessHistoryEntries.TryGetValue(processId, out var entry))
        {
            return;
        }

        entry.ClosedAtLocal = exitedUtc.ToLocalTime();
        entry.Status = "Closed";
        activeProcessHistoryEntries.Remove(processId);
    }

    private void RenderProcessHistoryGrid()
    {
        var selectedHistoryId = GetSelectedProcessHistoryId();
        var firstDisplayedRowIndex = GetFirstDisplayedProcessHistoryRowIndex();
        var filter = processHistorySearchTextBox.Text.Trim();
        var filteredEntries = processHistoryEntries
            .Where(entry => MatchesProcessHistory(entry, filter))
            .ToList();

        var visibleHistoryIds = filteredEntries.Select(entry => entry.Id).ToHashSet();
        var childrenByParent = filteredEntries
            .GroupBy(entry => entry.ParentHistoryId)
            .ToDictionary(group => group.Key ?? 0L, group => group.ToList());
        var latestActivityByHistoryId = new Dictionary<long, DateTime>();

        var rootEntries = filteredEntries
            .Where(entry => entry.ParentHistoryId is null
                || !visibleHistoryIds.Contains(entry.ParentHistoryId.Value))
            .OrderByDescending(entry => GetLatestProcessHistoryActivity(entry, childrenByParent, latestActivityByHistoryId, []))
            .ThenBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
            .ThenByDescending(entry => entry.Id)
            .ToList();

        foreach (var childList in childrenByParent.Values)
        {
            childList.Sort((left, right) =>
            {
                var latestComparison = GetLatestProcessHistoryActivity(right, childrenByParent, latestActivityByHistoryId, [])
                    .CompareTo(GetLatestProcessHistoryActivity(left, childrenByParent, latestActivityByHistoryId, []));
                if (latestComparison != 0)
                {
                    return latestComparison;
                }

                var nameComparison = string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase);
                return nameComparison != 0
                    ? nameComparison
                    : right.Id.CompareTo(left.Id);
            });
        }

        processHistoryGrid.SuspendLayout();
        processHistoryRows.Clear();
        foreach (var entry in rootEntries)
        {
            AddProcessHistoryRows(entry, 0, childrenByParent, [], processHistoryRows, collapsedProcessHistoryIds);
        }

        processHistoryGrid.Rows.Clear();
        foreach (var rowEntry in processHistoryRows)
        {
            var rowIndex = processHistoryGrid.Rows.Add(
                FormatHistoryName(rowEntry),
                rowEntry.Entry.ProcessName,
                rowEntry.Entry.ProcessId,
                rowEntry.Entry.StartedAtLocal.ToString("yyyy-MM-dd HH:mm:ss"),
                rowEntry.Entry.ClosedAtLocal?.ToString("yyyy-MM-dd HH:mm:ss") ?? "-",
                rowEntry.Entry.Status,
                rowEntry.Entry.BinaryPath,
                rowEntry.Entry.CommandLine);
            processHistoryGrid.Rows[rowIndex].Tag = rowEntry.Entry;
        }
        RestoreProcessHistorySelection(selectedHistoryId);
        RestoreProcessHistoryScrollPosition(firstDisplayedRowIndex);
        processHistoryGrid.ResumeLayout();

        processHistoryStatusLabel.Text = $"Process history: {filteredEntries.Count} shown, {processHistoryEntries.Count} total";
    }

    private static void AddProcessHistoryRows(
        ProcessHistoryEntry entry,
        int depth,
        IReadOnlyDictionary<long, List<ProcessHistoryEntry>> childrenByParent,
        HashSet<long> ancestry,
        ICollection<ProcessHistoryRowEntry> rows,
        ISet<long> collapsedHistoryIds)
    {
        if (!ancestry.Add(entry.Id))
        {
            return;
        }

        childrenByParent.TryGetValue(entry.Id, out var children);
        var hasChildren = children is { Count: > 0 };
        var isExpanded = hasChildren && !collapsedHistoryIds.Contains(entry.Id);
        rows.Add(new ProcessHistoryRowEntry(entry, depth, hasChildren, isExpanded));

        if (hasChildren && isExpanded && children is not null)
        {
            foreach (var child in children)
            {
                if (ancestry.Contains(child.Id))
                {
                    continue;
                }

                AddProcessHistoryRows(child, depth + 1, childrenByParent, [.. ancestry], rows, collapsedHistoryIds);
            }
        }
    }

    private static string FormatHistoryName(ProcessHistoryRowEntry rowEntry)
    {
        var glyph = rowEntry.HasChildren
            ? rowEntry.IsExpanded ? "- " : "+ "
            : "  ";
        var prefix = rowEntry.Depth == 0
            ? glyph
            : string.Concat(Enumerable.Repeat("   ", rowEntry.Depth)) + glyph;
        return prefix + rowEntry.Entry.Name;
    }

    private static bool MatchesProcessHistory(ProcessHistoryEntry entry, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter))
        {
            return true;
        }

        return entry.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || entry.ProcessName.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || entry.ProcessId.ToString().Contains(filter, StringComparison.OrdinalIgnoreCase)
            || entry.Status.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || entry.BinaryPath.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || entry.CommandLine.Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private void ExpandProcessHistoryAncestors(long? parentHistoryId)
    {
        var currentParentHistoryId = parentHistoryId;
        while (currentParentHistoryId.HasValue)
        {
            var parentEntry = processHistoryEntries.FirstOrDefault(entry => entry.Id == currentParentHistoryId.Value);
            if (parentEntry is null)
            {
                break;
            }

            collapsedProcessHistoryIds.Remove(parentEntry.Id);
            currentParentHistoryId = parentEntry.ParentHistoryId;
        }
    }

    private static DateTime GetLatestProcessHistoryActivity(
        ProcessHistoryEntry entry,
        IReadOnlyDictionary<long, List<ProcessHistoryEntry>> childrenByParent,
        IDictionary<long, DateTime> cache,
        HashSet<long> ancestry)
    {
        if (cache.TryGetValue(entry.Id, out var cached))
        {
            return cached;
        }

        if (!ancestry.Add(entry.Id))
        {
            return entry.ClosedAtLocal ?? entry.StartedAtLocal;
        }

        var latest = entry.ClosedAtLocal ?? entry.StartedAtLocal;
        if (childrenByParent.TryGetValue(entry.Id, out var children))
        {
            foreach (var child in children)
            {
                var childLatest = GetLatestProcessHistoryActivity(child, childrenByParent, cache, ancestry);
                if (childLatest > latest)
                {
                    latest = childLatest;
                }
            }
        }

        ancestry.Remove(entry.Id);
        cache[entry.Id] = latest;
        return latest;
    }

    private void LogCall(string area, string action, string details)
    {
        Debug.WriteLine($"{DateTime.Now:HH:mm:ss} [{area}] {action}: {details}");
    }

    private long? GetSelectedProcessHistoryId()
    {
        return processHistoryGrid.CurrentRow?.Tag is ProcessHistoryEntry entry
            ? entry.Id
            : null;
    }

    private int GetFirstDisplayedProcessHistoryRowIndex()
    {
        try
        {
            return processHistoryGrid.FirstDisplayedScrollingRowIndex >= 0
                ? processHistoryGrid.FirstDisplayedScrollingRowIndex
                : 0;
        }
        catch
        {
            return 0;
        }
    }

    private void RestoreProcessHistorySelection(long? selectedHistoryId)
    {
        if (processHistoryGrid.Rows.Count == 0)
        {
            return;
        }

        if (selectedHistoryId is null)
        {
            return;
        }

        foreach (DataGridViewRow row in processHistoryGrid.Rows)
        {
            if (row.Tag is ProcessHistoryEntry entry && entry.Id == selectedHistoryId.Value)
            {
                processHistoryGrid.CurrentCell = processHistoryGrid[0, row.Index];
                return;
            }
        }
    }

    private void RestoreProcessHistoryScrollPosition(int firstDisplayedRowIndex)
    {
        if (processHistoryGrid.Rows.Count == 0)
        {
            return;
        }

        var safeIndex = Math.Max(0, Math.Min(firstDisplayedRowIndex, processHistoryGrid.Rows.Count - 1));
        try
        {
            processHistoryGrid.FirstDisplayedScrollingRowIndex = safeIndex;
        }
        catch
        {
        }
    }

    private void processHistoryGrid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || processHistoryGrid.Rows[e.RowIndex].Tag is not ProcessHistoryEntry entry)
        {
            return;
        }

        var statusColor = entry.Status.Equals("Closed", StringComparison.OrdinalIgnoreCase)
            ? Color.FromArgb(220, 110, 110)
            : entry.WasSeenAfterLaunch
                ? Color.FromArgb(94, 234, 150)
                : entry.ParentHistoryId is null
                    ? Color.FromArgb(96, 165, 250)
                : Color.FromArgb(203, 213, 225);

        var statusColumn = processHistoryGrid.Columns["HistoryStatus"];
        if (statusColumn is not null && e.ColumnIndex == statusColumn.Index)
        {
            e.CellStyle.ForeColor = statusColor;
        }

        var nameColumn = processHistoryGrid.Columns["HistoryName"];
        if (nameColumn is not null && e.ColumnIndex == nameColumn.Index)
        {
            e.CellStyle.ForeColor = statusColor;
        }
    }

    private void processHistoryGrid_CellMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.RowIndex < 0)
        {
            return;
        }

        var nameColumn = processHistoryGrid.Columns["HistoryName"];
        if (nameColumn is null || e.ColumnIndex != nameColumn.Index)
        {
            return;
        }

        if (!TryGetProcessHistoryRow(e.RowIndex, out var rowEntry) || !rowEntry.HasChildren)
        {
            return;
        }

        ToggleProcessHistoryExpansion(rowEntry.Entry.Id);
    }

    private bool TryGetProcessHistoryRow(int rowIndex, out ProcessHistoryRowEntry rowEntry)
    {
        if (rowIndex >= 0 && rowIndex < processHistoryRows.Count)
        {
            rowEntry = processHistoryRows[rowIndex];
            return true;
        }

        rowEntry = default;
        return false;
    }

    private void ToggleProcessHistoryExpansion(long historyId)
    {
        if (!collapsedProcessHistoryIds.Add(historyId))
        {
            collapsedProcessHistoryIds.Remove(historyId);
        }

        RenderProcessHistoryGrid();
    }

    private void UpdateTabButtons()
    {
        StyleTabButton(processesTabButton, mainTabControl.SelectedTab == processesTabPage);
        StyleTabButton(processHistoryTabButton, mainTabControl.SelectedTab == processHistoryTabPage);
        StyleTabButton(startupTabButton, mainTabControl.SelectedTab == startupTabPage);
        StyleTabButton(tasksTabButton, mainTabControl.SelectedTab == tasksTabPage);
        StyleTabButton(servicesTabButton, mainTabControl.SelectedTab == servicesTabPage);
        StyleTabButton(conTabButton, mainTabControl.SelectedTab == conTabPage);
    }

    private static void StyleTabButton(Button button, bool isSelected)
    {
        button.BackColor = isSelected ? Color.FromArgb(39, 39, 39) : Color.FromArgb(25, 25, 25);
        button.ForeColor = isSelected ? Color.FromArgb(245, 245, 245) : Color.FromArgb(170, 170, 175);
    }

    private void QueueDnsLookup(string remoteAddress, ConnectionRecordKey key)
    {
        if (string.IsNullOrWhiteSpace(remoteAddress)
            || remoteAddress is "0.0.0.0" or "::" or "::1" or "127.0.0.1"
            || !pendingDnsLookups.Add(remoteAddress))
        {
            return;
        }

        _ = Task.Run(() =>
        {
            string resolvedHost;
            try
            {
                resolvedHost = System.Net.Dns.GetHostEntry(remoteAddress).HostName;
                if (string.IsNullOrWhiteSpace(resolvedHost))
                {
                    resolvedHost = remoteAddress;
                }
            }
            catch
            {
                resolvedHost = remoteAddress;
            }

            pendingDnsLookups.Remove(remoteAddress);
            if (!connectionRecords.TryGetValue(key, out var existing))
            {
                return;
            }

            connectionRecords[key] = existing with { DnsName = resolvedHost };
            if (!IsHandleCreated)
            {
                return;
            }

            BeginInvoke(UpdateConnectionGrid);
        });
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

    private void SortProcessStates(List<ProcessNodeState> states)
    {
        states.Sort(CompareProcessStates);
    }

    private int CompareProcessStates(ProcessNodeState? left, ProcessNodeState? right)
    {
        if (ReferenceEquals(left, right))
        {
            return 0;
        }

        if (left is null)
        {
            return -1;
        }

        if (right is null)
        {
            return 1;
        }

        var result = processSortColumnName switch
        {
            "NameColumn" => CompareText(left.Name, right.Name),
            "ProcessNameColumn" => CompareText(left.ProcessName, right.ProcessName),
            "PidColumn" => left.ProcessId.CompareTo(right.ProcessId),
            "CpuColumn" => left.CpuPercent.CompareTo(right.CpuPercent),
            "MemoryColumn" => left.WorkingSetBytes.CompareTo(right.WorkingSetBytes),
            "GpuColumn" => left.GpuPercent.CompareTo(right.GpuPercent),
            "DiskColumn" => left.DiskBytesPerSecond.CompareTo(right.DiskBytesPerSecond),
            "NetworkColumn" => CompareNetwork(left, right),
            "StatusColumn" => CompareStatus(left, right),
            "CommandLineColumn" => CompareText(left.CommandLine, right.CommandLine),
            _ => 0
        };

        if (result == 0)
        {
            result = left.IsExited.CompareTo(right.IsExited);
        }

        if (result == 0)
        {
            result = CompareText(left.Name, right.Name);
        }

        if (result == 0)
        {
            result = left.ProcessId.CompareTo(right.ProcessId);
        }

        return processSortOrder == SortOrder.Descending
            ? -result
            : result;
    }

    private static int CompareText(string? left, string? right)
    {
        return string.Compare(left ?? string.Empty, right ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    private static int CompareNetwork(ProcessNodeState left, ProcessNodeState right)
    {
        var throughputComparison = left.NetworkBytesPerSecond.CompareTo(right.NetworkBytesPerSecond);
        return throughputComparison != 0
            ? throughputComparison
            : left.NetworkConnectionCount.CompareTo(right.NetworkConnectionCount);
    }

    private static int CompareStatus(ProcessNodeState left, ProcessNodeState right)
    {
        if (left.IsExited != right.IsExited)
        {
            return left.IsExited.CompareTo(right.IsExited);
        }

        if (left.WasSeenAfterLaunch != right.WasSeenAfterLaunch)
        {
            return right.WasSeenAfterLaunch.CompareTo(left.WasSeenAfterLaunch);
        }

        var leftIsRoot = left.ParentProcessId is null || left.ParentProcessId == left.ProcessId;
        var rightIsRoot = right.ParentProcessId is null || right.ParentProcessId == right.ProcessId;
        return rightIsRoot.CompareTo(leftIsRoot);
    }

    private void UpdateProcessSortGlyphs()
    {
        foreach (DataGridViewColumn column in processGrid.Columns)
        {
            column.HeaderCell.SortGlyphDirection = processSortColumnName == column.Name
                ? processSortOrder
                : SortOrder.None;
        }
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

internal sealed class BorderlessTabControl : TabControl
{
    private const int TcmAdjustRect = 0x1328;

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == TcmAdjustRect && !DesignMode)
        {
            m.Result = 1;
            return;
        }

        base.WndProc(ref m);
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

internal readonly record struct ProcessHistoryRowEntry(ProcessHistoryEntry Entry, int Depth, bool HasChildren, bool IsExpanded);

internal sealed class ProcessHistoryEntry
{
    public ProcessHistoryEntry(long id, int processId, int? parentProcessId, long? parentHistoryId, string name, string processName, string binaryPath, string commandLine, DateTime startedAtLocal, DateTime? closedAtLocal, bool wasSeenAfterLaunch, string status)
    {
        Id = id;
        ProcessId = processId;
        ParentProcessId = parentProcessId;
        ParentHistoryId = parentHistoryId;
        Name = name;
        ProcessName = processName;
        BinaryPath = binaryPath;
        CommandLine = commandLine;
        StartedAtLocal = startedAtLocal;
        ClosedAtLocal = closedAtLocal;
        WasSeenAfterLaunch = wasSeenAfterLaunch;
        Status = status;
    }

    public long Id { get; }

    public int ProcessId { get; }

    public int? ParentProcessId { get; set; }

    public long? ParentHistoryId { get; set; }

    public string Name { get; set; }

    public string ProcessName { get; set; }

    public string BinaryPath { get; set; }

    public string CommandLine { get; set; }

    public DateTime StartedAtLocal { get; }

    public DateTime? ClosedAtLocal { get; set; }

    public bool WasSeenAfterLaunch { get; }

    public string Status { get; set; }
}

internal readonly record struct ConnectionRecordKey(int ProcessId, string Protocol, string RemoteAddress, int RemotePort);

internal sealed record ConnectionRecord(
    ConnectionRecordKey Key,
    int ProcessId,
    string BinaryName,
    string BinaryPath,
    string DnsName,
    string RemoteAddress,
    string Upload,
    string Download,
    DateTime FirstSeenLocal,
    DateTime LastSeenLocal);

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
    double NetworkUploadBytesPerSecond,
    double NetworkDownloadBytesPerSecond,
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
                    metrics.NetworkUploadBytesPerSecond,
                    metrics.NetworkDownloadBytesPerSecond,
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

internal readonly record struct ProcessRuntimeMetrics(
    double DiskBytesPerSecond,
    double NetworkUploadBytesPerSecond,
    double NetworkDownloadBytesPerSecond,
    double NetworkBytesPerSecond,
    int ConnectionCount,
    double GpuPercent);

internal sealed class ProcessRuntimeMetricsProvider : IDisposable
{
    private readonly object sync = new();
    private readonly Dictionary<int, (ulong Bytes, DateTime SampleUtc)> lastDiskSamples = [];
    private readonly Dictionary<string, CounterSample> lastGpuSamples = new(StringComparer.OrdinalIgnoreCase);
    private readonly NetworkUsageMonitor networkUsageMonitor = new();
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
            var networkRates = networkUsageMonitor.ReadCurrentRates();
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
                    networkRates.TryGetValue(pid, out var networkRate);

                    metrics[pid] = new ProcessRuntimeMetrics(
                        diskRate,
                        networkRate.UploadBytesPerSecond,
                        networkRate.DownloadBytesPerSecond,
                        networkRate.TotalBytesPerSecond,
                        connectionCount,
                        gpuPercent);
                }
            }

            networkUsageMonitor.CleanupState(metrics.Keys);
            CleanupState(metrics.Keys);
            return metrics;
        }
    }

    public void Dispose()
    {
        networkUsageMonitor.Dispose();
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

internal readonly record struct ProcessNetworkRate(double UploadBytesPerSecond, double DownloadBytesPerSecond, double TotalBytesPerSecond);

internal sealed class NetworkUsageMonitor : IDisposable
{
    private readonly object sync = new();
    private readonly Dictionary<int, (ulong UploadBytes, ulong DownloadBytes)> cumulativeBytes = [];
    private readonly Dictionary<int, (ulong UploadBytes, ulong DownloadBytes, DateTime SampleUtc)> lastSamples = [];
    private readonly string sessionName = $"TaskManager-Network-{Environment.ProcessId}";
    private TraceEventSession? session;
    private Task? processingTask;
    private bool started;

    public IReadOnlyDictionary<int, ProcessNetworkRate> ReadCurrentRates()
    {
        EnsureStarted();

        lock (sync)
        {
            var sampledAtUtc = DateTime.UtcNow;
            var rates = new Dictionary<int, ProcessNetworkRate>(cumulativeBytes.Count);

            foreach (var (pid, totals) in cumulativeBytes)
            {
                var uploadRate = 0d;
                var downloadRate = 0d;

                if (lastSamples.TryGetValue(pid, out var previous))
                {
                    var elapsedSeconds = (sampledAtUtc - previous.SampleUtc).TotalSeconds;
                    if (elapsedSeconds > 0)
                    {
                        if (totals.UploadBytes >= previous.UploadBytes)
                        {
                            uploadRate = (totals.UploadBytes - previous.UploadBytes) / elapsedSeconds;
                        }

                        if (totals.DownloadBytes >= previous.DownloadBytes)
                        {
                            downloadRate = (totals.DownloadBytes - previous.DownloadBytes) / elapsedSeconds;
                        }
                    }
                }

                lastSamples[pid] = (totals.UploadBytes, totals.DownloadBytes, sampledAtUtc);
                rates[pid] = new ProcessNetworkRate(uploadRate, downloadRate, uploadRate + downloadRate);
            }

            return rates;
        }
    }

    public void CleanupState(IEnumerable<int> activePids)
    {
        lock (sync)
        {
            var activePidSet = activePids.ToHashSet();
            foreach (var stalePid in cumulativeBytes.Keys.Where(pid => !activePidSet.Contains(pid)).ToArray())
            {
                cumulativeBytes.Remove(stalePid);
                lastSamples.Remove(stalePid);
            }
        }
    }

    public void Dispose()
    {
        try
        {
            session?.Dispose();
        }
        catch
        {
        }

        if (processingTask is { IsCompleted: false })
        {
            try
            {
                processingTask.Wait(TimeSpan.FromSeconds(1));
            }
            catch
            {
            }
        }
    }

    private void EnsureStarted()
    {
        if (started || !RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return;
        }

        started = true;

        try
        {
            if (TraceEventSession.IsElevated() == false)
            {
                return;
            }

            session = new TraceEventSession(sessionName, null)
            {
                StopOnDispose = true
            };
            session.EnableKernelProvider(KernelTraceEventParser.Keywords.NetworkTCPIP);

            processingTask = Task.Run(() =>
            {
                try
                {
                    session.Source.Kernel.TcpIpSend += data => RecordTransfer(data.ProcessID, checked((uint)data.size), isUpload: true);
                    session.Source.Kernel.TcpIpRecv += data => RecordTransfer(data.ProcessID, checked((uint)data.size), isUpload: false);
                    session.Source.Process();
                }
                catch
                {
                }
            });
        }
        catch
        {
            try
            {
                session?.Dispose();
            }
            catch
            {
            }

            session = null;
        }
    }

    private void RecordTransfer(int processId, uint bytes, bool isUpload)
    {
        if (processId <= 0 || bytes == 0)
        {
            return;
        }

        lock (sync)
        {
            cumulativeBytes.TryGetValue(processId, out var totals);
            cumulativeBytes[processId] = isUpload
                ? (totals.UploadBytes + bytes, totals.DownloadBytes)
                : (totals.UploadBytes, totals.DownloadBytes + bytes);
        }
    }
}

internal sealed record StartupEntry(string Name, string Scope, string Type, string Location, string Command, string LastRunTime, string NextExecutionTime);

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

    public static IEnumerable<StartupEntry> ReadScheduledTaskStartupEntries()
    {
        return ReadScheduledStartupTasks();
    }

    private static IEnumerable<StartupEntry> ReadStandardRegistryLocations()
    {
        var locations = new (RegistryHive Hive, RegistryView View, string SubKeyPath, string Scope, string Type)[]
        {
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\Run", "Current user", "Registry Run"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "Current user", "Registry RunOnce"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\RunServices", "Current user", "Registry RunServices"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", "Current user", "Registry RunServicesOnce"),
            (RegistryHive.CurrentUser, RegistryView.Default, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "Current user", "Registry Policy Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\Run", "All users x64", "Registry Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "All users x64", "Registry RunOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\RunServices", "All users x64", "Registry RunServices"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", "All users x64", "Registry RunServicesOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "All users x64", "Registry Policy Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\Run", "All users x86", "Registry Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\RunOnce", "All users x86", "Registry RunOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\RunServices", "All users x86", "Registry RunServices"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", "All users x86", "Registry RunServicesOnce"),
            (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "All users x86", "Registry Policy Run"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "System", "Winlogon Shell"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "System", "Winlogon Userinit"),
            (RegistryHive.LocalMachine, RegistryView.Registry64, @"System\CurrentControlSet\Control\Terminal Server\Wds\rdpwd", "System", "Terminal Server StartupPrograms")
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
            if (type == "Winlogon Shell" && !valueName.Equals("Shell", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (type == "Winlogon Userinit" && !valueName.Equals("Userinit", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (type == "Terminal Server StartupPrograms" && !valueName.Equals("StartupPrograms", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var nextExecution = type switch
            {
                "Registry Run" or "Registry RunServices" or "Registry Policy Run" => "At sign-in",
                "Registry RunOnce" or "Registry RunServicesOnce" => "Next sign-in",
                "Winlogon Shell" or "Winlogon Userinit" => "At sign-in",
                "Terminal Server StartupPrograms" => "At session start",
                _ => "Unavailable"
            };

            yield return new StartupEntry(
                valueName,
                scope,
                type,
                $@"{hive} [{view}]\{subKeyPath}",
                key.GetValue(valueName)?.ToString() ?? string.Empty,
                "Unavailable",
                nextExecution);
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
            yield return new StartupEntry(Path.GetFileNameWithoutExtension(file), scope, "Startup folder", path, file, "Unavailable", "At sign-in");
        }
    }

    private static IEnumerable<StartupEntry> ReadScheduledStartupTasks()
    {
        ProcessStartInfo startInfo = new()
        {
            FileName = "schtasks.exe",
            Arguments = "/query /fo csv /v",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            yield break;
        }

        var output = process.StandardOutput.ReadToEnd();
        process.WaitForExit(5000);
        if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
        {
            yield break;
        }

        var rows = ParseCsv(output);
        if (rows.Count < 2)
        {
            yield break;
        }

        var headers = rows[0];
        var headerMap = headers
            .Select((header, index) => new { header, index })
            .ToDictionary(item => item.header, item => item.index, StringComparer.OrdinalIgnoreCase);

        for (var index = 1; index < rows.Count; index++)
        {
            var row = rows[index];
            var scheduleType = GetCsvValue(row, headerMap, "Schedule Type");
            if (!(scheduleType.Contains("logon", StringComparison.OrdinalIgnoreCase)
                || scheduleType.Contains("startup", StringComparison.OrdinalIgnoreCase)
                || scheduleType.Contains("boot", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var taskName = GetCsvValue(row, headerMap, "TaskName");
            var command = GetCsvValue(row, headerMap, "Task To Run");
            if (string.IsNullOrWhiteSpace(command))
            {
                command = GetCsvValue(row, headerMap, "Actions");
            }

            yield return new StartupEntry(
                taskName,
                "Task Scheduler",
                "Scheduled Task",
                @"Task Scheduler Library",
                command,
                NormalizeTaskTime(GetCsvValue(row, headerMap, "Last Run Time")),
                NormalizeTaskTime(GetCsvValue(row, headerMap, "Next Run Time")));
        }
    }

    private static string NormalizeTaskTime(string value)
    {
        return string.IsNullOrWhiteSpace(value) || value.Equals("N/A", StringComparison.OrdinalIgnoreCase)
            ? "Unavailable"
            : value;
    }

    private static string GetCsvValue(IReadOnlyList<string> row, IReadOnlyDictionary<string, int> headerMap, string header)
    {
        return headerMap.TryGetValue(header, out var index) && index < row.Count
            ? row[index]
            : string.Empty;
    }

    private static List<List<string>> ParseCsv(string csv)
    {
        List<List<string>> rows = [];
        List<string> currentRow = [];
        System.Text.StringBuilder currentValue = new();
        var inQuotes = false;

        for (var index = 0; index < csv.Length; index++)
        {
            var ch = csv[index];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (index + 1 < csv.Length && csv[index + 1] == '"')
                    {
                        currentValue.Append('"');
                        index++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    currentValue.Append(ch);
                }

                continue;
            }

            if (ch == '"')
            {
                inQuotes = true;
            }
            else if (ch == ',')
            {
                currentRow.Add(currentValue.ToString());
                currentValue.Clear();
            }
            else if (ch == '\r')
            {
            }
            else if (ch == '\n')
            {
                currentRow.Add(currentValue.ToString());
                currentValue.Clear();
                if (currentRow.Count > 1 || (currentRow.Count == 1 && !string.IsNullOrWhiteSpace(currentRow[0])))
                {
                    rows.Add(currentRow);
                }

                currentRow = [];
            }
            else
            {
                currentValue.Append(ch);
            }
        }

        if (currentValue.Length > 0 || currentRow.Count > 0)
        {
            currentRow.Add(currentValue.ToString());
            rows.Add(currentRow);
        }

        return rows;
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
