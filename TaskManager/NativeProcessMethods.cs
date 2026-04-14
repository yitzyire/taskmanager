using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TaskManager;

internal static class NativeProcessMethods
{
    private const int AddressFamilyInterNetwork = 2;
    private const int AddressFamilyInterNetworkV6 = 23;
    private const uint SnapshotProcess = 0x00000002;
    private const int TcpTableOwnerPidAll = 5;
    private const int UdpTableOwnerPid = 1;
    private const uint QueryLimitedInformation = 0x1000;
    private const uint VirtualMemoryRead = 0x0010;

    public static Dictionary<int, int> ReadParentProcessMap()
    {
        var snapshot = CreateToolhelp32Snapshot(SnapshotProcess, 0);
        if (snapshot == IntPtr.Zero || snapshot == new IntPtr(-1))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        try
        {
            var entry = new PROCESSENTRY32
            {
                dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32>()
            };

            var map = new Dictionary<int, int>();
            if (!Process32First(snapshot, ref entry))
            {
                return map;
            }

            do
            {
                map[(int)entry.th32ProcessID] = (int)entry.th32ParentProcessID;
                entry.dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32>();
            }
            while (Process32Next(snapshot, ref entry));

            return map;
        }
        finally
        {
            CloseHandle(snapshot);
        }
    }

    public static string? TryGetProcessPath(int processId)
    {
        var handle = OpenProcess(QueryLimitedInformation, false, processId);
        if (handle == IntPtr.Zero)
        {
            return TryReadFromManagedProcess(processId);
        }

        try
        {
            var capacity = 1024;
            var builder = new System.Text.StringBuilder(capacity);
            if (QueryFullProcessImageName(handle, 0, builder, ref capacity))
            {
                return builder.ToString();
            }
        }
        finally
        {
            CloseHandle(handle);
        }

        return TryReadFromManagedProcess(processId);
    }

    public static string? TryGetProcessCommandLine(int processId)
    {
        var handle = OpenProcess(QueryLimitedInformation | VirtualMemoryRead, false, processId);
        if (handle == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var processInformation = new PROCESS_BASIC_INFORMATION();
            var status = NtQueryInformationProcess(
                handle,
                0,
                ref processInformation,
                Marshal.SizeOf<PROCESS_BASIC_INFORMATION>(),
                out _);

            if (status != 0 || processInformation.PebBaseAddress == IntPtr.Zero)
            {
                return null;
            }

            var processParametersAddress = ReadIntPtr(
                handle,
                processInformation.PebBaseAddress + (IntPtr.Size == 8 ? 0x20 : 0x10));

            if (processParametersAddress == IntPtr.Zero)
            {
                return null;
            }

            var commandLineUnicode = ReadStructure<UNICODE_STRING>(
                handle,
                processParametersAddress + (IntPtr.Size == 8 ? 0x70 : 0x40));

            if (commandLineUnicode.Length <= 0 || commandLineUnicode.Buffer == IntPtr.Zero)
            {
                return null;
            }

            var buffer = new byte[commandLineUnicode.Length];
            if (!ReadProcessMemory(handle, commandLineUnicode.Buffer, buffer, buffer.Length, out _))
            {
                return null;
            }

            return System.Text.Encoding.Unicode.GetString(buffer);
        }
        catch
        {
            return null;
        }
        finally
        {
            CloseHandle(handle);
        }
    }

    public static string TryReadServiceStartType(string serviceName)
    {
        using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}");
        return key?.GetValue("Start") switch
        {
            0 => "Boot",
            1 => "System",
            2 => "Automatic",
            3 => "Manual",
            4 => "Disabled",
            _ => "Unknown"
        };
    }

    public static bool TryGetProcessDiskBytes(int processId, out ulong totalBytes)
    {
        totalBytes = 0;
        var handle = OpenProcess(QueryLimitedInformation, false, processId);
        if (handle == IntPtr.Zero)
        {
            return false;
        }

        try
        {
            if (!GetProcessIoCounters(handle, out var counters))
            {
                return false;
            }

            totalBytes = counters.ReadTransferCount + counters.WriteTransferCount;
            return true;
        }
        finally
        {
            CloseHandle(handle);
        }
    }

    public static Dictionary<int, int> ReadNetworkConnectionCounts()
    {
        var counts = new Dictionary<int, int>();
        ReadTcpConnections(AddressFamilyInterNetwork, counts);
        ReadTcpConnections(AddressFamilyInterNetworkV6, counts);
        ReadUdpListeners(AddressFamilyInterNetwork, counts);
        ReadUdpListeners(AddressFamilyInterNetworkV6, counts);
        return counts;
    }

    private static string? TryReadFromManagedProcess(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return process.MainModule?.FileName;
        }
        catch
        {
            return null;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint processAccess, bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [Out] byte[] lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool QueryFullProcessImageName(IntPtr hProcess, int flags, System.Text.StringBuilder text, ref int size);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetProcessIoCounters(IntPtr hProcess, out IO_COUNTERS ioCounters);

    [DllImport("iphlpapi.dll", SetLastError = true)]
    private static extern uint GetExtendedTcpTable(
        IntPtr pTcpTable,
        ref int pdwSize,
        bool bOrder,
        int ulAf,
        int tableClass,
        uint reserved);

    [DllImport("iphlpapi.dll", SetLastError = true)]
    private static extern uint GetExtendedUdpTable(
        IntPtr pUdpTable,
        ref int pdwSize,
        bool bOrder,
        int ulAf,
        int tableClass,
        uint reserved);

    [DllImport("ntdll.dll")]
    private static extern int NtQueryInformationProcess(
        IntPtr processHandle,
        int processInformationClass,
        ref PROCESS_BASIC_INFORMATION processInformation,
        int processInformationLength,
        out int returnLength);

    private static IntPtr ReadIntPtr(IntPtr processHandle, IntPtr address)
    {
        var bytes = new byte[IntPtr.Size];
        if (!ReadProcessMemory(processHandle, address, bytes, bytes.Length, out _))
        {
            return IntPtr.Zero;
        }

        return IntPtr.Size == 8
            ? new IntPtr(BitConverter.ToInt64(bytes, 0))
            : new IntPtr(BitConverter.ToInt32(bytes, 0));
    }

    private static T ReadStructure<T>(IntPtr processHandle, IntPtr address) where T : struct
    {
        var size = Marshal.SizeOf<T>();
        var bytes = new byte[size];
        if (!ReadProcessMemory(processHandle, address, bytes, bytes.Length, out _))
        {
            return default;
        }

        var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
        try
        {
            return Marshal.PtrToStructure<T>(handle.AddrOfPinnedObject());
        }
        finally
        {
            handle.Free();
        }
    }

    private static void ReadTcpConnections(int addressFamily, IDictionary<int, int> counts)
    {
        var size = 0;
        _ = GetExtendedTcpTable(IntPtr.Zero, ref size, true, addressFamily, TcpTableOwnerPidAll, 0);
        if (size <= 0)
        {
            return;
        }

        var buffer = Marshal.AllocHGlobal(size);
        try
        {
            if (GetExtendedTcpTable(buffer, ref size, true, addressFamily, TcpTableOwnerPidAll, 0) != 0)
            {
                return;
            }

            var rowCount = Marshal.ReadInt32(buffer);
            var rowPtr = buffer + sizeof(int);

            if (addressFamily == AddressFamilyInterNetwork)
            {
                var rowSize = Marshal.SizeOf<MIB_TCPROW_OWNER_PID>();
                for (var index = 0; index < rowCount; index++)
                {
                    var row = Marshal.PtrToStructure<MIB_TCPROW_OWNER_PID>(rowPtr + (index * rowSize));
                    IncrementCount(counts, unchecked((int)row.owningPid));
                }
            }
            else
            {
                var rowSize = Marshal.SizeOf<MIB_TCP6ROW_OWNER_PID>();
                for (var index = 0; index < rowCount; index++)
                {
                    var row = Marshal.PtrToStructure<MIB_TCP6ROW_OWNER_PID>(rowPtr + (index * rowSize));
                    IncrementCount(counts, unchecked((int)row.owningPid));
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static void ReadUdpListeners(int addressFamily, IDictionary<int, int> counts)
    {
        var size = 0;
        _ = GetExtendedUdpTable(IntPtr.Zero, ref size, true, addressFamily, UdpTableOwnerPid, 0);
        if (size <= 0)
        {
            return;
        }

        var buffer = Marshal.AllocHGlobal(size);
        try
        {
            if (GetExtendedUdpTable(buffer, ref size, true, addressFamily, UdpTableOwnerPid, 0) != 0)
            {
                return;
            }

            var rowCount = Marshal.ReadInt32(buffer);
            var rowPtr = buffer + sizeof(int);

            if (addressFamily == AddressFamilyInterNetwork)
            {
                var rowSize = Marshal.SizeOf<MIB_UDPROW_OWNER_PID>();
                for (var index = 0; index < rowCount; index++)
                {
                    var row = Marshal.PtrToStructure<MIB_UDPROW_OWNER_PID>(rowPtr + (index * rowSize));
                    IncrementCount(counts, unchecked((int)row.owningPid));
                }
            }
            else
            {
                var rowSize = Marshal.SizeOf<MIB_UDP6ROW_OWNER_PID>();
                for (var index = 0; index < rowCount; index++)
                {
                    var row = Marshal.PtrToStructure<MIB_UDP6ROW_OWNER_PID>(rowPtr + (index * rowSize));
                    IncrementCount(counts, unchecked((int)row.owningPid));
                }
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static void IncrementCount(IDictionary<int, int> counts, int pid)
    {
        if (pid <= 0)
        {
            return;
        }

        counts[pid] = counts.TryGetValue(pid, out var current)
            ? current + 1
            : 1;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PROCESSENTRY32
    {
        public uint dwSize;
        public uint cntUsage;
        public uint th32ProcessID;
        public IntPtr th32DefaultHeapID;
        public uint th32ModuleID;
        public uint cntThreads;
        public uint th32ParentProcessID;
        public int pcPriClassBase;
        public uint dwFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szExeFile;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_BASIC_INFORMATION
    {
        public IntPtr Reserved1;
        public IntPtr PebBaseAddress;
        public IntPtr Reserved2_0;
        public IntPtr Reserved2_1;
        public IntPtr UniqueProcessId;
        public IntPtr InheritedFromUniqueProcessId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct UNICODE_STRING
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IO_COUNTERS
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MIB_TCPROW_OWNER_PID
    {
        public uint state;
        public uint localAddr;
        public uint localPort;
        public uint remoteAddr;
        public uint remotePort;
        public uint owningPid;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MIB_TCP6ROW_OWNER_PID
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public byte[] localAddr;
        public uint localScopeId;
        public uint localPort;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public byte[] remoteAddr;
        public uint remoteScopeId;
        public uint remotePort;
        public uint state;
        public uint owningPid;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MIB_UDPROW_OWNER_PID
    {
        public uint localAddr;
        public uint localPort;
        public uint owningPid;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MIB_UDP6ROW_OWNER_PID
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public byte[] localAddr;
        public uint localScopeId;
        public uint localPort;
        public uint owningPid;
    }
}

