using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TaskManager;

internal static class NativeProcessMethods
{
    private const uint SnapshotProcess = 0x00000002;
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
}

