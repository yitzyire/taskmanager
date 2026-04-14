using System.Runtime.InteropServices;

namespace NewWindowsTaskManager;

internal interface IProcessMetadataProvider
{
    string TryGetCommandLine(int processId);
}

internal static class ProcessMetadataProvider
{
    public static IProcessMetadataProvider Current { get; } =
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? new WindowsProcessMetadataProvider()
            : new UnsupportedProcessMetadataProvider();
}

internal sealed class WindowsProcessMetadataProvider : IProcessMetadataProvider
{
    public string TryGetCommandLine(int processId)
    {
        return NativeProcessMethods.TryGetProcessCommandLine(processId) ?? "Unavailable";
    }
}

internal sealed class UnsupportedProcessMetadataProvider : IProcessMetadataProvider
{
    public string TryGetCommandLine(int processId)
    {
        return "Unavailable on this platform";
    }
}
