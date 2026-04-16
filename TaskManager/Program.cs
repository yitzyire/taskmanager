namespace TaskManager;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, args) => ShowFatalError("UI thread exception", args.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            ShowFatalError("Unhandled exception", args.ExceptionObject as Exception ?? new Exception("Unknown fatal error."));

        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }

    private static void ShowFatalError(string title, Exception exception)
    {
        try
        {
            MessageBox.Show(
                $"{title}\n\n{exception.Message}",
                "TaskManager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        catch
        {
        }
    }
}
