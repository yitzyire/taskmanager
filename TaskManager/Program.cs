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
            var logPath = Path.Combine(AppContext.BaseDirectory, "crash.log");
            File.AppendAllText(
                logPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}{Environment.NewLine}{exception}{Environment.NewLine}{Environment.NewLine}");

            MessageBox.Show(
                $"{title}\n\n{exception.Message}\n\nA crash log was written to:\n{logPath}",
                "New Windows Task Manager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        catch
        {
        }
    }
}
