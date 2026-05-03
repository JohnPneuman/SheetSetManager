using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using SheetSetEditor.Services;

namespace SheetSetEditor;

public partial class App : Application
{
    private static Mutex? _mutex;

    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    private const int SW_RESTORE = 9;

    protected override void OnStartup(StartupEventArgs e)
    {
        _mutex = new Mutex(true, "BoekSolutions_SheetSetEditor_SingleInstance", out bool createdNew);
        if (!createdNew)
        {
            var me = Process.GetCurrentProcess();
            var existing = Process.GetProcessesByName(me.ProcessName)
                .FirstOrDefault(p => p.Id != me.Id && p.MainWindowHandle != IntPtr.Zero);
            if (existing != null)
            {
                ShowWindow(existing.MainWindowHandle, SW_RESTORE);
                SetForegroundWindow(existing.MainWindowHandle);
            }
            Shutdown();
            return;
        }

        // Apply persisted settings before the first window appears
        LocalizationService.Instance.Language = EditorSettingsService.GetLanguage();
        ApplyTheme(EditorSettingsService.GetDarkMode());

        base.OnStartup(e);
        var window = new MainWindow();
        window.Show();

        if (e.Args.Length > 0)
            window.OpenFromCommandLine(e.Args[0]);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mutex?.ReleaseMutex();
        base.OnExit(e);
    }

    public static void ApplyTheme(bool dark)
    {
        var dict = Application.Current.Resources.MergedDictionaries;
        dict.Clear();
        var uri = dark
            ? new Uri("Resources/DarkTheme.xaml",  UriKind.Relative)
            : new Uri("Resources/LightTheme.xaml", UriKind.Relative);
        dict.Add(new ResourceDictionary { Source = uri });
    }
}
