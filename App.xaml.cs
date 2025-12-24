using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace ZibllWindows
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Attach global exception handlers early
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            base.OnStartup(e);

            // Check if installed, if not show InstallPage first
            CheckInstallationAndStart();
        }

        private void CheckInstallationAndStart()
        {
            try
            {
                string installPath = @"C:\Program Files\ZibllWindows";
                bool isInstalled = Directory.Exists(installPath) && 
                                   File.Exists(Path.Combine(installPath, "ZibllWindows.exe"));

                // Get current executable path
                string currentExe = System.Reflection.Assembly.GetExecutingAssembly().Location;
                if (string.IsNullOrEmpty(currentExe))
                {
                    currentExe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                }
                
                string currentDir = Path.GetDirectoryName(currentExe) ?? "";
                bool isRunningFromInstall = !string.IsNullOrEmpty(currentDir) && 
                                           currentDir.Equals(installPath, StringComparison.OrdinalIgnoreCase);

                // If not installed or not running from install directory, show InstallPage
                if (!isInstalled || !isRunningFromInstall)
                {
                    var installPage = new InstallPage();
                    installPage.ShowDialog();
                    
                    // After install page closes, check if we should continue
                    if (!isInstalled && !Directory.Exists(installPath))
                    {
                        // User skipped installation, continue anyway
                    }
                }
            }
            catch (Exception ex)
            {
                // If error checking, continue to MainWindow anyway
                System.Diagnostics.Debug.WriteLine($"Error checking installation: {ex.Message}");
            }
        }

        private void LogAndReport(string title, Exception ex)
        {
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ZibllWindows");
                Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "alexherry_error.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}: {ex}\n\n");

                // Use a simple MessageBox here to ensure no dependency on UI libraries
                System.Windows.MessageBox.Show($"{title}: {ex.Message}\n\nLog: {logPath}", "ZibllWindows - Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch
            {
                // Swallow any logging errors to avoid recursive crashes
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogAndReport("Lỗi chưa xử lý (UI Thread)", e.Exception);
            e.Handled = true;
            // After reporting, shut down with error code
            Shutdown(1);
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogAndReport("Lỗi chưa xử lý (AppDomain)", ex);
            }
            else
            {
                try
                {
                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ZibllWindows");
                    Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "alexherry_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Unhandled (AppDomain): {e.ExceptionObject}\n\n");
                }
                catch { }
            }
            Shutdown(1);
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogAndReport("Lỗi Task chưa quan sát", e.Exception);
            e.SetObserved();
        }
    }
}
