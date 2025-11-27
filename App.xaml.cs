using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace ZibllWindows
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Attach global exception handlers early
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            base.OnStartup(e);
        }

        private void LogAndReport(string title, Exception ex)
        {
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ActivationTool");
                Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "activationtool_error.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}: {ex}\n\n");

                // Use a simple MessageBox here to ensure no dependency on UI libraries
                MessageBox.Show($"{title}: {ex.Message}\n\nLog: {logPath}", "ActivationTool - Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ActivationTool");
                    Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "activationtool_error.log");
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
