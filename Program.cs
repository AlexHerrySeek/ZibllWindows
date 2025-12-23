using System;
using System.IO;
using System.Windows;

namespace ZibllWindows
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                // Check for uninstall argument
                if (args.Length > 0 && args[0].ToLower() == "/uninstall")
                {
                    ShowUninstallDialog();
                    return;
                }

                var app = new App();
                app.InitializeComponent();
                app.Run();
            }
            catch (Exception ex)
            {
                try
                {
                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ZibllWindows");
                    Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "zibllwindows_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Startup exception: {ex}\n\n");
                    System.Windows.MessageBox.Show($"Lỗi khởi động: {ex.Message}\n\nLog: {logPath}", "ZibllWindows - Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch { }
            }
        }

        private static void ShowUninstallDialog()
        {
            try
            {
                var app = new App();
                app.InitializeComponent();
                var installPage = new InstallPage();
                installPage.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi khi mở giao diện gỡ cài đặt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
