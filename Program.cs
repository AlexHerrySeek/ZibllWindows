using System;
using System.IO;

namespace ZibllWindows
{
    public class Program
    {
        [STAThread]
        public static void Main()
        {
            try
            {
                var app = new App();
                app.InitializeComponent();
                app.Run();
            }
            catch (Exception ex)
            {
                try
                {
                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ActivationTool");
                    Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "activationtool_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Startup exception: {ex}\n\n");
                    System.Windows.MessageBox.Show($"Lỗi khởi động: {ex.Message}\n\nLog: {logPath}", "ActivationTool - Lỗi", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
                catch { }
            }
        }
    }
}
