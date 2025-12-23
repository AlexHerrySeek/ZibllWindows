using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;

namespace ZibllWindows
{
    public partial class UninstallToolsPage : Wpf.Ui.Controls.FluentWindow
    {
        private readonly string[] _bloatwarePackages = new string[]
        {
            "Microsoft.Microsoft3DViewer",
            "Microsoft.AppConnector",
            "Microsoft.BingFinance",
            "Microsoft.BingNews",
            "Microsoft.BingSports",
            "Microsoft.BingWeather",
            "Microsoft.GetHelp",
            "Microsoft.Getstarted",
            "Microsoft.Messaging",
            "Microsoft.Microsoft3DViewer",
            "Microsoft.MicrosoftOfficeHub",
            "Microsoft.MicrosoftSolitaireCollection",
            "Microsoft.NetworkSpeedTest",
            "Microsoft.News",
            "Microsoft.Office.Sway",
            "Microsoft.OneConnect",
            "Microsoft.People",
            "Microsoft.Print3D",
            "Microsoft.SkypeApp",
            "Microsoft.Wallet",
            "Microsoft.WindowsAlarms",
            "Microsoft.WindowsCamera",
            "microsoft.windowscommunicationsapps",
            "Microsoft.WindowsFeedbackHub",
            "Microsoft.WindowsMaps",
            "Microsoft.WindowsSoundRecorder",
            "Microsoft.XboxApp",
            "Microsoft.XboxGameOverlay",
            "Microsoft.XboxGamingOverlay",
            "Microsoft.XboxIdentityProvider",
            "Microsoft.XboxSpeechToTextOverlay",
            "Microsoft.YourPhone",
            "Microsoft.ZuneMusic",
            "Microsoft.ZuneVideo"
        };

        public UninstallToolsPage()
        {
            InitializeComponent();
        }

        private async void BtnBloatware_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LblStatus.Text = "Đang xóa Bloatware...";
                LblStatus.Foreground = System.Windows.Media.Brushes.Orange;

                int removed = 0;
                foreach (var package in _bloatwarePackages)
                {
                    try
                    {
                        var process = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = "powershell.exe",
                                Arguments = $"-Command \"Get-AppxPackage {package} | Remove-AppxPackage\"",
                                UseShellExecute = false,
                                CreateNoWindow = true,
                                RedirectStandardOutput = true
                            }
                        };
                        process.Start();
                        await process.WaitForExitAsync();
                        if (process.ExitCode == 0)
                            removed++;
                    }
                    catch { }
                }

                LblStatus.Text = $"Đã xóa {removed} ứng dụng Bloatware.";
                LblStatus.Foreground = System.Windows.Media.Brushes.Green;
            }
            catch (Exception ex)
            {
                LblStatus.Text = $"Lỗi: {ex.Message}";
                LblStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        private async void BtnBKAV_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndRunTool(
                "https://github.com/minhhaibui1987/Tien.ich.cai.win.dao.hai.do/releases/download/xoago/BKAVREMOVALTOOLPRO.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "BKAVREMOVALTOOLPRO.exe"),
                "BKAV Removal Tool");
        }

        private async void Btn360_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndRunTool(
                "https://github.com/minhhaibui1987/Tien.ich.cai.win.dao.hai.do/releases/download/xoago/go360.zip",
                @"C:\Go 360\go360.zip",
                "360 Uninstaller Tool",
                isZip: true,
                extractPath: @"C:\Go 360",
                exePath: @"C:\Go 360\360 Uninstaller Tool.exe");
        }

        private async void BtnWPS_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndRunTool(
                "https://github.com/minhhaibui1987/Tien.ich.cai.win.dao.hai.do/releases/download/xoago/WPS.OFFICE.UNINSTALLER.TOOL.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Go WPS Office.exe"),
                "WPS Removal Tool");
        }

        private async void BtnEdge_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndRunTool(
                "https://github.com/minhhaibui1987/Tien.ich.cai.win.dao.hai.do/releases/download/xoago/Microsoft.Edge.Remover.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Go MS Edge.exe"),
                "EDGE Removal Tool");
        }

        private async Task DownloadAndRunTool(string url, string savePath, string toolName, 
            bool isZip = false, string extractPath = "", string exePath = "")
        {
            try
            {
                LblStatus.Text = $"Đang tải {toolName}...";
                LblStatus.Foreground = System.Windows.Media.Brushes.Orange;

                // Create directory if needed
                string dir = Path.GetDirectoryName(savePath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(10);
                    var response = await httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    using (var fileStream = new FileStream(savePath, FileMode.Create))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }
                }

                if (isZip)
                {
                    LblStatus.Text = "Đang giải nén...";
                    var extractProcess = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = "powershell.exe",
                            Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"Expand-Archive -LiteralPath '{savePath}' -DestinationPath '{extractPath}' -Force\"",
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    extractProcess.Start();
                    await extractProcess.WaitForExitAsync();

                    if (File.Exists(exePath))
                    {
                        LblStatus.Text = $"Đang chạy {toolName}...";
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = exePath,
                            WorkingDirectory = extractPath,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"Không tìm thấy file: {Path.GetFileName(exePath)}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        LblStatus.Text = "Lỗi: Không tìm thấy file.";
                        LblStatus.Foreground = System.Windows.Media.Brushes.Red;
                        return;
                    }

                    File.Delete(savePath);
                }
                else
                {
                    LblStatus.Text = $"Tải xong. Đang chạy {toolName}...";
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = savePath,
                        UseShellExecute = true
                    });
                }

                LblStatus.Text = $"Hoàn tất. {toolName} đã được khởi động.";
                LblStatus.Foreground = System.Windows.Media.Brushes.Green;
            }
            catch (Exception ex)
            {
                LblStatus.Text = $"Lỗi: {ex.Message}";
                LblStatus.Foreground = System.Windows.Media.Brushes.Red;
                System.Windows.MessageBox.Show($"Lỗi khi tải/chạy {toolName}: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
