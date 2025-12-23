using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace ZibllWindows
{
    public partial class RemoveKeysPage : Wpf.Ui.Controls.FluentWindow
    {
        public RemoveKeysPage()
        {
            InitializeComponent();
        }

        private async void BtnRemoveWindowsKey_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "Bạn có chắc chắn muốn xóa Product Key Windows?\n\n" +
                "Thao tác này sẽ:\n" +
                "- Gỡ bỏ Product Key đã cài đặt\n" +
                "- Xóa key khỏi registry\n" +
                "- Máy tính trở về trạng thái chưa kích hoạt",
                "Xác nhận xóa Key Windows",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    LblStatus.Text = "Đang xóa key Windows...";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Orange;

                    await Task.Run(() =>
                    {
                        string slmgrPath = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows", "System32", "slmgr.vbs");
                        
                        // Remove key
                        var process1 = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = "cscript.exe",
                                Arguments = $"//nologo \"{slmgrPath}\" /upk",
                                UseShellExecute = false,
                                CreateNoWindow = true,
                                RedirectStandardOutput = true
                            }
                        };
                        process1.Start();
                        process1.WaitForExit();

                        // Clear key from registry
                        var process2 = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = "cscript.exe",
                                Arguments = $"//nologo \"{slmgrPath}\" /cpky",
                                UseShellExecute = false,
                                CreateNoWindow = true
                            }
                        };
                        process2.Start();
                        process2.WaitForExit();
                    });

                    LblStatus.Text = "Đã gỡ key Windows thành công!";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Green;
                    System.Windows.MessageBox.Show(
                        "Đã gỡ key Windows.\nMáy tính trở về trạng thái chưa kích hoạt.",
                        "Hoàn tất",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    LblStatus.Text = $"Lỗi: {ex.Message}";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Red;
                    System.Windows.MessageBox.Show($"Lỗi khi gỡ key Windows: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BtnRemoveOfficeKeys_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "Bạn có chắc chắn muốn xóa TẤT CẢ Product Key Office đã cài?\n\n" +
                "Thao tác này không thể hoàn tác.",
                "Xác nhận xóa Key Office",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    LblStatus.Text = "Đang quét các key Office...";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Orange;

                    string osppPath = FindOfficePath();
                    if (string.IsNullOrEmpty(osppPath))
                    {
                        System.Windows.MessageBox.Show("Không tìm thấy OSPP.VBS để xóa key.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Get status to find keys
                    string dstatus = await RunOfficeCommand(osppPath, "/dstatus");
                    
                    // Extract keys (Last 5 characters)
                    var keys = ExtractOfficeKeys(dstatus);

                    if (keys.Count == 0)
                    {
                        LblStatus.Text = "Không có key Office nào để xóa.";
                        LblStatus.Foreground = System.Windows.Media.Brushes.Gray;
                        System.Windows.MessageBox.Show("Không thấy key Office nào để xóa.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    LblStatus.Text = $"Đang xóa {keys.Count} key Office...";
                    int deleted = 0;

                    foreach (var key in keys)
                    {
                        string result2 = await RunOfficeCommand(osppPath, $"/unpkey:{key}");
                        if (!result2.Contains("ERROR") && !result2.Contains("failed"))
                        {
                            deleted++;
                        }
                    }

                    LblStatus.Text = $"Đã xóa {deleted}/{keys.Count} key Office.";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Green;
                    System.Windows.MessageBox.Show($"Đã xóa {deleted} key Office.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    LblStatus.Text = $"Lỗi: {ex.Message}";
                    LblStatus.Foreground = System.Windows.Media.Brushes.Red;
                    System.Windows.MessageBox.Show($"Lỗi khi xóa key Office: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string FindOfficePath()
        {
            string[] possiblePaths = {
                @"C:\Program Files\Microsoft Office\Office16",
                @"C:\Program Files (x86)\Microsoft Office\Office16",
                @"C:\Program Files\Microsoft Office\Office15",
                @"C:\Program Files (x86)\Microsoft Office\Office15",
                @"C:\Program Files\Microsoft Office\Office14",
                @"C:\Program Files (x86)\Microsoft Office\Office14",
                @"C:\Program Files\Microsoft Office\root\Office16",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16"
            };

            foreach (var path in possiblePaths)
            {
                string osppPath = Path.Combine(path, "ospp.vbs");
                if (File.Exists(osppPath))
                    return osppPath;
            }

            // Try system32
            string systemOspp = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows", "System32", "ospp.vbs");
            if (File.Exists(systemOspp))
                return systemOspp;

            return null;
        }

        private async Task<string> RunOfficeCommand(string osppPath, string arguments)
        {
            return await Task.Run(() =>
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "cscript.exe",
                        Arguments = $"//Nologo \"{osppPath}\" {arguments}",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        WorkingDirectory = Path.GetDirectoryName(osppPath)
                    }
                };
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                return output + error;
            });
        }

        private System.Collections.Generic.List<string> ExtractOfficeKeys(string dstatus)
        {
            var keys = new System.Collections.Generic.List<string>();
            var lines = dstatus.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                if (line.Contains("Last 5"))
                {
                    var match = Regex.Match(line, @"Last 5\s*:\s*([A-Z0-9]{5})", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        string key = match.Groups[1].Value.ToUpper();
                        if (!keys.Contains(key))
                            keys.Add(key);
                    }
                }
            }

            return keys;
        }
    }
}
