using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Text;

namespace ZibllWindows
{
    public partial class InstallPage : Wpf.Ui.Controls.FluentWindow
    {
        private string version = "1.2";
        private string appName = "ZibllWindows";
        private string installPath = @"C:\Program Files\ZibllWindows";
        private string desktopShortcutPath;
        private string startMenuShortcutPath;

        public InstallPage()
        {
            InitializeComponent();
            desktopShortcutPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{appName}.lnk");
            startMenuShortcutPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs", $"{appName}.lnk");
            
            Loaded += InstallPage_Loaded;
        }

        private async void InstallPage_Loaded(object sender, RoutedEventArgs e)
        {
            await CheckUpdateAndInstallStatus();
        }

        private async Task CheckUpdateAndInstallStatus()
        {
            try
            {
                StatusText.Text = "Đang kiểm tra cập nhật...";
                ProgressRing.Visibility = Visibility.Visible;
                ProgressBar.Visibility = Visibility.Collapsed;

                // Check if already installed
                bool isInstalled = IsInstalled();
                
                // Check for updates
                bool hasUpdate = false;
                string latestVersion = version;
                
                try
                {
                    hasUpdate = await CheckForUpdate();
                    if (hasUpdate)
                    {
                        latestVersion = await GetLatestVersion();
                    }
                }
                catch
                {
                    // If can't check update, continue anyway
                }

                if (isInstalled)
                {
                    if (hasUpdate)
                    {
                        StatusText.Text = $"Ứng dụng đã được cài đặt.\nPhát hiện phiên bản mới: {latestVersion}\nBạn có muốn cập nhật?";
                        InstallButton.Content = "Cập nhật";
                        InstallButton.Visibility = Visibility.Visible;
                        InstallButton.IsEnabled = true;
                        SkipButton.Visibility = Visibility.Visible;
                        UninstallButton.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        StatusText.Text = "Ứng dụng đã được cài đặt.\nBạn đang sử dụng phiên bản mới nhất.";
                        InstallButton.Visibility = Visibility.Collapsed;
                        SkipButton.Visibility = Visibility.Collapsed;
                        UninstallButton.Visibility = Visibility.Visible;
                        CloseButton.Visibility = Visibility.Visible;
                    }
                }
                else
                {
                    if (hasUpdate)
                    {
                        StatusText.Text = $"Chào mừng đến với ZibllWindows!\nPhát hiện phiên bản mới: {latestVersion}\nBạn có muốn cài đặt?";
                    }
                    else
                    {
                        StatusText.Text = "Chào mừng đến với ZibllWindows!\nBạn có muốn cài đặt ứng dụng?";
                    }
                    
                    InstallButton.Content = "Cài đặt";
                    InstallButton.Visibility = Visibility.Visible;
                    InstallButton.IsEnabled = true;
                    SkipButton.Visibility = Visibility.Visible;
                }

                ProgressRing.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Lỗi: {ex.Message}";
                ProgressRing.Visibility = Visibility.Collapsed;
                CloseButton.Visibility = Visibility.Visible;
            }
        }

        private async Task<bool> CheckForUpdate()
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);
                    string versionUrl = "https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/refs/heads/main/backend/version";
                    string latestVersion = await httpClient.GetStringAsync(versionUrl);
                    latestVersion = latestVersion.Trim();

                    return latestVersion != version;
                }
            }
            catch
            {
                return false;
            }
        }

        private async Task<string> GetLatestVersion()
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);
                    string versionUrl = "https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/refs/heads/main/backend/version";
                    return (await httpClient.GetStringAsync(versionUrl)).Trim();
                }
            }
            catch
            {
                return version;
            }
        }

        private bool IsInstalled()
        {
            return Directory.Exists(installPath) && 
                   File.Exists(Path.Combine(installPath, "ZibllWindows.exe"));
        }

        private async void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InstallButton.IsEnabled = false;
                SkipButton.IsEnabled = false;
                ProgressRing.Visibility = Visibility.Collapsed;
                ProgressBar.Visibility = Visibility.Visible;
                ProgressBar.Value = 0;

                StatusText.Text = "Đang cài đặt...";
                ProgressText.Text = "Đang sao chép files...";
                ProgressBar.Value = 20;

                // Get current executable path
                string currentExe = System.Reflection.Assembly.GetExecutingAssembly().Location;
                if (string.IsNullOrEmpty(currentExe))
                {
                    currentExe = Process.GetCurrentProcess().MainModule.FileName;
                }

                // Create install directory
                Directory.CreateDirectory(installPath);
                ProgressBar.Value = 30;

                // Copy files
                string targetExe = Path.Combine(installPath, "ZibllWindows.exe");
                File.Copy(currentExe, targetExe, true);
                ProgressBar.Value = 50;

                // Copy other files if needed
                string currentDir = Path.GetDirectoryName(currentExe);
                if (!string.IsNullOrEmpty(currentDir))
                {
                    string[] filesToCopy = { "zibllwindows.ico", "app.manifest" };
                    foreach (string file in filesToCopy)
                    {
                        string sourceFile = Path.Combine(currentDir, file);
                        if (File.Exists(sourceFile))
                        {
                            File.Copy(sourceFile, Path.Combine(installPath, file), true);
                        }
                    }
                }
                ProgressBar.Value = 60;

                ProgressText.Text = "Đang tạo shortcut...";
                await Task.Delay(100);
                CreateDesktopShortcut(targetExe);
                CreateStartMenuShortcut(targetExe);
                ProgressBar.Value = 70;

                ProgressText.Text = "Đang thêm vào Registry...";
                await Task.Delay(100);
                AddToRegistry(targetExe);
                ProgressBar.Value = 90;

                ProgressText.Text = "Hoàn tất!";
                ProgressBar.Value = 100;

                StatusText.Text = "Cài đặt thành công!";
                InfoBorder.Visibility = Visibility.Visible;
                InstallInfoText.Text = $"Đã cài đặt vào: {installPath}\n" +
                                      $"Shortcut đã được tạo trên Desktop\n" +
                                      $"Đã thêm vào Programs and Features";

                InstallButton.Visibility = Visibility.Collapsed;
                SkipButton.Visibility = Visibility.Collapsed;
                UninstallButton.Visibility = Visibility.Visible;
                CloseButton.Visibility = Visibility.Visible;

                await Task.Delay(2000);
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Lỗi cài đặt: {ex.Message}";
                ProgressBar.Visibility = Visibility.Collapsed;
                InstallButton.IsEnabled = true;
                SkipButton.IsEnabled = true;
                System.Windows.MessageBox.Show($"Lỗi khi cài đặt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateDesktopShortcut(string targetExe)
        {
            try
            {
                // Create .lnk file using IShellLink COM interface
                Type shellLinkType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellLinkType);
                dynamic shortcut = shell.CreateShortcut(desktopShortcutPath);
                shortcut.TargetPath = targetExe;
                shortcut.WorkingDirectory = installPath;
                shortcut.Description = "ZibllWindows - Công cụ kích hoạt Windows và Office";
                string iconPath = Path.Combine(installPath, "zibllwindows.ico");
                if (File.Exists(iconPath))
                {
                    shortcut.IconLocation = iconPath;
                }
                shortcut.Save();
            }
            catch
            {
                // Fallback: Create .url file
                try
                {
                    string urlPath = desktopShortcutPath.Replace(".lnk", ".url");
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("[InternetShortcut]");
                    sb.AppendLine($"URL=file:///{targetExe.Replace("\\", "/")}");
                    string iconPath = Path.Combine(installPath, "zibllwindows.ico");
                    if (File.Exists(iconPath))
                    {
                        sb.AppendLine($"IconFile={iconPath.Replace("\\", "/")}");
                        sb.AppendLine("IconIndex=0");
                    }
                    File.WriteAllText(urlPath, sb.ToString());
                }
                catch { }
            }
        }

        private void CreateStartMenuShortcut(string targetExe)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(startMenuShortcutPath));
                
                Type shellLinkType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellLinkType);
                dynamic shortcut = shell.CreateShortcut(startMenuShortcutPath);
                shortcut.TargetPath = targetExe;
                shortcut.WorkingDirectory = installPath;
                shortcut.Description = "ZibllWindows - Công cụ kích hoạt Windows và Office";
                string iconPath = Path.Combine(installPath, "zibllwindows.ico");
                if (File.Exists(iconPath))
                {
                    shortcut.IconLocation = iconPath;
                }
                shortcut.Save();
            }
            catch { }
        }

        private void AddToRegistry(string exePath)
        {
            try
            {
                string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ZibllWindows";
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(uninstallKey))
                {
                    if (key != null)
                    {
                        key.SetValue("DisplayName", appName);
                        key.SetValue("DisplayVersion", version);
                        key.SetValue("Publisher", "AlexHerrySeek");
                        key.SetValue("InstallLocation", installPath);
                        key.SetValue("UninstallString", $"\"{exePath}\" /uninstall");
                        key.SetValue("DisplayIcon", Path.Combine(installPath, "zibllwindows.ico"));
                        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                        key.SetValue("EstimatedSize", new FileInfo(exePath).Length / 1024, RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                // May need admin rights
                throw new Exception($"Không thể thêm vào Registry. Cần quyền Administrator: {ex.Message}");
            }
        }

        private async void UninstallButton_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "Bạn có chắc chắn muốn gỡ cài đặt ZibllWindows?\n\n" +
                "Thao tác này sẽ:\n" +
                "- Xóa thư mục cài đặt\n" +
                "- Xóa shortcut trên Desktop và Start Menu\n" +
                "- Xóa registry entries",
                "Xác nhận gỡ cài đặt",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    UninstallButton.IsEnabled = false;
                    ProgressBar.Visibility = Visibility.Visible;
                    ProgressBar.Value = 0;
                    StatusText.Text = "Đang gỡ cài đặt...";

                    ProgressText.Text = "Đang xóa registry...";
                    RemoveFromRegistry();
                    ProgressBar.Value = 30;

                    ProgressText.Text = "Đang xóa shortcut...";
                    await Task.Delay(100);
                    DeleteShortcuts();
                    ProgressBar.Value = 60;

                    ProgressText.Text = "Đang xóa files...";
                    await Task.Delay(100);
                    DeleteInstallDirectory();
                    ProgressBar.Value = 100;

                    StatusText.Text = "Đã gỡ cài đặt thành công!";
                    ProgressText.Text = "";
                    ProgressBar.Visibility = Visibility.Collapsed;
                    CloseButton.Visibility = Visibility.Visible;

                    System.Windows.MessageBox.Show("Đã gỡ cài đặt thành công!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    StatusText.Text = $"Lỗi gỡ cài đặt: {ex.Message}";
                    System.Windows.MessageBox.Show($"Lỗi khi gỡ cài đặt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    UninstallButton.IsEnabled = true;
                }
            }
        }

        private void RemoveFromRegistry()
        {
            try
            {
                string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ZibllWindows";
                Registry.LocalMachine.DeleteSubKey(uninstallKey, false);
            }
            catch { }
        }

        private void DeleteShortcuts()
        {
            try
            {
                if (File.Exists(desktopShortcutPath))
                    File.Delete(desktopShortcutPath);
                if (File.Exists(startMenuShortcutPath))
                    File.Delete(startMenuShortcutPath);
                
                // Also delete .url if exists
                string urlPath = desktopShortcutPath.Replace(".lnk", ".url");
                if (File.Exists(urlPath))
                    File.Delete(urlPath);
            }
            catch { }
        }

        private void DeleteInstallDirectory()
        {
            try
            {
                if (Directory.Exists(installPath))
                {
                    // Try to delete files first
                    string[] files = Directory.GetFiles(installPath);
                    foreach (string file in files)
                    {
                        try
                        {
                            File.SetAttributes(file, FileAttributes.Normal);
                            File.Delete(file);
                        }
                        catch { }
                    }

                    // Delete subdirectories
                    string[] dirs = Directory.GetDirectories(installPath);
                    foreach (string dir in dirs)
                    {
                        try
                        {
                            Directory.Delete(dir, true);
                        }
                        catch { }
                    }

                    // Delete main directory
                    Directory.Delete(installPath, true);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Không thể xóa thư mục cài đặt. Có thể đang được sử dụng: {ex.Message}");
            }
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
