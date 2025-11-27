using System;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace ZibllWindows
{
    public partial class MainWindow
    {
        // System Info Methods (async, non-blocking)
        private async Task LoadSystemInfoAsync()
        {
            try
            {
                ShowBusy("Đang tải thông tin hệ thống...");
                await Task.WhenAll(
                    LoadBasicSystemInfoAsync(),
                    LoadCPUInfoAsync(),
                    LoadMemoryInfoAsync(),
                    LoadGPUInfoAsync(),
                    LoadDiskInfoAsync(),
                    LoadNetworkInfoAsync(),
                    LoadActivationStatusAsync()
                );
            }
            catch (Exception ex)
            {
                _ = ShowInfoDialogAsync("Lỗi", $"Lỗi khi tải thông tin hệ thống: {ex.Message}");
            }
            finally
            {
                HideBusy();
            }
        }

        private async Task LoadBasicSystemInfoAsync()
        {
            var computer = Environment.MachineName;
            var user = Environment.UserName;
            var osName = Environment.OSVersion.ToString();
            var osVer = $"{Environment.OSVersion.Version.Major}.{Environment.OSVersion.Version.Minor}.{Environment.OSVersion.Version.Build}";
            var arch = Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit";
            await Dispatcher.InvokeAsync(() =>
            {
                TxtComputerName.Text = computer;
                TxtUserName.Text = user;
                TxtOSName.Text = osName;
                TxtOSVersion.Text = osVer;
                TxtOSArchitecture.Text = arch;
            });
        }

        private async Task LoadCPUInfoAsync()
        {
            string name = "Không thể lấy thông tin";
            string cores = "N/A";
            string speed = "N/A";
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    name = obj["Name"]?.ToString() ?? name;
                    cores = $"{obj["NumberOfCores"]} nhân / {obj["NumberOfLogicalProcessors"]} luồng";
                    var maxSpeed = obj["MaxClockSpeed"];
                    if (maxSpeed != null)
                    {
                        double speedGHz = Convert.ToDouble(maxSpeed) / 1000;
                        speed = $"{speedGHz:F2} GHz";
                    }
                    break;
                }
            }
            catch { }

            await Dispatcher.InvokeAsync(() =>
            {
                TxtCPUName.Text = name;
                TxtCPUCores.Text = cores;
                TxtCPUSpeed.Text = speed;
            });
        }

        private async Task LoadMemoryInfoAsync()
        {
            string total = "Không thể lấy thông tin";
            string available = "N/A";
            string used = "N/A";
            try
            {
                double totalRamGb = 0;
                using (var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        totalRamGb = Convert.ToDouble(obj["TotalPhysicalMemory"]) / (1024 * 1024 * 1024);
                        total = $"{totalRamGb:F2} GB";
                        break;
                    }
                }

                using (var searcher = new ManagementObjectSearcher("SELECT FreePhysicalMemory, TotalVisibleMemorySize FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var freeGb = Convert.ToDouble(obj["FreePhysicalMemory"]) / (1024 * 1024);
                        var totalVisibleGb = Convert.ToDouble(obj["TotalVisibleMemorySize"]) / (1024 * 1024);
                        var usedGb = totalVisibleGb - freeGb;
                        available = $"{freeGb:F2} GB";
                        used = $"{usedGb:F2} GB ({(usedGb / totalVisibleGb * 100):F1}%)";
                        break;
                    }
                }
            }
            catch { }

            await Dispatcher.InvokeAsync(() =>
            {
                TxtTotalRAM.Text = total;
                TxtAvailableRAM.Text = available;
                TxtUsedRAM.Text = used;
            });
        }

        private async Task LoadGPUInfoAsync()
        {
            string name = "Không thể lấy thông tin";
            string memory = "N/A";
            string driver = "N/A";

            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT Name, AdapterRAM, DriverVersion, AdapterCompatibility, PNPDeviceID FROM Win32_VideoController");
                var results = searcher.Get().Cast<ManagementObject>().ToList();
                ManagementObject? best = null;

                // Prefer discrete GPUs (NVIDIA/AMD/Radeon/GeForce)
                best = results.FirstOrDefault(o =>
                {
                    var n = (o["Name"]?.ToString() ?? "");
                    var comp = (o["AdapterCompatibility"]?.ToString() ?? "");
                    return n.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase)
                           || n.Contains("GeForce", StringComparison.OrdinalIgnoreCase)
                           || n.Contains("AMD", StringComparison.OrdinalIgnoreCase)
                           || n.Contains("Radeon", StringComparison.OrdinalIgnoreCase)
                           || comp.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase)
                           || comp.Contains("Advanced Micro Devices", StringComparison.OrdinalIgnoreCase)
                           || comp.Contains("AMD", StringComparison.OrdinalIgnoreCase);
                });

                // If none matched, pick the one with largest AdapterRAM
                if (best == null && results.Count > 0)
                {
                    best = results.OrderByDescending(o =>
                    {
                        try { return Convert.ToUInt64(o["AdapterRAM"] ?? 0); }
                        catch { return 0UL; }
                    }).First();
                }

                if (best != null)
                {
                    name = best["Name"]?.ToString() ?? name;
                    var adapterRAM = best["AdapterRAM"];
                    if (adapterRAM != null)
                    {
                        double ramGB = Convert.ToDouble(adapterRAM) / (1024 * 1024 * 1024);
                        memory = $"{ramGB:F2} GB";
                    }
                    driver = best["DriverVersion"]?.ToString() ?? driver;
                }
            }
            catch { }

            await Dispatcher.InvokeAsync(() =>
            {
                TxtGPUName.Text = name;
                TxtGPUMemory.Text = memory;
                TxtGPUDriver.Text = driver;
            });
        }

        private async Task LoadDiskInfoAsync()
        {
            var drives = new List<(string Name, string Type, string Label, string Total, string Free, string Used, bool IsSystem)>();
            try
            {
                var systemDrive = System.IO.Path.GetPathRoot(Environment.SystemDirectory)?.TrimEnd('\\');
                var allDrives = System.IO.DriveInfo.GetDrives().Where(d => d.IsReady).ToList();
                
                foreach (var d in allDrives)
                {
                    try
                    {
                        var driveName = d.Name.TrimEnd('\\');
                        var driveType = d.DriveType switch
                        {
                            System.IO.DriveType.Fixed => "Cứng",
                            System.IO.DriveType.Removable => "USB/Tháo lắp",
                            System.IO.DriveType.Network => "Mạng",
                            System.IO.DriveType.CDRom => "CD/DVD",
                            System.IO.DriveType.Ram => "RAM",
                            _ => "Khác"
                        };
                        var label = string.IsNullOrWhiteSpace(d.VolumeLabel) ? "(Không tên)" : d.VolumeLabel;
                        var totalGB = d.TotalSize / (1024.0 * 1024 * 1024);
                        var freeGB = d.AvailableFreeSpace / (1024.0 * 1024 * 1024);
                        var usedGB = totalGB - freeGB;
                        var usedPercent = (usedGB / totalGB * 100);
                        var isSystem = driveName.Equals(systemDrive, StringComparison.OrdinalIgnoreCase);
                        
                        drives.Add((driveName, driveType, label, $"{totalGB:F2} GB", $"{freeGB:F2} GB", $"{usedGB:F2} GB ({usedPercent:F1}%)", isSystem));
                    }
                    catch { }
                }
            }
            catch { }

            await Dispatcher.InvokeAsync(() =>
            {
                DiskInfoPanel.Children.Clear();
                if (!drives.Any())
                {
                    var noData = new System.Windows.Controls.TextBlock
                    {
                        Text = "Không thể lấy thông tin ổ đĩa",
                        FontStyle = System.Windows.FontStyles.Italic,
                        Foreground = System.Windows.Media.Brushes.Gray
                    };
                    DiskInfoPanel.Children.Add(noData);
                    return;
                }

                foreach (var (name, type, label, total, free, used, isSystem) in drives)
                {
                    var separator = new System.Windows.Controls.Border
                    {
                        Height = 1,
                        Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(30, 128, 128, 128)),
                        Margin = new System.Windows.Thickness(0, 10, 0, 10)
                    };
                    if (DiskInfoPanel.Children.Count > 0)
                        DiskInfoPanel.Children.Add(separator);

                    var header = new System.Windows.Controls.TextBlock
                    {
                        Text = $" {name} - {label} [{type}]" + (isSystem ? "  Hệ thống" : ""),
                        FontWeight = System.Windows.FontWeights.Bold,
                        FontSize = 14,
                        Margin = new System.Windows.Thickness(0, 0, 0, 8)
                    };
                    DiskInfoPanel.Children.Add(header);

                    var grid = new System.Windows.Controls.Grid
                    {
                        Margin = new System.Windows.Thickness(10, 0, 0, 0)
                    };
                    grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(150) });
                    grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
                    grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition());
                    grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition());
                    grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition());

                    var lbl1 = new System.Windows.Controls.TextBlock { Text = "Tổng dung lượng:", FontWeight = System.Windows.FontWeights.SemiBold };
                    System.Windows.Controls.Grid.SetRow(lbl1, 0);
                    System.Windows.Controls.Grid.SetColumn(lbl1, 0);
                    grid.Children.Add(lbl1);

                    var val1 = new System.Windows.Controls.TextBlock { Text = total, TextWrapping = System.Windows.TextWrapping.Wrap };
                    System.Windows.Controls.Grid.SetRow(val1, 0);
                    System.Windows.Controls.Grid.SetColumn(val1, 1);
                    grid.Children.Add(val1);

                    var lbl2 = new System.Windows.Controls.TextBlock { Text = "Còn trống:", FontWeight = System.Windows.FontWeights.SemiBold, Margin = new System.Windows.Thickness(0, 4, 0, 0) };
                    System.Windows.Controls.Grid.SetRow(lbl2, 1);
                    System.Windows.Controls.Grid.SetColumn(lbl2, 0);
                    grid.Children.Add(lbl2);

                    var val2 = new System.Windows.Controls.TextBlock { Text = free, TextWrapping = System.Windows.TextWrapping.Wrap, Margin = new System.Windows.Thickness(0, 4, 0, 0) };
                    System.Windows.Controls.Grid.SetRow(val2, 1);
                    System.Windows.Controls.Grid.SetColumn(val2, 1);
                    grid.Children.Add(val2);

                    var lbl3 = new System.Windows.Controls.TextBlock { Text = "Đã sử dụng:", FontWeight = System.Windows.FontWeights.SemiBold, Margin = new System.Windows.Thickness(0, 4, 0, 0) };
                    System.Windows.Controls.Grid.SetRow(lbl3, 2);
                    System.Windows.Controls.Grid.SetColumn(lbl3, 0);
                    grid.Children.Add(lbl3);

                    var val3 = new System.Windows.Controls.TextBlock { Text = used, TextWrapping = System.Windows.TextWrapping.Wrap, Margin = new System.Windows.Thickness(0, 4, 0, 0) };
                    System.Windows.Controls.Grid.SetRow(val3, 2);
                    System.Windows.Controls.Grid.SetColumn(val3, 1);
                    grid.Children.Add(val3);

                    DiskInfoPanel.Children.Add(grid);
                }
            });
        }

        private async Task LoadNetworkInfoAsync()
        {
            string ip = "Không thể lấy thông tin";
            string mac = "N/A";
            string name = "N/A";
            try
            {
                var networkInterfaces = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(ni => ni.OperationalStatus == OperationalStatus.Up &&
                                 ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    .ToList();

                if (networkInterfaces.Any())
                {
                    var activeInterface = networkInterfaces.First();
                    var ipProperties = activeInterface.GetIPProperties();
                    var ipv4 = ipProperties.UnicastAddresses
                        .FirstOrDefault(addr => addr.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                    ip = ipv4?.Address.ToString() ?? "N/A";
                    mac = string.Join(":", activeInterface.GetPhysicalAddress().GetAddressBytes().Select(b => b.ToString("X2")));
                    name = activeInterface.Name;
                }
                else
                {
                    ip = "Không có kết nối";
                }
            }
            catch { }

            await Dispatcher.InvokeAsync(() =>
            {
                TxtLocalIP.Text = ip;
                TxtMACAddress.Text = mac;
                TxtNetworkName.Text = name;
            });
        }

        private record ActivationStatus(
            bool WindowsActivated,
            string WindowsChannel,
            string WindowsExpiry,
            string WindowsEdition,
            string OfficeStatus,
            string OfficeDetail);

        private async Task LoadActivationStatusAsync()
        {
            ActivationStatus status;
            try
            {
                var windowsXpr = await RunCommand("slmgr.vbs", "/xpr");
                var windowsDli = await RunCommand("slmgr.vbs", "/dli");
                bool winActivated = windowsXpr.IndexOf("permanent", StringComparison.OrdinalIgnoreCase) >= 0
                                     || windowsXpr.IndexOf("vĩnh viễn", StringComparison.OrdinalIgnoreCase) >= 0
                                     || windowsDli.IndexOf("Licensed", StringComparison.OrdinalIgnoreCase) >= 0;

                string edition = ExtractLineContains(windowsDli, new[] { "Name:", "Windows", "Professional", "Enterprise" });
                string channel = ExtractLineContains(windowsDli, new[] { "Description:", "KMS", "RETAIL", "VOLUME" });
                string expiry = ParseKmsExpiry(windowsXpr);

                var osppPath = GetOsppVbsPath();
                string officeStatus = "Chưa cài";
                string officeDetail = string.Empty;
                if (!string.IsNullOrEmpty(osppPath) && File.Exists(osppPath))
                {
                    var officeInfo = await RunCommand("cscript.exe", $"\"{osppPath}\" /dstatus");
                    officeDetail = ShortenOneLine(officeInfo, 220);
                    if (officeInfo.IndexOf("LICENSE STATUS: ---LICENSED---", StringComparison.OrdinalIgnoreCase) >= 0)
                        officeStatus = "Đã kích hoạt";
                    else if (officeInfo.IndexOf("---NOTIFICATIONS---", StringComparison.OrdinalIgnoreCase) >= 0 || officeInfo.IndexOf("UNLICENSED", StringComparison.OrdinalIgnoreCase) >= 0)
                        officeStatus = "Chưa kích hoạt";
                    else
                        officeStatus = "Đã cài";
                }

                status = new ActivationStatus(winActivated, channel, expiry, edition, officeStatus, officeDetail);
            }
            catch
            {
                status = new ActivationStatus(false, "Không lấy được", "—", "Không lấy được", "Không lấy được", "");
            }

            await Dispatcher.InvokeAsync(() =>
            {
                TxtWinActivation.Text = status.WindowsActivated ? "ĐÃ KÍCH HOẠT" : "CHƯA KÍCH HOẠT";
                TxtWinEdition.Text = string.IsNullOrWhiteSpace(status.WindowsEdition) ? "—" : status.WindowsEdition;
                TxtWinExpiry.Text = status.WindowsActivated && !string.IsNullOrWhiteSpace(status.WindowsExpiry) ? status.WindowsExpiry : (status.WindowsActivated ? "Vĩnh viễn" : "—");
                TxtOfficeStatus.Text = status.OfficeStatus;
                TxtOfficeDetail.Text = string.IsNullOrWhiteSpace(status.OfficeDetail) ? "—" : status.OfficeDetail;
            });
        }

        // Duplicate methods removed to avoid conflicts.
        // Using existing implementations of ExtractLineContains and GetOsppVbsPath in MainWindow.xaml.cs.

        private static string ParseKmsExpiry(string xpr)
        {
            if (string.IsNullOrWhiteSpace(xpr)) return string.Empty;
            try
            {
                var parts = xpr.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var line = parts.FirstOrDefault(p => p.IndexOf("expire", StringComparison.OrdinalIgnoreCase) >= 0);
                if (line != null)
                {
                    var tokens = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    foreach (var t in tokens)
                    {
                        if (DateTime.TryParse(t, out var dt))
                        {
                            var remaining = dt > DateTime.Now ? (dt - DateTime.Now).TotalDays : 0;
                            return dt.ToString("dd/MM/yyyy") + (remaining > 0 ? $" (còn {(int)remaining} ngày)" : "");
                        }
                    }
                }
            }
            catch { }
            return string.Empty;
        }

        private static string ShortenOneLine(string input, int max)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var one = input.Replace('\r', ' ').Replace('\n', ' ').Trim();
            return one.Length <= max ? one : one.Substring(0, max) + "...";
        }

        private async void BtnRefreshSystemInfo_Click(object sender, RoutedEventArgs e)
        {
            await LoadSystemInfoAsync();
            await ShowInfoDialogAsync("Thông báo", "Đã làm mới thông tin hệ thống!");
        }

        // Settings Handlers
    }
}
