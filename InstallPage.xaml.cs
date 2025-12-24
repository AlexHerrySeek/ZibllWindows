using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Animation;
using Wpf.Ui.Controls;

namespace ZibllWindows
{
    public partial class InstallPage : Wpf.Ui.Controls.FluentWindow
    {
        private string version = "1.3";
        private string appName = "ZibllWindows";
        private string installPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ZibllWindows");

        public InstallPage()
        {
            InitializeComponent();
            Loaded += async (s, e) => await CheckStatusAsync();
        }

        private async Task CheckStatusAsync()
        {
            UpdateUIState("Checking", "Kiểm tra hệ thống", "Đang kết nối tới máy chủ...");

            bool isInstalled = IsInstalled();
            bool hasUpdate = await CheckForUpdateAsync();

            await Task.Delay(1200);

            if (isInstalled)
            {
                if (hasUpdate)
                {
                    UpdateUIState("Update", "Có bản cập nhật mới", "Phiên bản mới mang lại hiệu năng và tính năng tốt hơn.");
                    InstallButton.Content = "Cập nhật ngay";
                    InstallButton.Visibility = Visibility.Visible;
                    SkipButton.Content = "Vào App ngay";
                    SkipButton.Visibility = Visibility.Visible;
                    UninstallButton.Visibility = Visibility.Visible;
                }
                else
                {
                    UpdateUIState("Success", "Đang khởi động", "Bạn đang sử dụng phiên bản mới nhất.");
                    await Task.Delay(1000);
                    OpenMainWindow();
                }
            }
            else
            {
                UpdateUIState("Install", "Chào mừng bạn", "Cài đặt ZibllWindows để bắt đầu trải nghiệm tuyệt vời.");
                InstallButton.Content = "Cài đặt ngay";
                InstallButton.Visibility = Visibility.Visible;
                CloseButton.Visibility = Visibility.Visible;
            }
        }

        private void UpdateUIState(string state, string title, string subTitle)
        {
            StatusTitle.Text = title;
            StatusText.Text = subTitle;
            MainProgressRing.Visibility = Visibility.Collapsed;
            InstallProgressStack.Visibility = Visibility.Collapsed;

            switch (state)
            {
                case "Checking":
                    StateIcon.Symbol = SymbolRegular.ArrowSync24;
                    MainProgressRing.Visibility = Visibility.Visible;
                    break;
                case "Update":
                    StateIcon.Symbol = SymbolRegular.ArrowDownload24;
                    break;
                case "Install":
                    StateIcon.Symbol = SymbolRegular.AddSquare24;
                    break;
                case "Installing":
                    StateIcon.Symbol = SymbolRegular.ArrowCircleDown24;
                    InstallProgressStack.Visibility = Visibility.Visible;
                    break;
                case "Success":
                    StateIcon.Symbol = SymbolRegular.CheckmarkCircle24;
                    StateIcon.Foreground = System.Windows.Media.Brushes.Green;
                    break;
                case "Error":
                    StateIcon.Symbol = SymbolRegular.ErrorCircle24;
                    StateIcon.Foreground = System.Windows.Media.Brushes.Red;
                    break;
            }
        }

        private async void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InstallButton.IsEnabled = false;
                SkipButton.Visibility = Visibility.Collapsed;
                UpdateUIState("Installing", "Đang xử lý", "Vui lòng giữ kết nối và không đóng cửa sổ...");

                UpdateProgress(20, "Đang khởi tạo thư mục...");
                Directory.CreateDirectory(installPath);
                await Task.Delay(500);

                UpdateProgress(60, "Đang sao chép tệp tin hệ thống...");
                string currentExe = Process.GetCurrentProcess().MainModule.FileName;
                string targetExe = Path.Combine(installPath, "ZibllWindows.exe");
                await Task.Run(() => File.Copy(currentExe, targetExe, true));

                UpdateProgress(90, "Đang hoàn tất cấu hình...");
                await Task.Delay(500); // Giả lập xử lý registry

                UpdateProgress(100, "Hoàn tất cài đặt!");
                UpdateUIState("Success", "Thành công!", "Ứng dụng sẽ tự động khởi động.");

                await Task.Delay(1500);
                OpenMainWindow();
            }
            catch (Exception ex)
            {
                UpdateUIState("Error", "Thất bại", "Không thể ghi tệp vào hệ thống.");
                StatusCard.Visibility = Visibility.Visible;
                InstallInfoText.Text = "Lỗi: " + ex.Message + "\nGợi ý: Hãy chạy ứng dụng bằng quyền Administrator.";
                InstallButton.IsEnabled = true;
            }
        }

        private void UpdateProgress(double value, string detail)
        {
            DoubleAnimation anim = new DoubleAnimation(value, TimeSpan.FromMilliseconds(400));
            MainProgressBar.BeginAnimation(System.Windows.Controls.ProgressBar.ValueProperty, anim);
            ProgressDetailText.Text = detail;
        }

        private void OpenMainWindow()
        {
            var main = new MainWindow();
            main.Show();
            this.Close();
        }

        private bool IsInstalled() => File.Exists(Path.Combine(installPath, "ZibllWindows.exe"));

        private async Task<bool> CheckForUpdateAsync()
        {
            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                var res = await client.GetStringAsync("https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/main/backend/version");
                return res.Trim() != version;
            }
            catch { return false; }
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e) => OpenMainWindow();
        private void CloseButton_Click(object sender, RoutedEventArgs e) => this.Close();
        private void UninstallButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Chức năng đang được cập nhật!");
        }
    }
}