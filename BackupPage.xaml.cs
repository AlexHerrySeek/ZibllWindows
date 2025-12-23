using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;

namespace ZibllWindows
{
    public partial class BackupPage : Wpf.Ui.Controls.FluentWindow
    {
        private ObservableCollection<BackupFolderItem> _folders = new ObservableCollection<BackupFolderItem>();
        private bool _isRunning = false;
        private bool _isPaused = false;
        private bool _isCancelled = false;
        private CancellationTokenSource? _cancellationTokenSource;

        public BackupPage()
        {
            InitializeComponent();
            ListViewFolders.ItemsSource = _folders;
            Loaded += BackupPage_Loaded;
        }

        private async void BackupPage_Loaded(object sender, RoutedEventArgs e)
        {
            await ScanFolders();
        }

        private async Task ScanFolders()
        {
            try
            {
                LblStatus.Text = "Đang quét các ổ đĩa...";
                _folders.Clear();

                await Task.Run(() =>
                {
                    var drives = DriveInfo.GetDrives()
                        .Where(d => d.DriveType == DriveType.Fixed && d.IsReady)
                        .Where(d => !d.Name.StartsWith("X:", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    foreach (var drive in drives)
                    {
                        ScanUsersOnDrive(drive.Name);
                    }
                });

                LblStatus.Text = $"Đã tìm thấy {_folders.Count} thư mục.";
                
                if (_folders.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "Không tìm thấy Desktop/Downloads/Documents/Pictures trên các ổ đĩa.",
                        "Thông báo",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LblStatus.Text = $"Lỗi: {ex.Message}";
                System.Windows.MessageBox.Show($"Lỗi khi quét: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ScanUsersOnDrive(string drive)
        {
            try
            {
                string usersPath = Path.Combine(drive, "Users");
                if (!Directory.Exists(usersPath))
                {
                    // Try old Windows XP path
                    usersPath = Path.Combine(drive, "Documents and Settings");
                    if (!Directory.Exists(usersPath))
                        return;
                }

                var userDirs = Directory.GetDirectories(usersPath);
                string[] skipUsers = { "Default", "Default User", "DefaultAppPool", "All Users", "DefaultAccount", "WDAGUtilityAccount" };

                foreach (var userDir in userDirs)
                {
                    string userName = Path.GetFileName(userDir);
                    if (skipUsers.Contains(userName, StringComparer.OrdinalIgnoreCase))
                        continue;

                    CheckAndAddFolder(drive, userDir, userName, "Desktop", Path.Combine(userName, "Desktop"));
                    CheckAndAddFolder(drive, userDir, userName, "Documents", Path.Combine(userName, "Documents"));
                    CheckAndAddFolder(drive, userDir, userName, "Downloads", Path.Combine(userName, "Downloads"));
                    CheckAndAddFolder(drive, userDir, userName, "Pictures", Path.Combine(userName, "Pictures"));

                    // For XP
                    if (usersPath.Contains("Documents and Settings"))
                    {
                        CheckAndAddFolder(drive, Path.Combine(userDir, "My Documents"), userName, "Documents", Path.Combine(userName, "Documents"));
                        CheckAndAddFolder(drive, Path.Combine(userDir, "My Documents", "My Pictures"), userName, "Pictures", Path.Combine(userName, "Pictures"));
                    }
                }
            }
            catch { }
        }

        private void CheckAndAddFolder(string drive, string path, string userName, string folderName, string destRel)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Dispatcher.Invoke(() =>
                    {
                        _folders.Add(new BackupFolderItem
                        {
                            UserName = userName,
                            FolderName = folderName,
                            FullPath = path,
                            Drive = drive,
                            DestinationRelative = destRel,
                            IsSelected = true
                        });
                    });
                }
            }
            catch { }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _folders)
            {
                item.IsSelected = true;
            }
        }

        private void BtnUnselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _folders)
            {
                item.IsSelected = false;
            }
        }

        private async void BtnRescan_Click(object sender, RoutedEventArgs e)
        {
            await ScanFolders();
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtDestination.Text = dialog.SelectedPath;
            }
        }

        private async void BtnBackup_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtDestination.Text))
            {
                System.Windows.MessageBox.Show("Vui lòng chọn thư mục đích!", "Thiếu thư mục đích", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedFolders = _folders.Where(f => f.IsSelected).ToList();
            if (selectedFolders.Count == 0)
            {
                System.Windows.MessageBox.Show("Hãy chọn ít nhất một thư mục để sao lưu!", "Chưa chọn dữ liệu", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string destRoot = TxtDestination.Text;
            if (ChkSubFolder.IsChecked == true)
            {
                destRoot = Path.Combine(destRoot, "Du lieu sao luu", 
                    $"{DateTime.Now:yyyyMMdd}_{DateTime.Now:HHmmss}");
            }
            else
            {
                destRoot = Path.Combine(destRoot, "Du lieu sao luu");
            }

            Directory.CreateDirectory(destRoot);

            _isRunning = true;
            _isPaused = false;
            _isCancelled = false;
            _cancellationTokenSource = new CancellationTokenSource();

            BtnBackup.IsEnabled = false;
            BtnPause.Visibility = Visibility.Visible;
            BtnCancel.Visibility = Visibility.Visible;
            BtnPause.IsEnabled = true;
            BtnCancel.IsEnabled = true;
            ProgressBar.Visibility = Visibility.Visible;
            LblFile.Visibility = Visibility.Visible;
            ScrollLog.Visibility = Visibility.Visible;

            try
            {
                await StartBackup(selectedFolders, destRoot);
            }
            catch (Exception ex)
            {
                AppendLog($"Lỗi: {ex.Message}");
                LblStatus.Text = $"Lỗi: {ex.Message}";
            }
            finally
            {
                _isRunning = false;
                BtnBackup.IsEnabled = true;
                BtnPause.Visibility = Visibility.Collapsed;
                BtnCancel.Visibility = Visibility.Collapsed;
            }
        }

        private async Task StartBackup(List<BackupFolderItem> folders, string destRoot)
        {
            AppendLog($"Bắt đầu sao lưu đến: {destRoot}");

            // Count total files
            long totalFiles = 0;
            LblStatus.Text = "Đang đếm số file...";
            foreach (var folder in folders)
            {
                if (_isCancelled) return;
                await WaitIfPaused();
                totalFiles += CountFiles(folder.FullPath);
            }

            AppendLog($"Tổng số file dự kiến sao lưu: {totalFiles}");

            long copiedFiles = 0;
            foreach (var folder in folders)
            {
                if (_isCancelled)
                {
                    AppendLog("Đã dừng bởi người dùng.");
                    LblStatus.Text = "Đã dừng bởi người dùng.";
                    return;
                }

                await WaitIfPaused();
                string destPath = Path.Combine(destRoot, folder.DestinationRelative);
                LblStatus.Text = $"Đang sao chép: {folder.UserName}\\{folder.FolderName}";
                AppendLog($"Sao chép: {folder.FullPath} -> {destPath}");

                Directory.CreateDirectory(destPath);
                copiedFiles += await CopyDirectoryWithProgress(folder.FullPath, destPath, copiedFiles, totalFiles);
            }

            if (!_isCancelled)
            {
                ProgressBar.Value = 100;
                LblStatus.Text = "Hoàn tất sao lưu!";
                LblFile.Text = "";
                AppendLog($"Hoàn tất. Đích: {destRoot}");
                System.Windows.MessageBox.Show($"Sao lưu đã hoàn tất. Dữ liệu nằm trong:\n{destRoot}", 
                    "Xong", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private long CountFiles(string directory)
        {
            try
            {
                long count = 0;
                if (_isCancelled) return -1;

                var files = Directory.GetFiles(directory);
                count += files.Length;

                foreach (var subDir in Directory.GetDirectories(directory))
                {
                    long subCount = CountFiles(subDir);
                    if (subCount == -1) return -1;
                    count += subCount;
                }

                return count;
            }
            catch
            {
                return 0;
            }
        }

        private async Task<long> CopyDirectoryWithProgress(string source, string dest, long copied, long total)
        {
            long fileCount = 0;
            try
            {
                Directory.CreateDirectory(dest);

                foreach (var file in Directory.GetFiles(source))
                {
                    if (_isCancelled) return fileCount;
                    await WaitIfPaused();

                    string destFile = Path.Combine(dest, Path.GetFileName(file));
                    string displayPath = file.Length > 95 ? file.Substring(0, 45) + "..." + file.Substring(file.Length - 47) : file;
                    LblFile.Text = $"Đang sao lưu file: {displayPath}";
                    AppendLog($"File: {file}");

                    try
                    {
                        File.Copy(file, destFile, true);
                        copied++;
                        fileCount++;
                        if (total > 0)
                        {
                            ProgressBar.Value = (int)((copied / (double)total) * 100);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"Lỗi copy -> {destFile} ({ex.Message})");
                    }
                }

                foreach (var subDir in Directory.GetDirectories(source))
                {
                    if (_isCancelled) return fileCount;
                    await WaitIfPaused();

                    string destSubDir = Path.Combine(dest, Path.GetFileName(subDir));
                    fileCount += await CopyDirectoryWithProgress(subDir, destSubDir, copied, total);
                }
            }
            catch { }

            return fileCount;
        }

        private async Task WaitIfPaused()
        {
            while (_isPaused && !_isCancelled)
            {
                LblStatus.Text = "Tạm dừng...";
                await Task.Delay(50);
            }
        }

        private void BtnPause_Click(object sender, RoutedEventArgs e)
        {
            _isPaused = !_isPaused;
            BtnPause.Content = _isPaused ? "Tiếp tục" : "Tạm dừng";
            AppendLog(_isPaused ? "Tạm dừng bởi người dùng." : "Tiếp tục sao lưu.");
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _isCancelled = true;
            _cancellationTokenSource?.Cancel();
            LblStatus.Text = "Đang dừng...";
            AppendLog("Yêu cầu dừng sao lưu...");
        }

        private void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = $"{DateTime.Now:HH:mm:ss} - ";
                TxtLog.AppendText(timestamp + message + Environment.NewLine);
                TxtLog.ScrollToEnd();
            });
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox cb && cb.DataContext is BackupFolderItem item)
            {
                item.IsSelected = true;
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox cb && cb.DataContext is BackupFolderItem item)
            {
                item.IsSelected = false;
            }
        }
    }

    public class BackupFolderItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string UserName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string Drive { get; set; } = string.Empty;
        public string DestinationRelative { get; set; } = string.Empty;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
