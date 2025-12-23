using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Wpf.Ui.Controls;

namespace ZibllWindows
{
    public partial class MainWindow : Wpf.Ui.Controls.FluentWindow
    {
        private AppSettings settings;
        private System.Windows.Media.ImageBrush backgroundBrush = new System.Windows.Media.ImageBrush();
        private string version = "1.2";

        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;

            // Load settings
            settings = AppSettings.Load();
            ApplySettings();
            
            CheckAdminPrivileges();
            
            // Navigate to home page by default
            NavigateToPage("home");
            
            // Show intro
            Loaded += async (s, e) => 
            {
                await ShowIntro();
                // Check for updates after intro
                await CheckAndUpdateIfNeeded();
            };
            
            // Save window size on resize/close
            SizeChanged += MainWindow_SizeChanged;
            StateChanged += MainWindow_StateChanged;
            Closing += MainWindow_Closing;
        }

        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Normal && settings != null)
            {
                settings.WindowWidth = Width;
                settings.WindowHeight = Height;
            }
        }

        private void MainWindow_StateChanged(object? sender, EventArgs e)
        {
            if (settings != null)
            {
                settings.WindowState = (int)WindowState;
            }
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                settings?.Save();
            }
            catch
            {
                // bỏ qua lỗi save
            }
            e.Cancel = true;
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private async Task ShowIntro()
        {
            try
            {
                // Wait for intro animations to play (1.5 seconds)
                await Task.Delay(1500);
                
                // Show intro for 1 second
                IntroStatusText.Text = "Đang khởi động...";
                await Task.Delay(1000);
                
                // Fade out intro and show welcome message
                await FadeOutIntro();
                ShowSnackbar("Sẵn sàng", "Ứng dụng đã sẵn sàng để kích hoạt Windows và Office", ControlAppearance.Success, 4000);
            }
            catch (Exception ex)
            {
                IntroStatusText.Text = "Lỗi khi khởi động";
                await Task.Delay(1500);
                await FadeOutIntro();
            }
        }

        private async Task CheckAndUpdateIfNeeded()
        {
            try
            {
                string installPath = @"C:\Program Files\ZibllWindows";
                bool isInstalled = Directory.Exists(installPath) && 
                                   File.Exists(Path.Combine(installPath, "ZibllWindows.exe"));

                if (!isInstalled)
                    return; // Skip update check if not installed

                SetStatus("Đang kiểm tra cập nhật...", InfoBarSeverity.Informational);
                
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);
                    string versionUrl = "https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/refs/heads/main/backend/version";
                    string latestVersion = await httpClient.GetStringAsync(versionUrl);
                    latestVersion = latestVersion.Trim();

                    if (latestVersion != version)
                    {
                        SetStatus($"Phát hiện phiên bản mới: {latestVersion}", InfoBarSeverity.Success);
                        
                        var result = await ShowConfirmDialogAsync("Cập nhật có sẵn",
                            $"Phiên bản hiện tại: {version}\nPhiên bản mới: {latestVersion}\n\nBạn có muốn tải về và cập nhật?");

                        if (result)
                        {
                            await DownloadAndUpdate(latestVersion);
                        }
                    }
                    else
                    {
                        SetStatus("Bạn đang sử dụng phiên bản mới nhất", InfoBarSeverity.Success);
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Không thể kiểm tra cập nhật: {ex.Message}", InfoBarSeverity.Warning);
            }
        }

        private async Task DownloadAndUpdate(string latestVersion)
        {
            try
            {
                SetStatus("Đang tải bản cập nhật...", InfoBarSeverity.Informational);
                
                string downloadUrl = $"https://github.com/AlexHerrySeek/ZibllWindows/releases/download/v{latestVersion}/ZibllWindows.exe";
                string tempPath = Path.Combine(Path.GetTempPath(), "ZibllWindows_Update.exe");

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(5);
                    var response = await httpClient.GetAsync(downloadUrl);
                    response.EnsureSuccessStatusCode();

                    using (var fileStream = new FileStream(tempPath, FileMode.Create))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }
                }

                // Run updater
                var updaterProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = tempPath,
                        UseShellExecute = true
                    }
                };
                updaterProcess.Start();

                // Close current app
                System.Windows.Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi khi tải cập nhật: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Không thể tải cập nhật:\n\n{ex.Message}");
            }
        }

        private async Task FadeOutIntro()
        {
            var fadeOut = new System.Windows.Media.Animation.DoubleAnimation
            {
                From = 1.0,
                To = 0.0,
                Duration = TimeSpan.FromMilliseconds(500),
                EasingFunction = new System.Windows.Media.Animation.CubicEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseIn }
            };

            fadeOut.Completed += (s, e) => IntroOverlay.Visibility = Visibility.Collapsed;
            IntroOverlay.BeginAnimation(UIElement.OpacityProperty, fadeOut);
            await Task.Delay(500);
        }

        private void MainNavigation_PaneOpenedOrClosed(NavigationView sender, RoutedEventArgs args)
        {
            // Điều chỉnh Margin của ScrollViewer khi menu mở/đóng
            if (sender.IsPaneOpen)
            {
                // Menu mở - margin = OpenPaneLength
                if (MainContentGrid?.Parent is ScrollViewer scrollViewer)
                {
                    scrollViewer.Margin = new Thickness(220, 0, 0, 0);
                }
            }
            else
            {
                // Menu đóng - margin = CompactPaneLength
                if (MainContentGrid?.Parent is ScrollViewer scrollViewer)
                {
                    scrollViewer.Margin = new Thickness(60, 0, 0, 0);
                }
            }
        }

        private void ShowBusy(string message = "Đang xử lý...")
        {
            Dispatcher.Invoke(() =>
            {
                if (BusyText != null) BusyText.Text = message;
                if (BusyOverlay != null) BusyOverlay.Visibility = Visibility.Visible;
            });
        }

        private void HideBusy()
        {
            Dispatcher.Invoke(() =>
            {
                if (BusyOverlay != null) BusyOverlay.Visibility = Visibility.Collapsed;
            });
        }

        private void ShowSnackbar(string title, string message, ControlAppearance appearance = ControlAppearance.Secondary, int timeout = 3000)
        {
            Dispatcher.Invoke(() =>
            {
                try
                {
                    var snackbar = new Snackbar(SnackbarPresenter)
                    {
                        Title = title,
                        Content = message,
                        Appearance = appearance,
                        Timeout = TimeSpan.FromMilliseconds(timeout)
                    };
                    snackbar.Show();
                }
                catch { /* Ignore snackbar errors */ }
            });
        }

        private async void MainNavigation_SelectionChanged(NavigationView sender, RoutedEventArgs args)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"SelectionChanged fired. SelectedItem: {sender.SelectedItem}");
                
                if (sender.SelectedItem is NavigationViewItem item && item.Tag != null)
                {
                    var pageTag = item.Tag.ToString()?.ToLower();
                    System.Diagnostics.Debug.WriteLine($"Navigating to: {pageTag}");
                    NavigateToPage(pageTag ?? "home");
                    
                    // Load system info when navigating to that page
                    if (pageTag == "systeminfo")
                    {
                        await LoadSystemInfoAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SelectionChanged: {ex.Message}");
            }
        }

        private async void NavItem_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            e.Handled = true; // Prevent NavigationView from handling this
            
            if (sender is NavigationViewItem item && item.Tag != null)
            {
                var pageTag = item.Tag.ToString()?.ToLower();
                System.Diagnostics.Debug.WriteLine($"NavItem clicked: {pageTag}");
                
                if (!string.IsNullOrEmpty(pageTag))
                {
                    NavigateToPage(pageTag);
                    
                    // Load system info when navigating to that page
                    if (pageTag == "systeminfo")
                    {
                        await LoadSystemInfoAsync();
                    }
                }
            }
        }

        private void NavigateToWindows_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage("activate");
        }

        private void NavigateToOffice_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage("activate");
        }

        private void BtnOpenBackup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var backupPage = new BackupPage
                {
                    Owner = this
                };
                backupPage.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở công cụ sao lưu: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void NavigateToPage(string pageTag)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToPage called with: '{pageTag}'");
            
            Dispatcher.Invoke(() =>
            {
                // Hide all pages
                HomePage.Visibility = Visibility.Collapsed;
                if (WindowsPage != null) WindowsPage.Visibility = Visibility.Collapsed;
                if (OfficePage != null) OfficePage.Visibility = Visibility.Collapsed;
                if (ActivatePage != null) ActivatePage.Visibility = Visibility.Collapsed;
                if (BackupPage != null) BackupPage.Visibility = Visibility.Collapsed;
                if (DevelopmentPage != null) DevelopmentPage.Visibility = Visibility.Collapsed;
                ToolsPage.Visibility = Visibility.Collapsed;
                GhostToolboxPage.Visibility = Visibility.Collapsed;
                SystemInfoPage.Visibility = Visibility.Collapsed;
                SettingsPage.Visibility = Visibility.Collapsed;
                HelpPage.Visibility = Visibility.Collapsed;

                // Show selected page
                switch (pageTag?.ToLower())
                {
                    case "home":
                        HomePage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing HomePage");
                        break;
                    case "activate":
                        if (ActivatePage != null) ActivatePage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing ActivatePage");
                        break;
                    case "windows":
                        if (WindowsPage != null) WindowsPage.Visibility = Visibility.Visible;
                        else if (ActivatePage != null) ActivatePage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing WindowsPage/ActivatePage");
                        break;
                    case "office":
                        if (OfficePage != null) OfficePage.Visibility = Visibility.Visible;
                        else if (ActivatePage != null) ActivatePage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing OfficePage/ActivatePage");
                        break;
                    case "backup":
                        if (BackupPage != null) BackupPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing BackupPage");
                        break;
                    case "development":
                        if (DevelopmentPage != null) DevelopmentPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing DevelopmentPage");
                        break;
                    case "tools":
                        ToolsPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing ToolsPage");
                        break;
                    case "ghosttoolbox":
                        GhostToolboxPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing GhostToolboxPage");
                        break;
                    case "systeminfo":
                        SystemInfoPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing SystemInfoPage");
                        break;
                    case "settings":
                        SettingsPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing SettingsPage");
                        break;
                    case "help":
                        HelpPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing HelpPage");
                        break;
                    default:
                        System.Diagnostics.Debug.WriteLine($"No match found for pageTag: '{pageTag}'");
                        HomePage.Visibility = Visibility.Visible; // Default to home
                        break;
                }
            });
        }

        private void ApplySettings()
        {
            if (settings == null) return;
            
            // Apply window size and state
            if (settings.WindowWidth > 0 && settings.WindowHeight > 0)
            {
                Width = settings.WindowWidth;
                Height = settings.WindowHeight;
            }
            if (settings.WindowState >= 0 && settings.WindowState <= 2)
            {
                WindowState = (WindowState)settings.WindowState;
            }
            
            // Apply opacity
            this.Opacity = settings.Opacity;
            if (OpacitySlider != null) OpacitySlider.Value = settings.Opacity;
            if (OpacityValue != null) OpacityValue.Text = $"{(int)(settings.Opacity * 100)}%";

            // Apply font family
            if (MainContentGrid != null)
            {
                var fontFamily = new System.Windows.Media.FontFamily(settings.FontFamily);
                System.Windows.Documents.TextElement.SetFontFamily(MainContentGrid, fontFamily);
            }
            if (FontFamilyComboBox != null && FontFamilyComboBox.Items.Count > 0)
            {
                FontFamilyComboBox.SelectedItem = FontFamilyComboBox.Items.Cast<System.Windows.Controls.ComboBoxItem>()
                    .FirstOrDefault(item => item.Content?.ToString() == settings.FontFamily);
            }

            // Apply font size
            if (MainContentGrid != null)
            {
                System.Windows.Documents.TextElement.SetFontSize(MainContentGrid, settings.FontSize);
            }
            if (FontSizeSlider != null) FontSizeSlider.Value = settings.FontSize;
            if (FontSizeValue != null) FontSizeValue.Text = $"{(int)settings.FontSize}px";

            // Apply theme
            if (ThemeComboBox != null && ThemeComboBox.Items.Count > 0)
            {
                ThemeComboBox.SelectedItem = ThemeComboBox.Items.Cast<System.Windows.Controls.ComboBoxItem>()
                    .FirstOrDefault(item => item.Content?.ToString() == settings.Theme);
            }

            // Apply accent color
            if (AccentColorComboBox != null && AccentColorComboBox.Items.Count > 0)
            {
                AccentColorComboBox.SelectedItem = AccentColorComboBox.Items.Cast<System.Windows.Controls.ComboBoxItem>()
                    .FirstOrDefault(item => item.Content?.ToString() == settings.AccentColor);
            }
            
            // Apply saved accent color
            try
            {
                System.Windows.Media.Color accentColor;
                switch (settings.AccentColor)
                {
                    case "Xanh dương":
                        accentColor = System.Windows.Media.Color.FromRgb(59, 130, 246);
                        break;
                    case "Tím":
                        accentColor = System.Windows.Media.Color.FromRgb(168, 85, 247);
                        break;
                    case "Hồng":
                        accentColor = System.Windows.Media.Color.FromRgb(236, 72, 153);
                        break;
                    case "Đỏ":
                        accentColor = System.Windows.Media.Color.FromRgb(239, 68, 68);
                        break;
                    case "Cam":
                        accentColor = System.Windows.Media.Color.FromRgb(249, 115, 22);
                        break;
                    case "Xanh lá":
                        accentColor = System.Windows.Media.Color.FromRgb(34, 197, 94);
                        break;
                    default:
                        accentColor = System.Windows.Media.Color.FromRgb(0, 120, 212);
                        break;
                }
                Wpf.Ui.Appearance.ApplicationAccentColorManager.Apply(accentColor);
            }
            catch { /* Ignore accent color errors */ }

            // Apply background image
            if (BackgroundImageCheckBox != null)
            {
                BackgroundImageCheckBox.IsChecked = settings.UseBackgroundImage;
            }
            if (BackgroundOpacityPanel != null)
            {
                BackgroundOpacityPanel.Visibility = settings.UseBackgroundImage ? Visibility.Visible : Visibility.Collapsed;
            }
            
            // Apply background theme
            if (BackgroundThemeComboBox != null && BackgroundThemeComboBox.Items.Count > 0)
            {
                var content = settings.BackgroundTheme;
                foreach (System.Windows.Controls.ComboBoxItem item in BackgroundThemeComboBox.Items)
                {
                    var itemContent = item.Content.ToString() ?? "";
                    if (itemContent.Contains(content) || (content == "Không" && itemContent == "Không"))
                    {
                        BackgroundThemeComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
            
            // Apply animations setting
            if (AnimationCheckBox != null)
            {
                AnimationCheckBox.IsChecked = settings.UseAnimations;
            }
            
            // Apply border radius
            if (BorderRadiusSlider != null)
            {
                BorderRadiusSlider.Value = settings.BorderRadius;
            }
            if (BorderRadiusValue != null)
            {
                BorderRadiusValue.Text = $"{(int)settings.BorderRadius}px";
            }
            
            // Apply compact mode
            if (CompactModeCheckBox != null)
            {
                CompactModeCheckBox.IsChecked = settings.CompactMode;
            }
            var navView = this.FindName("MainNavigation") as Wpf.Ui.Controls.NavigationView;
            if (navView != null && settings.CompactMode)
            {
                navView.IsPaneOpen = false;
            }
            
            // Apply backdrop type
            if (BackdropComboBox != null && BackdropComboBox.Items.Count > 0)
            {
                BackdropComboBox.SelectedItem = BackdropComboBox.Items.Cast<System.Windows.Controls.ComboBoxItem>()
                    .FirstOrDefault(item => item.Content?.ToString() == settings.BackdropType);
            }
            
            if (settings.UseBackgroundImage && !string.IsNullOrEmpty(settings.BackgroundImagePath) && System.IO.File.Exists(settings.BackgroundImagePath))
            {
                try
                {
                    backgroundBrush = new System.Windows.Media.ImageBrush(new System.Windows.Media.Imaging.BitmapImage(new Uri(settings.BackgroundImagePath)));
                    backgroundBrush.Opacity = settings.BackgroundOpacity;
                    backgroundBrush.Stretch = System.Windows.Media.Stretch.UniformToFill;
                    if (RootGrid != null) RootGrid.Background = backgroundBrush;
                    if (BackgroundOpacitySlider != null) BackgroundOpacitySlider.Value = settings.BackgroundOpacity;
                    if (BackgroundOpacityValue != null) BackgroundOpacityValue.Text = $"{(int)(settings.BackgroundOpacity * 100)}%";
                }
                catch
                {
                    // Failed to load image, reset to transparent
                    if (RootGrid != null) RootGrid.Background = System.Windows.Media.Brushes.Transparent;
                }
            }
            else
            {
                // No background image, use transparent to show Mica
                if (RootGrid != null) RootGrid.Background = System.Windows.Media.Brushes.Transparent;
            }
        }

        private void OpacitySlider_ValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
        {
                if (settings == null || OpacityValue == null) return;

                // When using Mica/Acrylic backdrops, changing Window.Opacity can cause white overlay.
                // Apply opacity to content instead and keep window fully opaque.
                if (this.WindowBackdropType != Wpf.Ui.Controls.WindowBackdropType.None)
                {
                    this.Opacity = 1.0;
                    if (MainContentGrid != null)
                        MainContentGrid.Opacity = e.NewValue;
                }
                else
                {
                    this.Opacity = e.NewValue;
                    if (MainContentGrid != null)
                        MainContentGrid.Opacity = 1.0;
                }

                OpacityValue.Text = $"{(int)(e.NewValue * 100)}%";
                settings.Opacity = e.NewValue;
                settings.Save();
        }

        private void FontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (settings == null || FontFamilyComboBox?.SelectedItem == null || MainContentGrid == null) return;

            var selectedItem = FontFamilyComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (selectedItem?.Content == null) return;
            var fontName = selectedItem.Content.ToString() ?? "Segoe UI";
            
            System.Windows.Documents.TextElement.SetFontFamily(MainContentGrid, new System.Windows.Media.FontFamily(fontName));
            settings.FontFamily = fontName;
            settings.Save();
        }

        private void FontSizeSlider_ValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
        {
            if (settings == null || FontSizeValue == null) return;
            
            System.Windows.Documents.TextElement.SetFontSize(MainContentGrid, e.NewValue);
            FontSizeValue.Text = $"{(int)e.NewValue}px";
            settings.FontSize = e.NewValue;
            settings.Save();
        }

        private void BackgroundImageCheckBox_Changed(object sender, System.Windows.RoutedEventArgs e)
        {
            if (settings == null || RootGrid == null) return;

            settings.UseBackgroundImage = BackgroundImageCheckBox.IsChecked == true;
            if (BackgroundOpacityPanel != null)
                BackgroundOpacityPanel.Visibility = settings.UseBackgroundImage ? Visibility.Visible : Visibility.Collapsed;

            if (!settings.UseBackgroundImage)
            {
                RootGrid.Background = System.Windows.Media.Brushes.Transparent;
            }
            else if (!string.IsNullOrEmpty(settings.BackgroundImagePath) && System.IO.File.Exists(settings.BackgroundImagePath))
            {
                try
                {
                    backgroundBrush = new System.Windows.Media.ImageBrush(new System.Windows.Media.Imaging.BitmapImage(new Uri(settings.BackgroundImagePath)));
                    backgroundBrush.Opacity = settings.BackgroundOpacity;
                    backgroundBrush.Stretch = System.Windows.Media.Stretch.UniformToFill;
                    RootGrid.Background = backgroundBrush;
                }
                catch { 
                    RootGrid.Background = System.Windows.Media.Brushes.Transparent;
                }
            }

            settings.Save();
        }

        private void BtnSelectBackground_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*",
                Title = "Chọn ảnh nền"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // Save image to AppData
                    string savedImagePath = AppSettings.SaveBackgroundImage(dialog.FileName);
                    
                    // Update settings
                    settings.BackgroundImagePath = savedImagePath;
                    settings.UseBackgroundImage = true;
                    settings.Save();
                    
                    // Apply image
                    backgroundBrush = new System.Windows.Media.ImageBrush(new System.Windows.Media.Imaging.BitmapImage(new Uri(savedImagePath)));
                    backgroundBrush.Opacity = settings.BackgroundOpacity;
                    backgroundBrush.Stretch = System.Windows.Media.Stretch.UniformToFill;
                    if (RootGrid != null)
                        RootGrid.Background = backgroundBrush;
                    
                    // Update UI
                    if (BackgroundImageCheckBox != null) BackgroundImageCheckBox.IsChecked = true;
                    if (BackgroundOpacityPanel != null)
                        BackgroundOpacityPanel.Visibility = Visibility.Visible;
                    
                    ShowSnackbar("Thành công", "Đã lưu và áp dụng ảnh nền", ControlAppearance.Success);
                }
                catch (Exception ex)
                {
                    ShowSnackbar("Lỗi", $"Không thể tải ảnh: {ex.Message}", ControlAppearance.Caution);
                    if (RootGrid != null)
                        RootGrid.Background = System.Windows.Media.Brushes.Transparent;
                }
            }
        }

        private void BackgroundOpacitySlider_ValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
        {
            if (settings == null || BackgroundOpacityValue == null) return;

            BackgroundOpacityValue.Text = $"{(int)(e.NewValue * 100)}%";
            
            if (backgroundBrush != null)
            {
                backgroundBrush.Opacity = e.NewValue;
            }

            settings.BackgroundOpacity = e.NewValue;
            settings.Save();
        }

        private async void BtnCheckUpdate_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                BtnCheckUpdate.IsEnabled = false;
                UpdateStatusText.Text = "Đang kiểm tra phiên bản mới...";

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);
                    
                    // Get version from GitHub
                    string versionUrl = "https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/refs/heads/main/backend/version";
                    string latestVersion = await httpClient.GetStringAsync(versionUrl);
                    latestVersion = latestVersion.Trim();

                    // Current version from field
                    string currentVersion = version;

                    if (latestVersion != currentVersion)
                    {
                        UpdateStatusText.Text = $"Phát hiện phiên bản mới: {latestVersion}";
                        
                        var result = System.Windows.MessageBox.Show(
                            $"Phiên bản hiện tại: {currentVersion}\nPhiên bản mới: {latestVersion}\n\nBạn có muốn tải về và cập nhật?",
                            "Cập nhật có sẵn",
                            System.Windows.MessageBoxButton.YesNo,
                            System.Windows.MessageBoxImage.Information);

                        if (result == System.Windows.MessageBoxResult.Yes)
                        {
                            await DownloadAndUpdate();
                        }
                    }
                    else
                    {
                        UpdateStatusText.Text = "Bạn đang sử dụng phiên bản mới nhất";
                        ShowSnackbar("Cập nhật", "Bạn đang sử dụng phiên bản mới nhất", ControlAppearance.Success);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatusText.Text = "Lỗi khi kiểm tra cập nhật";
                ShowSnackbar("Lỗi", $"Không thể kiểm tra cập nhật: {ex.Message}\n\nỨng dụng sẽ đóng sau 3 giây...", ControlAppearance.Danger, 3000);
                
                // Đóng ứng dụng sau 3 giây
                await Task.Delay(3000);
                System.Windows.Application.Current.Shutdown();
            }
            finally
            {
                BtnCheckUpdate.IsEnabled = true;
            }
        }

        private async Task DownloadAndUpdate()
        {
            try
            {
                UpdateStatusText.Text = "Đang tải về phiên bản mới...";
                ShowBusy("Đang tải về bản cập nhật...");

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(5);
                    
                    string downloadUrl = "https://github.com/AlexHerrySeek/ZibllWindows/releases/download/AutoUpdate/ZibllWindows.exe";

                    // Download to temp file
                    string tempPath = Path.Combine(Path.GetTempPath(), "ZibllWindows_Update.exe");
                    byte[] fileBytes = await httpClient.GetByteArrayAsync(downloadUrl);
                    File.WriteAllBytes(tempPath, fileBytes);

                    UpdateStatusText.Text = "Tải về hoàn tất. Đang chuẩn bị cập nhật...";

                    // Create update batch script
                    string currentExe = Process.GetCurrentProcess().MainModule?.FileName ?? Path.Combine(AppContext.BaseDirectory, "ZibllWindows.exe");
                    string batchScript = Path.Combine(Path.GetTempPath(), "update_zibll.bat");
                    
                    string batchContent = $@"@echo off
timeout /t 2 /nobreak > nul
taskkill /F /IM ZibllWindows.exe > nul 2>&1
timeout /t 1 /nobreak > nul
copy /Y ""{tempPath}"" ""{currentExe}"" > nul
del ""{tempPath}"" > nul
start """" ""{currentExe}""
del ""%~f0"" & exit";

                    File.WriteAllText(batchScript, batchContent);

                    // Start batch script and close current app
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = batchScript,
                        CreateNoWindow = true,
                        UseShellExecute = false
                    });

                    System.Windows.Application.Current.Shutdown();
                }
            }
            catch (Exception ex)
            {
                HideBusy();
                UpdateStatusText.Text = "Lỗi khi tải về cập nhật";
                ShowSnackbar("Lỗi", $"Không thể tải về: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void ThemeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ThemeComboBox == null) return;

            var selectedItem = ThemeComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (selectedItem == null) return;

            var theme = selectedItem.Content.ToString() ?? "Dark";
            
            try
            {
                switch (theme)
                {
                    case "Light":
                        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Light);
                        break;
                    case "Dark":
                        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);
                        break;
                    case "System":
                        // Default to dark theme
                        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(Wpf.Ui.Appearance.ApplicationTheme.Dark);
                        break;
                }

                if (settings != null)
                {
                    settings.Theme = theme;
                    settings.Save();
                }
                
                SetStatus($"Đã chuyển sang theme {theme}", InfoBarSeverity.Success);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Theme change error: {ex.Message}");
            }
        }

        private void AccentColorComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (AccentColorComboBox == null) return;

            var selectedItem = AccentColorComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (selectedItem == null) return;

            var colorName = selectedItem.Content.ToString() ?? "Mặc định";
            
            try
            {
                System.Windows.Media.Color accentColor;
                
                switch (colorName)
                {
                    case "Xanh dương":
                        accentColor = System.Windows.Media.Color.FromRgb(59, 130, 246); // Blue
                        break;
                    case "Tím":
                        accentColor = System.Windows.Media.Color.FromRgb(168, 85, 247); // Purple
                        break;
                    case "Hồng":
                        accentColor = System.Windows.Media.Color.FromRgb(236, 72, 153); // Pink
                        break;
                    case "Đỏ":
                        accentColor = System.Windows.Media.Color.FromRgb(239, 68, 68); // Red
                        break;
                    case "Cam":
                        accentColor = System.Windows.Media.Color.FromRgb(249, 115, 22); // Orange
                        break;
                    case "Xanh lá":
                        accentColor = System.Windows.Media.Color.FromRgb(34, 197, 94); // Green
                        break;
                    default: // Mặc định
                        accentColor = System.Windows.Media.Color.FromRgb(0, 120, 212); // Default Windows Blue
                        break;
                }

                Wpf.Ui.Appearance.ApplicationAccentColorManager.Apply(accentColor);

                if (settings != null)
                {
                    settings.AccentColor = colorName;
                    settings.Save();
                }
                
                ShowSnackbar("Thành công", $"Đã đổi màu chủ đạo sang {colorName}", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Accent color change error: {ex.Message}");
                ShowSnackbar("Lỗi", "Không thể thay đổi màu chủ đạo", ControlAppearance.Caution);
            }
        }

        private void BackgroundThemeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (BackgroundThemeComboBox == null) return;

            var selectedItem = BackgroundThemeComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (selectedItem == null) return;

            var content = selectedItem.Content.ToString() ?? "Không";
            var themeName = content.Contains(" ") ? content.Split(' ')[1] : content;
            
            try
            {
                if (themeName == "Không")
                {
                    if (backgroundBrush != null)
                        backgroundBrush.ImageSource = null;
                    if (RootGrid != null)
                        RootGrid.Background = System.Windows.Media.Brushes.Transparent;
                    if (settings != null)
                    {
                        settings.BackgroundTheme = "Không";
                        settings.UseBackgroundImage = false;
                        settings.Save();
                    }
                    if (BackgroundImageCheckBox != null)
                        BackgroundImageCheckBox.IsChecked = false;
                    ShowSnackbar("Thành công", "Đã tắt hình nền", ControlAppearance.Success);
                    return;
                }

                // Create gradient backgrounds for themes
                var gradientBrush = new System.Windows.Media.LinearGradientBrush();
                gradientBrush.StartPoint = new System.Windows.Point(0, 0);
                gradientBrush.EndPoint = new System.Windows.Point(1, 1);

                switch (themeName)
                {
                    case "Galaxy":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(17, 24, 39), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(88, 28, 135), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(0, 0, 0), 1));
                        break;
                    case "Sunset":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(251, 146, 60), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(239, 68, 68), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(124, 58, 237), 1));
                        break;
                    case "Ocean":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(6, 182, 212), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(59, 130, 246), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(30, 64, 175), 1));
                        break;
                    case "Forest":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(134, 239, 172), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(34, 197, 94), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(21, 128, 61), 1));
                        break;
                    case "Mountain":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(226, 232, 240), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(148, 163, 184), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(51, 65, 85), 1));
                        break;
                    case "City":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(30, 41, 59), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(51, 65, 85), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(15, 23, 42), 1));
                        break;
                    case "Abstract":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(236, 72, 153), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(168, 85, 247), 0.3));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(59, 130, 246), 0.7));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(34, 197, 94), 1));
                        break;
                    case "Gemstone":
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(56, 189, 248), 0));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(147, 51, 234), 0.5));
                        gradientBrush.GradientStops.Add(new System.Windows.Media.GradientStop(System.Windows.Media.Color.FromRgb(219, 39, 119), 1));
                        break;
                }

                gradientBrush.Opacity = settings?.BackgroundOpacity ?? 0.3;
                this.Background = gradientBrush;

                if (settings != null)
                {
                    settings.BackgroundTheme = themeName;
                    settings.UseBackgroundImage = true;
                    settings.Save();
                }

                if (BackgroundImageCheckBox != null)
                    BackgroundImageCheckBox.IsChecked = true;
                if (BackgroundOpacityPanel != null)
                    BackgroundOpacityPanel.Visibility = System.Windows.Visibility.Visible;

                ShowSnackbar("Thành công", $"Đã áp dụng giao diện {themeName}", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Background theme error: {ex.Message}");
                ShowSnackbar("Lỗi", "Không thể thay đổi giao diện", ControlAppearance.Caution);
            }
        }

        private void AnimationCheckBox_Changed(object sender, System.Windows.RoutedEventArgs e)
        {
            if (AnimationCheckBox == null || settings == null) return;
            settings.UseAnimations = AnimationCheckBox.IsChecked ?? true;
            settings.Save();
            ShowSnackbar("Thành công", $"Đã {(settings.UseAnimations ? "bật" : "tắt")} hiệu ứng động", ControlAppearance.Success);
        }

        private void BorderRadiusSlider_ValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
        {
            if (BorderRadiusValue == null || settings == null || BorderRadiusSlider == null) return;
            var value = (int)BorderRadiusSlider.Value;
            BorderRadiusValue.Text = $"{value}px";
            settings.BorderRadius = value;
            settings.Save();
        }

        private void CompactModeCheckBox_Changed(object sender, System.Windows.RoutedEventArgs e)
        {
            if (CompactModeCheckBox == null || settings == null) return;
            settings.CompactMode = CompactModeCheckBox.IsChecked ?? false;
            settings.Save();
            
            // Toggle NavigationView pane display mode
            var navView = this.FindName("MainNavigation") as Wpf.Ui.Controls.NavigationView;
            if (navView != null)
            {
                navView.IsPaneOpen = !settings.CompactMode;
            }
            
            ShowSnackbar("Thành công", $"Đã {(settings.CompactMode ? "bật" : "tắt")} chế độ thu gọn", ControlAppearance.Success);
        }

        private void BackdropComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (BackdropComboBox == null || settings == null) return;

            var selectedItem = BackdropComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem;
            if (selectedItem == null) return;

            var backdropType = selectedItem.Content.ToString() ?? "Mặc Định";
            
            try
            {
                switch (backdropType)
                {
                    case "Mặc Định":
                        Wpf.Ui.Appearance.SystemThemeWatcher.UnWatch(this);
                        break;
                    case "Mica":
                        Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this, Wpf.Ui.Controls.WindowBackdropType.Mica);
                        break;
                    case "Acrylic":
                        Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this, Wpf.Ui.Controls.WindowBackdropType.Acrylic);
                        break;
                    case "Tabbed":
                        Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this, Wpf.Ui.Controls.WindowBackdropType.Tabbed);
                        break;
                }

                settings.BackdropType = backdropType;
                settings.Save();
                ShowSnackbar("Thành công", $"Đã áp dụng hiệu ứng nền {backdropType}", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Backdrop change error: {ex.Message}");
                ShowSnackbar("Lỗi", "Không thể thay đổi hiệu ứng nền", ControlAppearance.Caution);
            }
        }

        private async void BtnResetSettings_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = "Bạn có chắc chắn muốn đặt lại tất cả cài đặt về mặc định?\n\nỨng dụng sẽ khởi động lại sau khi đặt lại.",
                PrimaryButtonText = "Đặt lại",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                settings = new AppSettings();
                settings.Save();
                
                ShowSnackbar("Thành công", "Đang khởi động lại...", ControlAppearance.Success);
                await Task.Delay(1000);
                
                System.Diagnostics.Process.Start(Environment.ProcessPath ?? "ZibllWindows.exe");
                System.Windows.Application.Current.Shutdown();
            }
        }

        private void CheckAdminPrivileges()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            
            if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                ShowSnackbar(" Cảnh báo", "Chưa có quyền Administrator. Một số chức năng có thể không hoạt động.", ControlAppearance.Caution, 5000);
            }
        }

        // Windows Activation Methods
        private async void BtnActivateWindowsHWID_Click(object sender, RoutedEventArgs e)
        {
            await ActivateWindowsWithProductKey();
        }

        private async void BtnActivateWindowsHWID2_Click(object sender, RoutedEventArgs e)
        {
            await ActivateWindowsWithProductKey();
        }

        private async void BtnActivateWindowsKMS38_Click(object sender, RoutedEventArgs e)
        {
            await ActivateWindowsWithProductKey();
        }

        private async void BtnActivateWindowsOnlineKMS_Click(object sender, RoutedEventArgs e)
        {
            await ActivateWindowsWithProductKey();
        }

        // Office Activation Methods
        private async void BtnActivateOfficeOhook_Click(object sender, RoutedEventArgs e)
        {
            await ActivateOfficeWithProductKey();
        }

        private async void BtnActivateOfficeOhook2_Click(object sender, RoutedEventArgs e)
        {
            await ActivateOfficeWithProductKey();
        }

        private async void BtnActivateOfficeKMS_Click(object sender, RoutedEventArgs e)
        {
            await ActivateOfficeWithProductKey();
        }

        // New activation methods using product key
        private async Task ActivateWindowsWithProductKey()
        {
            try
            {
                var dialog = new ProductKeyDialog
                {
                    Owner = this
                };

                if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.ProductKey))
                {
                    await ActivateWindowsWithKey(dialog.ProductKey);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi mở dialog: {ex.Message}");
            }
        }

        private async Task ActivateOfficeWithProductKey()
        {
            try
            {
                var dialog = new ProductKeyDialog
                {
                    Owner = this
                };

                if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.ProductKey))
                {
                    await ActivateOfficeWithKey(dialog.ProductKey);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi mở dialog: {ex.Message}");
            }
        }

        private async Task ActivateWindowsWithKey(string productKey)
        {
            try
            {
                ShowBusy("Đang kích hoạt Windows...");
                SetStatus("Đang kích hoạt Windows với Product Key...", InfoBarSeverity.Informational);
                DisableButtons();

                // Format key with hyphens
                string formattedKey = FormatProductKey(productKey);

                // Step 1: Uninstall existing key (if any)
                await ExecuteSlmgrCommand("/upk");

                // Step 2: Install the new product key
                bool installSuccess = await ExecuteSlmgrCommand($"/ipk {formattedKey}");

                if (!installSuccess)
                {
                    SetStatus("Lỗi: Không thể cài đặt Product Key!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync("Lỗi", 
                        "Không thể cài đặt Product Key.\n\n" +
                        "Nguyên nhân có thể:\n" +
                        "1. Product Key không hợp lệ\n" +
                        "2. Product Key không phù hợp với phiên bản Windows hiện tại\n" +
                        "3. Cần quyền Administrator");
                    return;
                }

                // Step 3: Activate Windows
                SetStatus("Đang kích hoạt Windows...", InfoBarSeverity.Warning);
                bool activateSuccess = await ExecuteSlmgrCommand("/ato");

                if (activateSuccess)
                {
                    await Task.Delay(2000);
                    await CheckActivationStatus();
                    SetStatus("Windows đã được kích hoạt thành công!", InfoBarSeverity.Success);
                    await ShowInfoDialogAsync("Thành công", 
                        "Windows đã được kích hoạt thành công với Product Key của bạn!");
                }
                else
                {
                    SetStatus("Kích hoạt thất bại!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync("Lỗi", 
                        "Không thể kích hoạt Windows.\n\n" +
                        "Nguyên nhân có thể:\n" +
                        "1. Product Key đã được sử dụng trên máy khác\n" +
                        "2. Không có kết nối Internet để xác thực\n" +
                        "3. Product Key không hợp lệ hoặc đã hết hạn");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi kích hoạt Windows: {ex.Message}");
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async Task ActivateOfficeWithKey(string productKey)
        {
            try
            {
                ShowBusy("Đang kích hoạt Office...");
                SetStatus("Đang kích hoạt Office với Product Key...", InfoBarSeverity.Informational);
                DisableButtons();

                string officePath = FindOfficePath();
                if (string.IsNullOrEmpty(officePath))
                {
                    SetStatus("Không tìm thấy Office!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync("Lỗi", "Không tìm thấy cài đặt Microsoft Office trên hệ thống.");
                    return;
                }

                // Format key with hyphens
                string formattedKey = FormatProductKey(productKey);

                // Step 1: Install the product key
                SetStatus("Đang cài đặt Product Key...", InfoBarSeverity.Informational);
                bool installSuccess = await ExecuteOfficeCommand($"/inpkey:{formattedKey}");

                if (!installSuccess)
                {
                    SetStatus("Lỗi: Không thể cài đặt Product Key!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync("Lỗi", 
                        "Không thể cài đặt Product Key.\n\n" +
                        "Nguyên nhân có thể:\n" +
                        "1. Product Key không hợp lệ\n" +
                        "2. Product Key không phù hợp với phiên bản Office hiện tại");
                    return;
                }

                // Step 2: Activate Office
                SetStatus("Đang kích hoạt Office...", InfoBarSeverity.Warning);
                bool activateSuccess = await ExecuteOfficeCommand("/act");

                if (activateSuccess)
                {
                    await Task.Delay(2000);
                    await ExecuteOfficeCommand("/dstatus");
                    SetStatus("Office đã được kích hoạt thành công!", InfoBarSeverity.Success);
                    await ShowInfoDialogAsync("Thành công", 
                        "Office đã được kích hoạt thành công với Product Key của bạn!");
                }
                else
                {
                    SetStatus("Kích hoạt thất bại!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync("Lỗi", 
                        "Không thể kích hoạt Office.\n\n" +
                        "Nguyên nhân có thể:\n" +
                        "1. Product Key đã được sử dụng trên máy khác\n" +
                        "2. Không có kết nối Internet để xác thực\n" +
                        "3. Product Key không hợp lệ hoặc đã hết hạn");
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi kích hoạt Office: {ex.Message}");
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private string FormatProductKey(string key)
        {
            // Remove all non-alphanumeric characters
            string cleanKey = System.Text.RegularExpressions.Regex.Replace(key.ToUpper(), "[^A-Z0-9]", "");
            
            // Format as XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
            if (cleanKey.Length == 25)
            {
                return $"{cleanKey.Substring(0, 5)}-{cleanKey.Substring(5, 5)}-{cleanKey.Substring(10, 5)}-{cleanKey.Substring(15, 5)}-{cleanKey.Substring(20, 5)}";
            }
            
            return key; // Return as-is if not 25 characters
        }

        private async Task ShowConfirmAndActivate(string message, string method)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = $"Bạn có muốn kích hoạt {message}?",
                PrimaryButtonText = "OK",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await ActivateWithMethod(method);
            }
        }

        private async Task ActivateWithMethod(string method)
        {
            try
            {
                ShowBusy($"Đang kích hoạt ({method.ToUpper()})...");
                SetStatus($"Đang kích hoạt bằng phương pháp {method.ToUpper()}...", InfoBarSeverity.Warning);
                DisableButtons();

                bool success = false;
                string productName = GetWindowsProductName();

                if (method == "kms" || method == "kms_office" || method == "kms38")
                {
                    if (method == "kms_office")
                    {
                        success = await ActivateOfficeWithKMS();
                    }
                    else if (IsWindowsServer(productName))
                    {
                        success = await ActivateWindowsServerWithKMS(productName);
                    }
                    else
                    {
                        success = await ActivateWindowsClientWithKMS(productName);
                    }
                }
                else if (method == "hwid")
                {
                    success = await ActivateWithHWID();
                }
                else if (method == "ohook")
                {
                    success = await ActivateWithOhook();
                }
                else
                {
                    SetStatus("Phương pháp không được hỗ trợ!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync(" Lỗi", "Phương pháp kích hoạt không được hỗ trợ.");
                    return;
                }

                if (success)
                {
                    SetStatus("Kích hoạt thành công!", InfoBarSeverity.Success);
                    await ShowInfoDialogAsync(" Kích hoạt thành công", "Quá trình kích hoạt đã hoàn tất.");
                }
                else
                {
                    SetStatus("Kích hoạt thất bại!", InfoBarSeverity.Error);
                    await ShowInfoDialogAsync(" Kích hoạt thất bại", "Có lỗi xảy ra trong quá trình kích hoạt.");
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
                await ShowInfoDialogAsync(" Lỗi", $"Lỗi: {ex.Message}");
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private string GetWindowsProductName()
        {
            try
            {
                object value = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "");
                return value?.ToString() ?? "Unknown";
            }
            catch (Exception)
            {
                return "Unknown";
            }
        }

        private bool IsWindowsServer(string productName)
        {
            if (string.IsNullOrEmpty(productName))
                return false;

            string lowerProductName = productName.ToLower();
            return lowerProductName.Contains("server") ||
                   lowerProductName.Contains("datacenter") ||
                   (lowerProductName.Contains("standard") && lowerProductName.Contains("server"));
        }

        private async Task<bool> ActivateWindowsClientWithKMS(string productName)
        {
            try
            {
                // Dictionary với mapping chính xác cho tất cả các phiên bản Windows
                Dictionary<string, string> winKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Windows 11
            {"Windows 11 Home", "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"},
            {"Windows 11 Home N", "3KHY7-WNT83-DGQKR-F7HPR-844BM"},
            {"Windows 11 Pro", "W269N-WFGWX-YVC9B-4J6C9-T83GX"},
            {"Windows 11 Pro N", "MH37W-N47XK-V7XM9-C7227-GCQG9"},
            {"Windows 11 Pro for Workstations", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"},
            {"Windows 11 Pro for Workstations N", "9FNHH-K3HBT-3W4TD-6383H-6XYWF"},
            {"Windows 11 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43"},
            {"Windows 11 Enterprise N", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"},
            {"Windows 11 Education", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"},
            {"Windows 11 Education N", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"},
            
            // Windows 10
            {"Windows 10 Home", "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"},
            {"Windows 10 Home N", "3KHY7-WNT83-DGQKR-F7HPR-844BM"},
            {"Windows 10 Pro", "W269N-WFGWX-YVC9B-4J6C9-T83GX"},
            {"Windows 10 Pro N", "MH37W-N47XK-V7XM9-C7227-GCQG9"},
            {"Windows 10 Pro for Workstations", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"},
            {"Windows 10 Pro for Workstations N", "9FNHH-K3HBT-3W4TD-6383H-6XYWF"},
            {"Windows 10 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43"},
            {"Windows 10 Enterprise N", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"},
            {"Windows 10 Education", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"},
            {"Windows 10 Education N", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"},
            
            // Windows 8.1
            {"Windows 8.1 Core", "M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK"},
            {"Windows 8.1 Core N", "7B9N3-D94CG-YTVHR-QBPX3-RJP64"},
            {"Windows 8.1 Pro", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9"},
            {"Windows 8.1 Pro N", "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY"},
            {"Windows 8.1 Enterprise", "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7"},
            {"Windows 8.1 Enterprise N", "TT4HM-HN7YT-62K67-RGRQJ-JFFXW"},
            
            // Windows 7
            {"Windows 7 Home Basic", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"},
            {"Windows 7 Home Premium", "HQRJW-XYQRP-JVBXW-GV3P3-M3B47"},
            {"Windows 7 Professional", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"},
            {"Windows 7 Enterprise", "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"},
            {"Windows 7 Ultimate", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"}
        };

                // Chuẩn hóa productName
                string normalizedProductName = productName.Trim();
                SetStatus($"Đang xác định phiên bản: {normalizedProductName}", InfoBarSeverity.Informational);

                // Tìm key chính xác
                string key = "";
                bool foundExactMatch = false;

                // Tìm kiếm chính xác trong dictionary
                foreach (var winVersion in winKeys.Keys)
                {
                    if (normalizedProductName.Equals(winVersion, StringComparison.OrdinalIgnoreCase) ||
                        normalizedProductName.Contains(winVersion, StringComparison.OrdinalIgnoreCase))
                    {
                        key = winKeys[winVersion];
                        foundExactMatch = true;
                        SetStatus($"Đã tìm thấy key cho: {winVersion}", InfoBarSeverity.Success);
                        break;
                    }
                }

                // Nếu không tìm thấy exact match, phân tích tự động
                if (!foundExactMatch)
                {
                    SetStatus("Đang phân tích phiên bản Windows...", InfoBarSeverity.Warning);

                    // Xác định các thông tin từ productName
                    bool isWindows11 = normalizedProductName.Contains("11", StringComparison.OrdinalIgnoreCase);
                    bool isWindows10 = normalizedProductName.Contains("10", StringComparison.OrdinalIgnoreCase);
                    bool isWindows8 = normalizedProductName.Contains("8", StringComparison.OrdinalIgnoreCase);
                    bool isWindows7 = normalizedProductName.Contains("7", StringComparison.OrdinalIgnoreCase);

                    // Xác định edition
                    bool isHomeEdition = normalizedProductName.Contains("Home", StringComparison.OrdinalIgnoreCase) ||
                                        normalizedProductName.Contains("Core", StringComparison.OrdinalIgnoreCase);

                    bool isProEdition = normalizedProductName.Contains("Professional", StringComparison.OrdinalIgnoreCase) ||
                                       normalizedProductName.Contains("Pro", StringComparison.OrdinalIgnoreCase) ||
                                       normalizedProductName.Contains("Business", StringComparison.OrdinalIgnoreCase);

                    bool isEnterpriseEdition = normalizedProductName.Contains("Enterprise", StringComparison.OrdinalIgnoreCase);
                    bool isEducationEdition = normalizedProductName.Contains("Education", StringComparison.OrdinalIgnoreCase);
                    bool isUltimateEdition = normalizedProductName.Contains("Ultimate", StringComparison.OrdinalIgnoreCase);
                    bool isWorkstationEdition = normalizedProductName.Contains("Workstation", StringComparison.OrdinalIgnoreCase);
                    bool hasN = normalizedProductName.Contains(" N", StringComparison.OrdinalIgnoreCase) ||
                               normalizedProductName.EndsWith(" N", StringComparison.OrdinalIgnoreCase);

                    // Chọn key dựa trên phiên bản và edition
                    if (isEnterpriseEdition)
                    {
                        key = hasN ? "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" : "NPPR9-FWDCX-D2C8J-H872K-2YT43";
                        SetStatus($"Đã xác định: Windows {(isWindows11 ? "11" : isWindows10 ? "10" : "")} Enterprise Edition", InfoBarSeverity.Informational);
                    }
                    else if (isEducationEdition)
                    {
                        key = hasN ? "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" : "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
                        SetStatus($"Đã xác định: Windows {(isWindows11 ? "11" : isWindows10 ? "10" : "")} Education Edition", InfoBarSeverity.Informational);
                    }
                    else if (isWorkstationEdition)
                    {
                        key = hasN ? "9FNHH-K3HBT-3W4TD-6383H-6XYWF" : "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J";
                        SetStatus($"Đã xác định: Windows {(isWindows11 ? "11" : isWindows10 ? "10" : "")} Pro for Workstations", InfoBarSeverity.Informational);
                    }
                    else if (isProEdition || isUltimateEdition)
                    {
                        key = hasN ? "MH37W-N47XK-V7XM9-C7227-GCQG9" : "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                        SetStatus($"Đã xác định: Windows {(isWindows11 ? "11" : isWindows10 ? "10" : "")} Professional Edition", InfoBarSeverity.Informational);
                    }
                    else if (isHomeEdition)
                    {
                        key = hasN ? "3KHY7-WNT83-DGQKR-F7HPR-844BM" : "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
                        SetStatus($"Đã xác định: Windows {(isWindows11 ? "11" : isWindows10 ? "10" : "")} Home Edition", InfoBarSeverity.Informational);
                    }
                    else
                    {
                        // Mặc định sử dụng Professional key (phổ biến nhất)
                        key = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                        SetStatus($"Không xác định rõ edition, sử dụng key mặc định", InfoBarSeverity.Warning);
                    }
                }

                // Kiểm tra và đề xuất chuyển edition nếu cần
                await CheckAndConvertEditionIfNeeded(productName, key);

                SetStatus($"Đang cài đặt Product Key...", InfoBarSeverity.Informational);

                // Danh sách KMS servers
                string[] kmsServers = {
            "kms8.msguides.com",
            "kms.chinancce.com",
            "kms.digiboy.ir",
            "kms.lotro.cc",
            "kms.srv.crsoo.com",
            "kms.03k.org"
        };

                // Bước 1: Gỡ bỏ key hiện tại trước (nếu có)
                await ExecuteSlmgrCommand("/upk");

                // Bước 2: Cài đặt key mới
                bool installKeySuccess = await ExecuteSlmgrCommand($"/ipk {key}");

                if (!installKeySuccess)
                {
                    // Thử sửa lỗi edition mismatch
                    SetStatus("Đang thử sửa lỗi không khớp edition...", InfoBarSeverity.Warning);
                    bool fixSuccess = await FixEditionMismatch(productName);

                    if (!fixSuccess)
                    {
                        SetStatus("Lỗi cài đặt product key!", InfoBarSeverity.Error);
                        await ShowInfoDialogAsync("Lỗi",
                            $"Không thể cài đặt product key cho phiên bản: {productName}\n\n" +
                            "Nguyên nhân có thể:\n" +
                            "1. Phiên bản Windows không hỗ trợ KMS\n" +
                            "2. Cần chuyển đổi edition trước\n" +
                            "3. Product key không phù hợp\n\n" +
                            "Hãy thử:\n" +
                            "1. Kiểm tra lại phiên bản Windows\n" +
                            "2. Sử dụng phương pháp HWID cho Windows Home\n" +
                            "3. Sử dụng 'Đổi phiên bản' trong công cụ");
                        return false;
                    }
                }

                // Bước 3: Thiết lập KMS server
                bool setKmsSuccess = false;
                foreach (var kmsServer in kmsServers)
                {
                    SetStatus($"Đang thử KMS server: {kmsServer}...", InfoBarSeverity.Warning);
                    setKmsSuccess = await ExecuteSlmgrCommand($"/skms {kmsServer}");

                    if (setKmsSuccess)
                    {
                        SetStatus($"Đã kết nối đến KMS server: {kmsServer}", InfoBarSeverity.Success);
                        break;
                    }
                }

                if (!setKmsSuccess)
                {
                    SetStatus("Không thể kết nối đến bất kỳ KMS server nào!", InfoBarSeverity.Error);
                    return false;
                }

                // Bước 4: Kích hoạt
                SetStatus("Đang kích hoạt Windows...", InfoBarSeverity.Warning);
                bool activateSuccess = await ExecuteSlmgrCommand("/ato");

                if (!activateSuccess)
                {
                    // Thử lại với phương pháp khác
                    SetStatus("Đang thử kích hoạt lại...", InfoBarSeverity.Warning);

                    // Reset và thử lại
                    await ExecuteSlmgrCommand("/rearm");
                    await Task.Delay(2000);

                    activateSuccess = await ExecuteSlmgrCommand("/ato");

                    if (!activateSuccess)
                    {
                        // Kiểm tra lỗi cụ thể
                        await CheckActivationStatus();

                        // Hiển thị hướng dẫn chi tiết
                        await ShowInfoDialogAsync("Lỗi kích hoạt",
                            $"Không thể kích hoạt Windows phiên bản: {productName}\n\n" +
                            "Các giải pháp:\n" +
                            "1. Kiểm tra kết nối Internet\n" +
                            "2. Thử phương pháp HWID thay vì KMS\n" +
                            "3. Chạy 'Khắc phục sự cố' trong công cụ\n" +
                            "4. Thử kích hoạt thủ công bằng lệnh:\n" +
                            $"   slmgr /ipk {key}\n" +
                            "   slmgr /skms kms8.msguides.com\n" +
                            "   slmgr /ato");

                        return false;
                    }
                }

                // Bước 5: Kiểm tra trạng thái kích hoạt
                if (activateSuccess)
                {
                    await Task.Delay(3000);
                    bool statusSuccess = await CheckActivationStatus2();

                    if (statusSuccess)
                    {
                        SetStatus("Windows đã được kích hoạt thành công!", InfoBarSeverity.Success);
                        await ShowInfoDialogAsync("Thành công",
                            $"Windows đã được kích hoạt thành công!\n\n" +
                            $"Phiên bản: {productName}\n" +
                            $"Phương pháp: Online KMS (180 ngày)\n\n" +
                            "Kích hoạt sẽ tự động gia hạn mỗi 180 ngày.");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi KMS: {ex.Message}", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi kích hoạt Windows:\n\n{ex.Message}");
                return false;
            }
        }

        private async Task<bool> CheckAndConvertEditionIfNeeded(string productName, string targetKey)
        {
            try
            {
                // Kiểm tra xem có cần chuyển edition không
                // Ví dụ: từ Core/Home lên Professional
                if (productName.Contains("Home", StringComparison.OrdinalIgnoreCase) ||
                    productName.Contains("Core", StringComparison.OrdinalIgnoreCase))
                {
                    if (targetKey == "W269N-WFGWX-YVC9B-4J6C9-T83GX") // Professional key
                    {
                        SetStatus("Đang nâng cấp từ Home/Core lên Professional...", InfoBarSeverity.Warning);

                        // Sử dụng changepk.exe để thay đổi product key và edition
                        string changePkCommand = $"/ProductKey {targetKey}";
                        bool changeSuccess = await RunProcessCheckSuccess("changepk.exe", changePkCommand);

                        if (changeSuccess)
                        {
                            SetStatus("Đã yêu cầu nâng cấp edition. Cần khởi động lại!", InfoBarSeverity.Success);
                            await ShowInfoDialogAsync("Thông báo",
                                "Để hoàn tất nâng cấp lên phiên bản Professional, cần khởi động lại máy tính.");
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi chuyển edition: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        private async Task<bool> CheckActivationStatus2()
        {
            try
            {
                SetStatus("Đang kiểm tra trạng thái kích hoạt...", InfoBarSeverity.Informational);

                // Kiểm tra bằng slmgr
                bool status = await ExecuteSlmgrCommand("/xpr"); // Hiển thị thời hạn
                await Task.Delay(1000);
                await ExecuteSlmgrCommand("/dlv"); // Hiển thị chi tiết

                return status;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi kiểm tra trạng thái: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        private async Task<bool> ExecuteSlmgrCommand(string arguments)
        {
            try
            {
                string systemRoot = Environment.GetEnvironmentVariable("SystemRoot");
                if (string.IsNullOrEmpty(systemRoot))
                {
                    systemRoot = @"C:\Windows";
                }

                string slmgrPath = Path.Combine(systemRoot, "System32", "slmgr.vbs");

                if (!File.Exists(slmgrPath))
                {
                    SetStatus("Không tìm thấy slmgr.vbs!", InfoBarSeverity.Error);
                    return false;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = "cscript.exe",
                    Arguments = $"//Nologo \"{slmgrPath}\" {arguments}",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = Path.Combine(systemRoot, "System32")
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();

                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();

                    await Task.Run(() => process.WaitForExit(30000));

                    // Xử lý output để kiểm tra thành công
                    bool success = process.ExitCode == 0;

                    // Kiểm tra các thông báo lỗi phổ biến
                    if (output.Contains("0xC004F069") || output.Contains("non-core edition"))
                    {
                        SetStatus("Lỗi: Không thể kích hoạt phiên bản này với key hiện tại. Cần đúng edition!", InfoBarSeverity.Error);
                        return false;
                    }

                    if (output.Contains("0x8007007B") || output.Contains("0xC004C003"))
                    {
                        SetStatus("Lỗi: Product key không hợp lệ hoặc đã hết hạn!", InfoBarSeverity.Error);
                        return false;
                    }

                    if (output.Contains("successfully") ||
                        output.Contains("Product activated successfully") ||
                        output.Contains("Key Management Service") ||
                        output.Contains("will expire"))
                    {
                        success = true;
                    }

                    if (output.Contains("Error") || output.Contains("failed"))
                    {
                        success = false;
                    }

                    if (!success)
                    {
                        string errorMsg = !string.IsNullOrEmpty(error) ? error : output;

                        // Hiển thị thông báo lỗi chi tiết
                        if (errorMsg.Contains("0xC004F069"))
                        {
                            SetStatus("Lỗi: Phiên bản Windows không khớp với product key!", InfoBarSeverity.Error);
                            await ShowInfoDialogAsync("Lỗi kích hoạt",
                                "Lỗi 0xC004F069: Phiên bản Windows hiện tại không tương thích với product key được cài đặt.\n\n" +
                                "Nguyên nhân:\n" +
                                "1. Đang cố gắng kích hoạt Windows Home bằng key Professional\n" +
                                "2. Đang cố gắng kích hoạt Windows Pro bằng key Enterprise\n" +
                                "3. Key không đúng cho edition hiện tại\n\n" +
                                "Giải pháp:\n" +
                                "1. Kiểm tra phiên bản Windows chính xác\n" +
                                "2. Sử dụng đúng key cho phiên bản đó\n" +
                                "3. Nâng cấp edition nếu cần");
                        }
                        else
                        {
                            SetStatus($"slmgr: {errorMsg.Trim()}", InfoBarSeverity.Warning);
                        }
                    }
                    else
                    {
                        SetStatus($"slmgr: Thành công", InfoBarSeverity.Success);
                    }

                    return success;
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi thực thi slmgr: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        // Phương thức mới: Tự động phát hiện và sửa lỗi edition
        private async Task<bool> FixEditionMismatch(string currentProductName)
        {
            try
            {
                SetStatus("Đang xử lý lỗi không khớp edition...", InfoBarSeverity.Warning);

                // Phát hiện edition hiện tại
                string detectedEdition = DetectWindowsEdition(currentProductName);

                // Lấy key phù hợp cho edition hiện tại
                string correctKey = GetKeyForEdition(detectedEdition);

                if (string.IsNullOrEmpty(correctKey))
                {
                    SetStatus("Không thể xác định key phù hợp!", InfoBarSeverity.Error);
                    return false;
                }

                // Cài đặt key đúng
                SetStatus($"Đang cài đặt key phù hợp cho {detectedEdition}...", InfoBarSeverity.Informational);
                return await ExecuteSlmgrCommand($"/ipk {correctKey}");
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi sửa edition: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        private string DetectWindowsEdition(string productName)
        {
            if (productName.Contains("Enterprise", StringComparison.OrdinalIgnoreCase))
                return "Enterprise";
            if (productName.Contains("Professional", StringComparison.OrdinalIgnoreCase) ||
                productName.Contains("Pro", StringComparison.OrdinalIgnoreCase))
                return "Professional";
            if (productName.Contains("Home", StringComparison.OrdinalIgnoreCase) ||
                productName.Contains("Core", StringComparison.OrdinalIgnoreCase))
                return "Home";
            if (productName.Contains("Education", StringComparison.OrdinalIgnoreCase))
                return "Education";

            return "Unknown";
        }

        private string GetKeyForEdition(string edition)
        {
            return edition.ToLower() switch
            {
                "enterprise" => "NPPR9-FWDCX-D2C8J-H872K-2YT43",
                "professional" or "pro" => "W269N-WFGWX-YVC9B-4J6C9-T83GX",
                "home" or "core" => "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
                "education" => "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2",
                _ => "W269N-WFGWX-YVC9B-4J6C9-T83GX" // Mặc định là Professional
            };
        }

        // THÊM PHƯƠNG THỨC ActivateWindowsServerWithKMS với tham số
        private async Task<bool> ActivateWindowsServerWithKMS(string productName)
        {
            try
            {
                Dictionary<string, string> serverKeys = new Dictionary<string, string>()
                {
                    {"Windows Server 2022 Standard", "VDYBN-27WPP-V4HQT-9VMD4-VMK7H"},
                    {"Windows Server 2022 Datacenter", "WX4NM-KYWYW-QJJR4-XV3QB-6VM33"},
                    {"Windows Server 2019 Standard", "N69G4-B89J2-4G8F4-WWYCC-J464C"},
                    {"Windows Server 2019 Datacenter", "WMDGN-G9PQG-XVVXX-R3X43-63DFG"},
                    {"Windows Server 2016 Standard", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY"},
                    {"Windows Server 2016 Datacenter", "CB7KF-BWN84-R7R2Y-793K2-8XDDG"},
                    {"Windows Server 2012 R2 Standard", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX"},
                    {"Windows Server 2012 R2 Datacenter", "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"},
                    {"Windows Server 2008 R2 Standard", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC"},
                    {"Windows Server 2008 R2 Datacenter", "74YFP-3QFB3-KQT8W-PMXWJ-7M648"},
                    {"Windows Server 2008 R2 Enterprise", "489J6-VHDMP-X63PK-3K798-CPX3Y"}
                };

                string key = "";
                foreach (var serverVersion in serverKeys.Keys)
                {
                    if (productName.Contains(serverVersion))
                    {
                        key = serverKeys[serverVersion];
                        break;
                    }
                }

                if (string.IsNullOrEmpty(key))
                {
                    if (productName.ToLower().Contains("evaluation"))
                    {
                        SetStatus("Đang chuyển đổi Evaluation version...", InfoBarSeverity.Warning);

                        string edition = "";
                        if (productName.Contains("Standard")) edition = "Standard";
                        else if (productName.Contains("Datacenter")) edition = "Datacenter";
                        else if (productName.Contains("Enterprise")) edition = "Enterprise";

                        if (!string.IsNullOrEmpty(edition))
                        {
                            string serverKey = GetServerKeyByEdition(edition);
                            string args = $"/online /set-edition:Server{edition} /productkey:{serverKey} /accepteula";
                            bool dismResult = await RunProcessCheckSuccess("dism.exe", args);

                            if (dismResult)
                            {
                                SetStatus("Yêu cầu khởi động lại để hoàn tất!", InfoBarSeverity.Success);
                                await ShowInfoDialogAsync("Thông báo",
                                    "Để nâng cấp lên phiên bản có bản quyền, cần khởi động lại máy tính.");
                                return true;
                            }
                        }
                    }

                    if (productName.Contains("Standard"))
                        key = "VDYBN-27WPP-V4HQT-9VMD4-VMK7H";
                    else if (productName.Contains("Datacenter"))
                        key = "WX4NM-KYWYW-QJJR4-XV3QB-6VM33";
                    else if (productName.Contains("Enterprise"))
                        key = "489J6-VHDMP-X63PK-3K798-CPX3Y";
                    else
                    {
                        SetStatus($"Phiên bản Server không được hỗ trợ: {productName}", InfoBarSeverity.Error);
                        return false;
                    }
                }

                SetStatus($"Đang cài đặt Product Key cho {productName}...", InfoBarSeverity.Informational);

                string[] kmsServers = {
                    "kms8.msguides.com",
                    "kms.chinancce.com",
                    "kms.digiboy.ir",
                    "kms.lotro.cc"
                };

                bool installKeySuccess = await ExecuteSlmgrCommand($"/ipk {key}");

                if (!installKeySuccess)
                {
                    SetStatus("Lỗi cài đặt product key!", InfoBarSeverity.Error);
                    return false;
                }

                bool setKmsSuccess = false;
                foreach (var kmsServer in kmsServers)
                {
                    SetStatus($"Đang thử KMS server: {kmsServer}...", InfoBarSeverity.Warning);
                    setKmsSuccess = await ExecuteSlmgrCommand($"/skms {kmsServer}");

                    if (setKmsSuccess)
                    {
                        SetStatus($"Đã kết nối đến KMS server: {kmsServer}", InfoBarSeverity.Success);
                        break;
                    }
                }

                if (!setKmsSuccess)
                {
                    SetStatus("Không thể kết nối đến bất kỳ KMS server nào!", InfoBarSeverity.Error);
                    return false;
                }

                SetStatus("Đang kích hoạt Windows Server...", InfoBarSeverity.Warning);
                bool activateSuccess = await ExecuteSlmgrCommand("/ato");

                if (activateSuccess)
                {
                    await Task.Delay(2000);
                    await ExecuteSlmgrCommand("/dlv");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi KMS Server: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        // THÊM PHƯƠNG THỨC GetServerKeyByEdition
        private string GetServerKeyByEdition(string edition)
        {
            return edition.ToLower() switch
            {
                "standard" => "VDYBN-27WPP-V4HQT-9VMD4-VMK7H",
                "datacenter" => "WX4NM-KYWYW-QJJR4-XV3QB-6VM33",
                "enterprise" => "489J6-VHDMP-X63PK-3K798-CPX3Y",
                _ => "VDYBN-27WPP-V4HQT-9VMD4-VMK7H"
            };
        }

        // THÊM PHƯƠNG THỨC ActivateOfficeWithKMS
        private async Task<bool> ActivateOfficeWithKMS()
        {
            try
            {
                string officePath = FindOfficePath();

                if (string.IsNullOrEmpty(officePath))
                {
                    SetStatus("Không tìm thấy Office!", InfoBarSeverity.Error);
                    return false;
                }

                string kmsServer = "kms8.msguides.com";

                Dictionary<string, string> officeKeys = new Dictionary<string, string>()
                {
                    {"Office 2021 Professional Plus", "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"},
                    {"Office 2021 Standard", "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3"},
                    {"Office 2019 Professional Plus", "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"},
                    {"Office 2019 Standard", "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK"},
                    {"Office 2016 Professional Plus", "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"},
                    {"Office 2016 Standard", "JNRGM-WHDWX-FJJG3-K47QV-DRTFM"}
                };

                string officeVersion = DetectOfficeVersion();
                string officeKey = officeKeys.ContainsKey(officeVersion) ? officeKeys[officeVersion] : "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";

                SetStatus($"Đang kích hoạt {officeVersion}...", InfoBarSeverity.Informational);

                bool step1 = await ExecuteOfficeCommand($"/inpkey:{officeKey}");
                bool step2 = await ExecuteOfficeCommand($"/sethst:{kmsServer}");
                bool step3 = await ExecuteOfficeCommand("/act");

                if (step1 && step2 && step3)
                {
                    await Task.Delay(1000);
                    await ExecuteOfficeCommand("/dstatus");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi Office KMS: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        // THÊM PHƯƠNG THỨC ActivateWithHWID
        private async Task<bool> ActivateWithHWID()
        {
            try
            {
                SetStatus("Đang kích hoạt bằng HWID...", InfoBarSeverity.Informational);

                string productName = GetWindowsProductName();
                string key = "";

                if (productName.Contains("Home") || productName.Contains("11") || productName.Contains("10"))
                {
                    key = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99";
                }
                else if (productName.Contains("Professional") || productName.Contains("Pro"))
                {
                    key = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                }
                else if (productName.Contains("Enterprise"))
                {
                    key = "NPPR9-FWDCX-D2C8J-H872K-2YT43";
                }
                else
                {
                    key = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                }

                bool step1 = await ExecuteSlmgrCommand($"/ipk {key}");
                bool step2 = await ExecuteSlmgrCommand("/ato");

                if (step1 && step2)
                {
                    await Task.Delay(2000);
                    await ExecuteSlmgrCommand("/dlv");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi HWID: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        // THÊM PHƯƠNG THỨC ActivateWithOhook
        private async Task<bool> ActivateWithOhook()
        {
            try
            {
                SetStatus("Đang chuẩn bị Ohook...", InfoBarSeverity.Informational);

                string officePath = FindOfficePath();
                if (string.IsNullOrEmpty(officePath))
                {
                    SetStatus("Không tìm thấy Office!", InfoBarSeverity.Error);
                    return false;
                }

                try
                {
                    string ohookUrl = "https://github.com/asdcorp/ohook/releases/latest/download/ohook.dll";
                    string ohookPath = Path.Combine(officePath, "ohook.dll");

                    SetStatus("Đang tải ohook.dll...", InfoBarSeverity.Warning);

                    using (var client = new System.Net.WebClient())
                    {
                        await client.DownloadFileTaskAsync(new Uri(ohookUrl), ohookPath);
                    }

                    if (!File.Exists(ohookPath))
                    {
                        SetStatus("Không thể tải ohook.dll!", InfoBarSeverity.Error);
                        return false;
                    }

                    SetStatus("Đang đăng ký ohook...", InfoBarSeverity.Warning);
                    bool registerSuccess = await RunProcessCheckSuccess("regsvr32.exe", $"/s \"{ohookPath}\"");

                    if (registerSuccess)
                    {
                        SetStatus("Ohook đã được cài đặt thành công!", InfoBarSeverity.Success);
                        await ShowInfoDialogAsync("Thành công", "Office đã được kích hoạt bằng Ohook. Khởi động lại Office để áp dụng.");
                        return true;
                    }
                }
                catch (System.Net.WebException webEx)
                {
                    SetStatus($"Không thể tải ohook: {webEx.Message}", InfoBarSeverity.Error);
                    return false;
                }

                return false;
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi Ohook: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        private async Task<bool> ExecuteOfficeCommand(string arguments)
        {
            try
            {
                string officePath = FindOfficePath();
                if (string.IsNullOrEmpty(officePath))
                {
                    SetStatus("Không tìm thấy Office!", InfoBarSeverity.Error);
                    return false;
                }

                string osppPath = Path.Combine(officePath, "ospp.vbs");

                if (!File.Exists(osppPath))
                {
                    string systemRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
                    osppPath = Path.Combine(systemRoot, "System32", "ospp.vbs");
                }

                if (!File.Exists(osppPath))
                {
                    SetStatus("Không tìm thấy ospp.vbs!", InfoBarSeverity.Error);
                    return false;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = "cscript.exe",
                    Arguments = $"//Nologo \"{osppPath}\" {arguments}",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WorkingDirectory = officePath
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    await Task.Run(() => process.WaitForExit(30000));

                    bool success = process.ExitCode == 0 &&
                                   !output.Contains("ERROR") &&
                                   !output.Contains("failed") &&
                                   (output.Contains("successful") ||
                                    output.Contains("activated") ||
                                    output.Contains("Product activation successful"));

                    if (!success)
                    {
                        string errorMsg = !string.IsNullOrEmpty(error) ? error : output;
                        SetStatus($"Office: {errorMsg.Trim()}", InfoBarSeverity.Warning);
                    }
                    else
                    {
                        SetStatus($"Office: Thành công", InfoBarSeverity.Success);
                    }

                    return success;
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi Office command: {ex.Message}", InfoBarSeverity.Error);
                return false;
            }
        }

        // THÊM PHƯƠNG THỨC RunProcessAsync
        private async Task<string> RunProcessAsync(string fileName, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    await Task.Run(() => process.WaitForExit(30000));

                    if (process.ExitCode == 0)
                    {
                        return output;
                    }
                    else
                    {
                        return $"Error: {error}";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Exception: {ex.Message}";
            }
        }

        // THÊM PHƯƠNG THỨC RunProcessCheckSuccess
        private async Task<bool> RunProcessCheckSuccess(string fileName, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    await Task.Run(() => process.WaitForExit(30000));
                    return process.ExitCode == 0;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        // THÊM PHƯƠNG THỨC FindOfficePath
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
                if (Directory.Exists(path) && File.Exists(Path.Combine(path, "winword.exe")))
                {
                    return path;
                }
            }

            return null;
        }

        // THÊM PHƯƠNG THỨC DetectOfficeVersion
        private string DetectOfficeVersion()
        {
            string officePath = FindOfficePath();
            if (string.IsNullOrEmpty(officePath)) return "Unknown";

            if (officePath.Contains("Office16")) return "Office 2021/2019";
            if (officePath.Contains("Office15")) return "Office 2013";
            if (officePath.Contains("Office14")) return "Office 2010";

            return "Office Unknown Version";
        }

        private async void BtnRemoveActivation_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = "Bạn có muốn gỡ bỏ tất cả kích hoạt KMS/Ohook?\n\nThao tác này sẽ xóa:\n- Ohook activation\n- KMS activation\n- Scheduled tasks",
                PrimaryButtonText = "OK",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await RemoveAllActivations();
            }
        }

        // THÊM PHƯƠNG THỨC RemoveAllActivations
        private async Task RemoveAllActivations()
        {
            try
            {
                ShowBusy("Đang gỡ bỏ kích hoạt...");
                SetStatus("Đang gỡ bỏ kích hoạt...", InfoBarSeverity.Warning);
                DisableButtons();

                string psCommand = "irm https://massgrave.dev/get | iex; ohook /r; kms /r";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    await Task.Run(() => process.WaitForExit());

                    if (process.ExitCode == 0)
                    {
                        SetStatus("Đã gỡ bỏ kích hoạt", InfoBarSeverity.Success);
                    }
                    else
                    {
                        SetStatus("Gỡ bỏ không thành công", InfoBarSeverity.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async void BtnInstallKey_Click(object sender, RoutedEventArgs e)
        {
            var inputBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = "Cài đặt Product Key",
                Content = "Nhập Product Key (25 ký tự):\n\nVí dụ: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX",
                PrimaryButtonText = "OK",
            };
            await inputBox.ShowDialogAsync();
            
            // Note: WPF-UI MessageBox doesn't have input, create simple window
            // AppendOutput("Chức năng cài đặt key: Vui lòng chạy lệnh thủ công:\n");
            // AppendOutput("Windows: slmgr /ipk YOUR-KEY-HERE\n");
            // AppendOutput("Office: cscript \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /inpkey:YOUR-KEY-HERE\n\n");
        }

        private async void BtnChangeEdition_Click(object sender, RoutedEventArgs e)
        {
            await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Đổi phiên bản Windows",
                Content = "Để đổi phiên bản Windows:\n\n1. Mở Settings > Update & Security > Activation\n2. Click 'Change product key'\n3. Nhập key tương ứng với phiên bản muốn chuyển.",
                PrimaryButtonText = "OK",
            }.ShowDialogAsync();
        }

        private async void BtnConvertOffice_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = "Chuyển đổi Office từ Retail/C2R sang Volume?\n\nĐiều này cần thiết để kích hoạt KMS.",
                PrimaryButtonText = "OK",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await ConvertOfficeToVolume();
            }
        }

        private async Task ConvertOfficeToVolume()
        {
            try
            {
                ShowBusy("Đang chuyển đổi Office...");
                SetStatus("Đang chuyển đổi Office...", InfoBarSeverity.Warning);
                DisableButtons();
                

                // AppendOutput("=== Chuyển đổi Office sang Volume ===\n");
                
                string psCommand = "irm https://massgrave.dev/get | iex; kms /c2v";
                
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.OutputDataReceived += (s, e) => 
                    {
                        if (!string.IsNullOrEmpty(e.Data)) { }
                    };
                    
                    process.Start();
                    process.BeginOutputReadLine();
                    await Task.Run(() => process.WaitForExit());

                    SetStatus("Chuyển đổi hoàn tất", InfoBarSeverity.Success);
                    // AppendOutput("\n=== Hoàn thành chuyển đổi ===\n");
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
                // AppendOutput($"\n[EXCEPTION] {ex.Message}\n");
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async void BtnOfficeToolPlus_Click(object sender, RoutedEventArgs e)
        {
            await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Office Tool Plus",
                Content = "Tải Office Tool Plus tại:\nhttps://otp.landian.vip/\n\nĐây là công cụ quản lý Office toàn diện:\n- Cài đặt/Gỡ Office\n- Quản lý licenses\n- Cấu hình nâng cao",
                PrimaryButtonText = "OK"
            }.ShowDialogAsync();
            
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://otp.landian.vip/redirect/download.php?type=runtime&arch=x64",
                UseShellExecute = true
            });
        }

        private async void BtnExtractOEM_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Trích xuất $OEM$ Folder",
                Content = "Tạo $OEM$ folder để cài Windows tự động kích hoạt?\n\nFolder sẽ được tạo trên Desktop.",
                PrimaryButtonText = "OK",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await ExtractOEMFolder();
            }
        }

        private async Task ExtractOEMFolder()
        {
            try
            {
                ShowBusy("Đang tạo $OEM$ folder...");
                SetStatus("Đang tạo $OEM$ folder...", InfoBarSeverity.Informational);
                DisableButtons();
                

                string psCommand = @"
                    # Extract OEM folder for auto-activation
                    $desktop = [Environment]::GetFolderPath('Desktop')
                    $oemPath = Join-Path $desktop '$OEM$\$$\Setup\Scripts'
                    New-Item -ItemType Directory -Path $oemPath -Force | Out-Null
                    
                    # Create SetupComplete.cmd for HWID activation
                    $setupCmd = @'
@echo off
title Auto-Activation Script

:: Check for internet connection
ping -n 1 8.8.8.8 >nul 2>&1
if %errorlevel% neq 0 (
    echo No internet connection, waiting...
    timeout /t 30
)

:: Run HWID activation
powershell -NoProfile -ExecutionPolicy Bypass -Command ""irm https://massgrave.dev/get | iex; hwid""

:: Delete this script after execution
del /f /q ""%~f0""
'@
                    $setupCmd | Out-File -FilePath (Join-Path (Split-Path $oemPath) 'SetupComplete.cmd') -Encoding ASCII
                    
                    Write-Host ""OEM folder created at: $desktop\$OEM$""
                    Write-Host ""Copy this folder to ISO:\sources\$OEM$ for auto-activation""
                ";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    await Task.Run(() => process.WaitForExit());

                    if (process.ExitCode == 0)
                    {
                        SetStatus("Đã tạo $OEM$ folder thành công!", InfoBarSeverity.Success);
                        await ShowInfoDialogAsync(" Thành công", "Copy folder $OEM$ vào thư mục sources\\ của Windows ISO để tự động kích hoạt khi cài đặt.");
                    }
                    else
                    {
                        SetStatus("Lỗi tạo OEM folder", InfoBarSeverity.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async void BtnTroubleshoot_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Khắc phục sự cố",
                Content = "Tự động sửa các lỗi activation và rebuild licensing?\n\nQuá trình này có thể mất vài phút.",
                PrimaryButtonText = "OK",
                CloseButtonText = "Hủy",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await TroubleshootActivation();
            }
        }

        private async Task TroubleshootActivation()
        {
            try
            {
                ShowBusy("Đang khắc phục sự cố...");
                SetStatus("Đang khắc phục sự cố...", InfoBarSeverity.Warning);
                DisableButtons();
                

                string psCommand = @"
                    Write-Host 'Troubleshooting activation issues...'
                    
                    # Stop services
                    Stop-Service sppsvc -Force
                    Stop-Service wuauserv -Force
                    
                    # Clear cache
                    Remove-Item -Path ""$env:SystemRoot\System32\spp\store\2.0\cache"" -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path ""$env:SystemRoot\ServiceProfiles\LocalService\AppData\Local\Microsoft\WSLicense"" -Recurse -Force -ErrorAction SilentlyContinue
                    
                    # Reset licensing
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /rilc
                    
                    # Restart services
                    Start-Service sppsvc
                    Start-Service wuauserv
                    
                    # Rearm
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /rearm
                    
                    Write-Host 'Troubleshooting completed. Please restart your computer.'
                ";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    await Task.Run(() => process.WaitForExit());

                    SetStatus("Hoàn thành khắc phục sự cố", InfoBarSeverity.Success);
                    await ShowInfoDialogAsync(" Hoàn thành", "Vui lòng khởi động lại máy tính để áp dụng thay đổi.");
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async void BtnResetLicensing_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Reset Licensing",
                Content = " CẢNH BÁO: Reset toàn bộ hệ thống licensing?\n\nĐiều này sẽ xóa tất cả licenses và activation hiện tại. Máy tính sẽ cần khởi động lại.",
                PrimaryButtonText = "OK",
                CloseButtonText = "Hủy",
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await ResetWindowsLicensing();
            }
        }

        private async Task ResetWindowsLicensing()
        {
            try
            {
                ShowBusy("Đang reset licensing...");
                SetStatus("Đang reset licensing...", InfoBarSeverity.Warning);
                DisableButtons();
                

                string psCommand = @"
                    Write-Host 'Resetting Windows Licensing System...'
                    
                    # Stop services
                    Stop-Service sppsvc -Force
                    Stop-Service ClipSVC -Force
                    
                    # Delete tokens and cache
                    Remove-Item -Path ""$env:SystemRoot\System32\spp\store\2.0\*"" -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path ""$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\*"" -Recurse -Force -ErrorAction SilentlyContinue
                    
                    # Reset licensing
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /upk
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /ckms
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /rilc
                    cscript //nologo $env:SystemRoot\System32\slmgr.vbs /rearm
                    
                    # Start services
                    Start-Service sppsvc
                    Start-Service ClipSVC
                    
                    Write-Host 'Licensing system has been reset. Please restart your computer.'
                ";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.Start();
                    await Task.Run(() => process.WaitForExit());

                    SetStatus("Đã reset licensing", InfoBarSeverity.Success);
                    await ShowInfoDialogAsync(" Hoàn thành", " BẮT BUỘC khởi động lại máy tính ngay!");
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
            }
            finally
            {
                HideBusy();
                EnableButtons();
            }
        }

        private async void BtnAdvancedKMS_Click(object sender, RoutedEventArgs e)
        {
            await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Cài đặt KMS nâng cao",
                Content = " Các tùy chỉnh KMS:\n\n" +
                          "• KMS Server: Tự động chọn server tối ưu\n" +
                          "• Port: 1688 (mặc định)\n" +
                          "• Auto-renewal: Tự động gia hạn mỗi 7 ngày\n" +
                          "• Task Schedule: Tạo task tự động\n\n" +
                          "Sử dụng Online KMS để kích hoạt,\n" +
                          "hệ thống sẽ tự động tạo task gia hạn.",
                PrimaryButtonText = "OK"
            }.ShowDialogAsync();
        }

        private async void BtnCheckActivation_Click(object sender, RoutedEventArgs e)
        {
            await CheckActivationStatus();
        }

        // Keep old methods for backward compatibility
        private async void BtnActivateWindows_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Windows bằng phương pháp HWID (Vĩnh viễn)", "hwid");
        }

        private async void BtnActivateOffice_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Office bằng phương pháp Ohook (Vĩnh viễn)", "ohook");
        }

        private async Task ActivateWindows()
        {
            await ActivateWithMethod("hwid");
        }

        private async Task ActivateOffice()
        {
            await ActivateWithMethod("ohook");
        }

        private async Task CheckActivationStatus()
        {
            try
            {
                ShowBusy("Đang kiểm tra trạng thái kích hoạt...");
                SetStatus("Đang kiểm tra trạng thái kích hoạt...", InfoBarSeverity.Informational);
                DisableButtons();
                
                // Windows activation
                var windowsInfo = await RunCommand("slmgr.vbs", "/xpr");
                var windowsDetails = await RunCommand("slmgr.vbs", "/dli");
                bool winActivated = windowsInfo.IndexOf("permanent", StringComparison.OrdinalIgnoreCase) >= 0
                                    || windowsInfo.IndexOf("vĩnh viễn", StringComparison.OrdinalIgnoreCase) >= 0
                                    || windowsDetails.IndexOf("Licensed", StringComparison.OrdinalIgnoreCase) >= 0;

                // Office detection + activation
                string officeStatus = "Chưa cài";
                string officeDetails = "";
                var osppPath = GetOsppVbsPath();
                if (!string.IsNullOrEmpty(osppPath) && File.Exists(osppPath))
                {
                    var officeInfo = await RunCommand("cscript.exe", $"\"{osppPath}\" /dstatus");
                    officeDetails = officeInfo;
                    if (officeInfo.IndexOf("LICENSE STATUS: ---LICENSED---", StringComparison.OrdinalIgnoreCase) >= 0
                        || officeInfo.IndexOf("LICENSED", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        officeStatus = "Đã kích hoạt";
                    }
                    else if (officeInfo.IndexOf("---NOTIFICATIONS---", StringComparison.OrdinalIgnoreCase) >= 0
                             || officeInfo.IndexOf("UNLICENSED", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        officeStatus = "Chưa kích hoạt";
                    }
                    else
                    {
                        officeStatus = "Đã cài (không rõ trạng thái)";
                    }
                }

                // Build summary
                var summary = new System.Text.StringBuilder();
                summary.AppendLine("Kết quả kiểm tra:");
                summary.AppendLine();
                summary.AppendLine($"Windows: {(winActivated ? "ĐÃ KÍCH HOẠT" : "CHƯA KÍCH HOẠT")}");
                if (!string.IsNullOrWhiteSpace(windowsInfo))
                {
                    var oneLine = windowsInfo.Replace("\r", " ").Replace("\n", " ").Trim();
                    if (oneLine.Length > 220) oneLine = oneLine.Substring(0, 220) + "...";
                    summary.AppendLine($"- /xpr: {oneLine}");
                }
                if (!string.IsNullOrWhiteSpace(windowsDetails))
                {
                    var product = ExtractLineContains(windowsDetails, new[] { "Name:", "Tên sản phẩm", "KMS", "Windows" });
                    if (!string.IsNullOrEmpty(product)) summary.AppendLine($"- Sản phẩm: {product}");
                }
                summary.AppendLine();
                summary.AppendLine($"Office: {officeStatus}");
                if (!string.IsNullOrEmpty(osppPath))
                    summary.AppendLine($"- Phát hiện OSPP: {osppPath}");
                
                // Hide overlay before dialog to avoid blocking
                HideBusy();
                await ShowInfoDialogAsync("Trạng thái kích hoạt", summary.ToString());
                SetStatus("Hoàn thành kiểm tra", InfoBarSeverity.Success);
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi: " + ex.Message, InfoBarSeverity.Error);
                // AppendOutput($"\n[EXCEPTION] {ex.Message}\n");
            }
            finally
            {
                // If not already hidden due to dialog, ensure it's hidden
                HideBusy();
                EnableButtons();
            }
        }

        private static string ExtractLineContains(string text, string[] keys)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            try
            {
                var lines = text.Replace("\r", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var k in keys)
                {
                    var line = lines.FirstOrDefault(l => l.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!string.IsNullOrWhiteSpace(line))
                        return line.Trim();
                }
            }
            catch { }
            return string.Empty;
        }

        private static string GetOsppVbsPath()
        {
            // Common Office 16 paths (C2R/Volume). Cover both x64/x86 installations.
            var candidates = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "Microsoft Office", "Office16", "OSPP.VBS"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Microsoft Office", "Office16", "OSPP.VBS")
            };
            return candidates.FirstOrDefault(File.Exists) ?? string.Empty;
        }

        private async Task<string> RunCommand(string fileName, string arguments)
        {
            var output = string.Empty;
            
            // Fix slmgr.vbs path issue - use cscript with full path
            if (fileName == "slmgr.vbs")
            {
                fileName = "cscript.exe";
                arguments = $"//nologo {Environment.GetFolderPath(Environment.SpecialFolder.System)}\\slmgr.vbs {arguments}";
            }
            
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = new Process { StartInfo = psi })
            {
                process.Start();
                output = await process.StandardOutput.ReadToEndAsync();
                await Task.Run(() => process.WaitForExit());
            }

            return output;
        }

        private void SetStatus(string text, InfoBarSeverity severity)
        {
            // Convert InfoBarSeverity to ControlAppearance for Snackbar
            var appearance = severity switch
            {
                InfoBarSeverity.Success => ControlAppearance.Success,
                InfoBarSeverity.Error => ControlAppearance.Danger,
                InfoBarSeverity.Warning => ControlAppearance.Caution,
                _ => ControlAppearance.Secondary
            };
            
            ShowSnackbar(" Thông báo", text, appearance);
        }

        private void DisableButtons()
        {
            if (BtnActivateWindowsHWID != null) BtnActivateWindowsHWID.IsEnabled = false;
            // Buttons đã được đổi tên hoặc xóa
            if (BtnActivateOfficeOhook != null) BtnActivateOfficeOhook.IsEnabled = false;
            if (BtnCheckActivation != null) BtnCheckActivation.IsEnabled = false;
            if (BtnRemoveActivation != null) BtnRemoveActivation.IsEnabled = false;
            if (BtnInstallKey != null) BtnInstallKey.IsEnabled = false;
            if (BtnChangeEdition != null) BtnChangeEdition.IsEnabled = false;
            if (BtnConvertOffice != null) BtnConvertOffice.IsEnabled = false;
            if (BtnOfficeToolPlus != null) BtnOfficeToolPlus.IsEnabled = false;
            if (BtnExtractOEM != null) BtnExtractOEM.IsEnabled = false;
            if (BtnTroubleshoot != null) BtnTroubleshoot.IsEnabled = false;
            if (BtnResetLicensing != null) BtnResetLicensing.IsEnabled = false;
            if (BtnAdvancedKMS != null) BtnAdvancedKMS.IsEnabled = false;
        }

        private void EnableButtons()
        {
            if (BtnActivateWindowsHWID != null) BtnActivateWindowsHWID.IsEnabled = true;
            // Buttons đã được đổi tên hoặc xóa
            if (BtnActivateOfficeOhook != null) BtnActivateOfficeOhook.IsEnabled = true;
            if (BtnCheckActivation != null) BtnCheckActivation.IsEnabled = true;
            if (BtnRemoveActivation != null) BtnRemoveActivation.IsEnabled = true;
            if (BtnInstallKey != null) BtnInstallKey.IsEnabled = true;
            if (BtnChangeEdition != null) BtnChangeEdition.IsEnabled = true;
            if (BtnConvertOffice != null) BtnConvertOffice.IsEnabled = true;
            if (BtnOfficeToolPlus != null) BtnOfficeToolPlus.IsEnabled = true;
            if (BtnExtractOEM != null) BtnExtractOEM.IsEnabled = true;
            if (BtnTroubleshoot != null) BtnTroubleshoot.IsEnabled = true;
            if (BtnResetLicensing != null) BtnResetLicensing.IsEnabled = true;
            if (BtnAdvancedKMS != null) BtnAdvancedKMS.IsEnabled = true;
        }

        // ===== Ghost Toolbox Features =====
        
        // Performance & Gaming
        private async void BtnDisableHPET_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("bcdedit /deletevalue useplatformclock", "Disable HPET");
        }

        private async void BtnEnableHPET_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("bcdedit /set useplatformclock true", "Enable HPET");
        }

        private async void BtnDisableStartupDelay_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize\" /v \"StartupDelayInMSec\" /t REG_DWORD /d 0 /f", "Disable Startup Delay");
        }

        private async void BtnEnableGameMode_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKCU\Software\Microsoft\GameBar"" /v ""AutoGameModeEnabled"" /t REG_DWORD /d 1 /f
reg add ""HKCU\Software\Microsoft\GameBar"" /v ""AllowAutoGameMode"" /t REG_DWORD /d 1 /f
powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
powercfg -setactive e9a42b02-d5df-448d-aa00-03f14749eb61";
            await ExecuteGhostCommand(commands, "Enable Game Mode");
        }

        private async void BtnDisableMitigations_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"" /v ""FeatureSettingsOverride"" /t REG_DWORD /d 3 /f
reg add ""HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"" /v ""FeatureSettingsOverrideMask"" /t REG_DWORD /d 3 /f
bcdedit /set hypervisorschedulertype off";
            await ExecuteGhostCommand(commands, "Disable CPU Mitigations");
        }

        private async void BtnEnableGPUScheduling_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\" /v \"HwSchMode\" /t REG_DWORD /d 2 /f", "Enable GPU Scheduling");
        }

        // Memory & Storage
        private async void BtnDisablePagefile_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wmic computersystem where name=\"%computername%\" set AutomaticManagedPagefile=False & wmic pagefileset delete", "Disable Pagefile");
        }

        private async void BtnSetPagefile_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
wmic computersystem where name=""%computername%"" set AutomaticManagedPagefile=False
wmic pagefileset where name=""C:\\pagefile.sys"" delete
wmic pagefileset create name=""C:\pagefile.sys""
wmic pagefileset where name=""C:\\pagefile.sys"" set InitialSize=4096,MaximumSize=4096";
            await ExecuteGhostCommand(commands, "Set Pagefile 4GB");
        }

        private async void BtnDisableHibernation_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powercfg -h off", "Disable Hibernation");
        }

        private async void BtnEnableHibernation_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powercfg -h on", "Enable Hibernation");
        }

        private async void BtnDisableSuperfetch_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("sc config SysMain start=disabled & sc stop SysMain", "Disable Superfetch/Sysmain");
        }

        private async void BtnDisableFastboot_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power\" /v \"HiberbootEnabled\" /t REG_DWORD /d 0 /f", "Disable Fast Startup");
        }

        private async void BtnEnableFastboot_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power\" /v \"HiberbootEnabled\" /t REG_DWORD /d 1 /f", "Enable Fast Startup");
        }

        private async void BtnDisableSleep_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powercfg -x -standby-timeout-ac 0 & powercfg -x -standby-timeout-dc 0", "Disable Sleep Mode");
        }

        private async void BtnPagefile256_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wmic pagefileset where name=\"C:\\\\pagefile.sys\" set InitialSize=256,MaximumSize=256", "Set Pagefile 256MB");
        }

        private async void BtnPagefile3GB_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wmic pagefileset where name=\"C:\\\\pagefile.sys\" set InitialSize=3072,MaximumSize=3072", "Set Pagefile 3GB");
        }

        private async void BtnPagefile8GB_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wmic pagefileset where name=\"C:\\\\pagefile.sys\" set InitialSize=8192,MaximumSize=8192", "Set Pagefile 8GB");
        }

        private async void BtnPagefile16GB_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wmic pagefileset where name=\"C:\\\\pagefile.sys\" set InitialSize=16384,MaximumSize=16384", "Set Pagefile 16GB");
        }

        private async void BtnDisableWin2077_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System\" /v \"DisableAcrylicBackgroundOnLogon\" /t REG_DWORD /d 1 /f", "Disable Windows 2077 Blur Effect");
        }

        private async void BtnStopStartup_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("taskmgr", "Open Task Manager to manage startup programs");
        }

        private async void BtnYellowTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC4FFFF00 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFFFFFF00 /f", "Yellow Theme");
        }

        private async void BtnCyanTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC400FFFF /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF00FFFF /f", "Cyan Theme");
        }

        private async void BtnBrownTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC48B4513 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF8B4513 /f", "Brown Theme");
        }

        private async void BtnGrayTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC4808080 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF808080 /f", "Gray Theme");
        }

        private async void BtnUncompactOS_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("compact /compactos:never", "Uncompact OS");
        }

        private async void BtnCleanCache_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("del /q /f /s %TEMP%\\* & del /q /f /s C:\\Windows\\Prefetch\\*", "Clean Cache & Prefetch");
        }

        private async void BtnOptimizeDelivery_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization\" /v \"DODownloadMode\" /t REG_DWORD /d 0 /f", "Disable Delivery Optimization");
        }

        private async void BtnCleanWindowsOld_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("rd /s /q C:\\Windows.old", "Delete Windows.old Folder");
        }

        private async void BtnFixStore_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wsreset.exe", "Reset Windows Store");
        }

        private async void BtnSystemCheck_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("sfc /scannow", "System File Checker");
        }

        private async void BtnResetNetwork_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("netsh winsock reset & netsh int ip reset & ipconfig /flushdns", "Reset Network Settings");
        }

        private async void BtnCleanPrefetch_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("del /q /f /s C:\\Windows\\Prefetch\\*", "Clean Prefetch");
        }

        private async void BtnCompactOS_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("compact /compactos:always", "Compact OS (LZX)");
        }

        // Windows Updates & Services
        private async void BtnStopUpdates_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseFeatureUpdatesStartTime"" /t REG_SZ /d ""2024-01-01T00:00:00Z"" /f
reg add ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseFeatureUpdatesEndTime"" /t REG_SZ /d ""2050-01-01T00:00:00Z"" /f
reg add ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseQualityUpdatesStartTime"" /t REG_SZ /d ""2024-01-01T00:00:00Z"" /f
reg add ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseQualityUpdatesEndTime"" /t REG_SZ /d ""2050-01-01T00:00:00Z"" /f
reg add ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseUpdatesExpiryTime"" /t REG_SZ /d ""2050-01-01T00:00:00Z"" /f";
            await ExecuteGhostCommand(commands, "Stop Updates Until 2050");
        }

        private async void BtnEnableUpdates_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg delete ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseFeatureUpdatesStartTime"" /f
reg delete ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseFeatureUpdatesEndTime"" /f
reg delete ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseQualityUpdatesStartTime"" /f
reg delete ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseQualityUpdatesEndTime"" /f
reg delete ""HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"" /v ""PauseUpdatesExpiryTime"" /f";
            await ExecuteGhostCommand(commands, "Enable Updates");
        }

        private async void BtnDisableCortana_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search\" /v \"AllowCortana\" /t REG_DWORD /d 0 /f", "Disable Cortana");
        }

        private async void BtnCleanTemp_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
del /q /f /s %TEMP%\*
del /q /f /s C:\Windows\Temp\*
del /q /f /s C:\Windows\Prefetch\*";
            await ExecuteGhostCommand(commands, "Clean Temp Files");
        }

        // Registry & Explorer
        private async void BtnTakeOwnership_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKCR\*\shell\runas"" /ve /d ""Take Ownership"" /f
reg add ""HKCR\*\shell\runas"" /v ""NoWorkingDirectory"" /t REG_SZ /d """" /f
reg add ""HKCR\*\shell\runas\command"" /ve /d ""cmd.exe /c takeown /f \""%1\"" && icacls \""%1\"" /grant administrators:F"" /f
reg add ""HKCR\*\shell\runas\command"" /v ""IsolatedCommand"" /d ""cmd.exe /c takeown /f \""%1\"" && icacls \""%1\"" /grant administrators:F"" /f
reg add ""HKCR\Directory\shell\runas"" /ve /d ""Take Ownership"" /f
reg add ""HKCR\Directory\shell\runas\command"" /ve /d ""cmd.exe /c takeown /f \""%1\"" /r /d y && icacls \""%1\"" /grant administrators:F /t"" /f";
            await ExecuteGhostCommand(commands, "Add Take Ownership");
        }

        private async void BtnDisableBlur_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System\" /v \"DisableAcrylicBackgroundOnLogon\" /t REG_DWORD /d 1 /f", "Disable Login Blur");
        }

        private async void BtnDisableRibbon_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Ribbon\" /v \"MinimizedStateTabletModeOff\" /t REG_DWORD /d 1 /f", "Disable Explorer Ribbon");
        }

        private async void BtnEnableRibbon_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Ribbon\" /v \"MinimizedStateTabletModeOff\" /t REG_DWORD /d 0 /f", "Enable Explorer Ribbon");
        }

        private async void BtnDarkMode_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"" /v ""AppsUseLightTheme"" /t REG_DWORD /d 0 /f
reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"" /v ""SystemUsesLightTheme"" /t REG_DWORD /d 0 /f";
            await ExecuteGhostCommand(commands, "Enable Dark Mode");
        }

        private async void BtnLightMode_Click(object sender, RoutedEventArgs e)
        {
            var commands = @"
reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"" /v ""AppsUseLightTheme"" /t REG_DWORD /d 1 /f
reg add ""HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"" /v ""SystemUsesLightTheme"" /t REG_DWORD /d 1 /f";
            await ExecuteGhostCommand(commands, "Enable Light Mode");
        }

        // System Cleanup
        private async void BtnClearEventLogs_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("for /F \"tokens=*\" %1 in ('wevtutil.exe el') DO wevtutil.exe cl \"%1\"", "Clear Event Logs");
        }

        private async void BtnCleanStore_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("wsreset", "Clean Store Cache");
        }

        private async void BtnDiskCleanup_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("cleanmgr /sagerun:1", "Disk Cleanup");
        }

        private async void BtnDefrag_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("dfrgui", "Defragment");
        }

        // Helper method for Ghost Toolbox commands
        private async Task ExecuteGhostCommand(string psScript, string featureName)
        {
            try
            {
                var confirmed = await ShowConfirmDialogAsync(
                    "Xác nhận",
                    $"Bạn có chắc muốn thực hiện: {featureName}?\n\nYêu cầu quyền Administrator."
                );
                if (!confirmed) return;

                ShowBusy($"Đang áp dụng {featureName}...");
                SetStatus($"Đang áp dụng {featureName}...", InfoBarSeverity.Informational);

                if (string.IsNullOrWhiteSpace(psScript))
                {
                    throw new ArgumentException("PowerShell script không được rỗng");
                }

                // Clean script
                psScript = psScript.Trim();
                if (psScript.StartsWith("powershell", StringComparison.OrdinalIgnoreCase))
                {
                    var match = System.Text.RegularExpressions.Regex.Match(
                        psScript,
                        @"^powershell(?:\.exe)?\s+(?:-Command|-c)\s+""(.+)""$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                        System.Text.RegularExpressions.RegexOptions.Singleline
                    );

                    if (match.Success)
                    {
                        psScript = match.Groups[1].Value.Replace("\\\"", "\"");
                    }
                }

                // Encode script to Base64 để tránh mọi vấn đề về escape
                var bytes = Encoding.Unicode.GetBytes(psScript);
                var encodedScript = Convert.ToBase64String(bytes);

                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encodedScript}",
                        UseShellExecute = true,
                        Verb = "runas",
                        CreateNoWindow = false,
                        RedirectStandardOutput = false,
                        RedirectStandardError = false
                    }
                };

                try
                {
                    process.Start();
                }
                catch (System.ComponentModel.Win32Exception win32Ex)
                {
                    if (win32Ex.NativeErrorCode == 1223)
                    {
                        SetStatus("Đã hủy bởi người dùng", InfoBarSeverity.Warning);
                        return;
                    }
                    throw;
                }

                await process.WaitForExitAsync();

                if (process.ExitCode == 0)
                {
                    SetStatus($"Thành công: {featureName}", InfoBarSeverity.Success);
                    ShowSnackbar("Thành công", featureName, ControlAppearance.Success);
                }
                else
                {
                    SetStatus($"Thất bại: {featureName} (Exit code: {process.ExitCode})", InfoBarSeverity.Error);
                    ShowSnackbar("Thất bại", $"{featureName} - Exit code: {process.ExitCode}", ControlAppearance.Danger);
                }
            }
            catch (Exception ex)
            {
                SetStatus("Lỗi Installer", InfoBarSeverity.Error);
                await ShowInfoDialogAsync("Lỗi", ex.Message);
            }
            finally
            {
                HideBusy();
            }
        }

        // ===== Ghost Toolbox Installer Features =====

        private async void BtnInstallEdge_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Microsoft Edge", "https://go.microsoft.com/fwlink/?linkid=2108834&Channel=Stable&language=en");
        }

        private async void BtnInstallFirefox_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Firefox", "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US");
        }

        private async void BtnInstallChrome_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Google Chrome", "https://dl.google.com/chrome/install/latest/chrome_installer.exe");
        }

        private async void BtnInstallBrave_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Brave", "https://laptop-updates.brave.com/latest/winx64");
        }

        private async void BtnInstallOperaGX_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Opera GX", "https://net.geo.opera.com/opera_gx/stable/windows");
        }

        private async void BtnInstallVLC_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("VLC", "https://get.videolan.org/vlc/last/win64/vlc-3.0.20-win64.exe");
        }

        private async void BtnInstall7Zip_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("7-Zip", "https://www.7-zip.org/a/7z2301-x64.exe");
        }

        private async void BtnInstallWinRAR_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("WinRAR", "https://www.rarlab.com/rar/winrar-x64-623.exe");
        }

        private async void BtnInstallNotepadPP_Click(object sender, RoutedEventArgs e)
        {
            await DownloadAndInstall("Notepad++", "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.6/npp.8.6.Installer.x64.exe");
        }

        private async void BtnDownloadUniKey_Click(object sender, RoutedEventArgs e)
        {
            await DownloadFile("UniKey 4.6 RC2", "https://www.unikey.org/assets/release/unikey46RC2-230919-win64.zip", "unikey46RC2-win64.zip");
        }

        private async void BtnDownloadEVKey_Click(object sender, RoutedEventArgs e)
        {
            await DownloadFile("EVKey", "https://github.com/lamquangminh/EVKey/releases/download/Release/EVKey.zip", "EVKey.zip");
        }

        private async Task DownloadFile(string appName, string url, string fileName)
        {
            try
            {
                var confirmed = await ShowConfirmDialogAsync($"Tải và cài đặt {appName}", $"Tải xuống và cài đặt {appName}?\n\nFile sẽ được tải và giải nén tự động.");
                if (!confirmed) return;

                ShowBusy($"Đang tải {appName}...");

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(5);

                    var fileBytes = await httpClient.GetByteArrayAsync(url);
                    var tempPath = Path.Combine(Path.GetTempPath(), fileName);
                    var extractPath = Path.Combine(Path.GetTempPath(), $"Extract_{Guid.NewGuid()}");

                    // Save file
                    await File.WriteAllBytesAsync(tempPath, fileBytes);

                    ShowBusy($"Đang giải nén {appName}...");

                    // Extract if it's a zip file
                    if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        System.IO.Compression.ZipFile.ExtractToDirectory(tempPath, extractPath);

                        // Find and execute .exe or .msi files
                        ShowBusy($"Đang cài đặt {appName}...");

                        var exeFiles = Directory.GetFiles(extractPath, "*.exe", SearchOption.AllDirectories);
                        var msiFiles = Directory.GetFiles(extractPath, "*.msi", SearchOption.AllDirectories);

                        string installFile = exeFiles.FirstOrDefault() ?? msiFiles.FirstOrDefault();

                        if (!string.IsNullOrEmpty(installFile))
                        {
                            var process = new Process
                            {
                                StartInfo = new ProcessStartInfo
                                {
                                    FileName = installFile,
                                    UseShellExecute = true,
                                    Verb = "runas",
                                    WindowStyle = ProcessWindowStyle.Normal
                                }
                            };
                            process.Start();
                            await Task.Run(() => process.WaitForExit());

                            HideBusy();
                            ShowSnackbar("Thành công", $"Đã cài đặt {appName} thành công!", ControlAppearance.Success);
                            await ShowInfoDialogAsync("Hoàn thành", $"{appName} đã được cài đặt thành công!");
                        }
                        else
                        {
                            HideBusy();
                            ShowSnackbar("Cảnh báo", "Không tìm thấy file cài đặt (.exe/.msi) trong zip", ControlAppearance.Caution);

                            // Open extracted folder
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = "explorer.exe",
                                Arguments = extractPath,
                                UseShellExecute = true
                            });
                        }
                    }
                    else
                    {
                        // Direct installation for .exe/.msi files
                        ShowBusy($"Đang cài đặt {appName}...");

                        var process = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = tempPath,
                                UseShellExecute = true,
                                Verb = "runas",
                                WindowStyle = ProcessWindowStyle.Normal
                            }
                        };
                        process.Start();
                        await Task.Run(() => process.WaitForExit());

                        HideBusy();
                        ShowSnackbar("Thành công", $"Đã cài đặt {appName} thành công!", ControlAppearance.Success);
                        await ShowInfoDialogAsync("Hoàn thành", $"{appName} đã được cài đặt thành công!");
                    }

                    // Cleanup
                    try { File.Delete(tempPath); } catch { }
                    try { Directory.Delete(extractPath, true); } catch { }
                }
            }
            catch (Exception ex)
            {
                HideBusy();
                ShowSnackbar("Lỗi", $"Lỗi: {ex.Message}", ControlAppearance.Danger);
                await ShowInfoDialogAsync("Lỗi", $"Không thể cài đặt {appName}:\n\n{ex.Message}");
            }
        }

        private async Task DownloadAndInstall(string appName, string url)
        {
            try
            {
                var confirmed = await ShowConfirmDialogAsync($"Cài đặt {appName}", $"Tải và cài đặt {appName}?");
                if (!confirmed) return;
                ShowBusy($"Đang tải {appName}...");

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(5);
                    var fileBytes = await httpClient.GetByteArrayAsync(url);
                    var tempPath = Path.Combine(Path.GetTempPath(), $"{appName.Replace(" ", "_")}_{Guid.NewGuid()}.exe");
                    await File.WriteAllBytesAsync(tempPath, fileBytes);

                    ShowBusy($"Đang cài đặt {appName}...");
                    var process = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = tempPath,
                            UseShellExecute = true,
                            Verb = "runas",
                            WindowStyle = ProcessWindowStyle.Hidden
                        }
                    };
                    process.Start();
                    await Task.Run(() => process.WaitForExit());

                    HideBusy();
                    await ShowInfoDialogAsync("Thông báo", $"Đã cài đặt xong {appName}!");
                }
            }
            catch (Exception ex)
            {
                HideBusy();
                await ShowInfoDialogAsync("Lỗi", $"Lỗi: {ex.Message}");
            }
        }

        // UWP Apps Management
        private async void BtnInstallStore_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.WindowsStore | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Install Microsoft Store");
        }

        private async void BtnRemoveStore_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *WindowsStore* | Remove-AppxPackage\"", "Remove Microsoft Store");
        }

        private async void BtnInstallXbox_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers *Xbox* | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Install Xbox Apps");
        }

        private async void BtnRemoveXbox_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *Xbox* | Remove-AppxPackage\"", "Remove Xbox Apps");
        }

        private async void BtnInstallOfficeHub_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.MicrosoftOfficeHub | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Install Office Hub");
        }

        private async void BtnInstallOneDrive_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Start-Process \\\"$env:SystemRoot\\SysWOW64\\OneDriveSetup.exe\\\"\"", "Install OneDrive");
        }

        // Calculator
        private async void BtnInstallCalculator_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.WindowsCalculator | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Máy tính");
        }

        private async void BtnRemoveCalculator_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *WindowsCalculator* | Remove-AppxPackage\"", "Gỡ bỏ Máy tính");
        }

        // Photos
        private async void BtnInstallPhotos_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.Windows.Photos | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Hình ảnh");
        }

        private async void BtnRemovePhotos_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *Photos* | Remove-AppxPackage\"", "Gỡ bỏ Hình ảnh");
        }

        // Camera
        private async void BtnInstallCamera_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.WindowsCamera | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Camera");
        }

        private async void BtnRemoveCamera_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *WindowsCamera* | Remove-AppxPackage\"", "Gỡ bỏ Camera");
        }

        // Maps
        private async void BtnInstallMaps_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.WindowsMaps | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Bản đồ");
        }

        private async void BtnRemoveMaps_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *WindowsMaps* | Remove-AppxPackage\"", "Gỡ bỏ Bản đồ");
        }

        // Weather
        private async void BtnInstallWeather_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.BingWeather | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Thời tiết");
        }

        private async void BtnRemoveWeather_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *BingWeather* | Remove-AppxPackage\"", "Gỡ bỏ Thời tiết");
        }

        // News
        private async void BtnInstallNews_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.BingNews | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Tin tức");
        }

        private async void BtnRemoveNews_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *BingNews* | Remove-AppxPackage\"", "Gỡ bỏ Tin tức");
        }

        // Solitaire
        private async void BtnInstallSolitaire_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Microsoft.MicrosoftSolitaireCollection | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Solitaire");
        }

        private async void BtnRemoveSolitaire_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *Solitaire* | Remove-AppxPackage\"", "Gỡ bỏ Solitaire");
        }

        // Clipchamp
        private async void BtnInstallClipchamp_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage -AllUsers Clipchamp.Clipchamp | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\"}\"", "Cài đặt Clipchamp");
        }

        private async void BtnRemoveClipchamp_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("powershell -Command \"Get-AppxPackage *Clipchamp* | Remove-AppxPackage\"", "Gỡ bỏ Clipchamp");
        }

        // Action Center & Personalization
        private async void BtnDisableActionCenter_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer\" /v \"DisableNotificationCenter\" /t REG_DWORD /d 1 /f", "Disable Action Center");
        }

        private async void BtnEnableActionCenter_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg delete \"HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer\" /v \"DisableNotificationCenter\" /f", "Enable Action Center");
        }

        private async void BtnDisableNotifications_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\PushNotifications\" /v \"ToastEnabled\" /t REG_DWORD /d 0 /f", "Disable Notifications");
        }

        private async void BtnEnableNotifications_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\PushNotifications\" /v \"ToastEnabled\" /t REG_DWORD /d 1 /f", "Enable Notifications");
        }

        // Windows Theme Colors
        private async void BtnPurpleTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC44595D64 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF8856D6 /f", "Purple Theme");
        }

        private async void BtnPinkTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC4E81DB4 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFFE81DB4 /f", "Pink Theme");
        }

        private async void BtnBlueTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC40078D7 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF0078D7 /f", "Blue Theme");
        }

        private async void BtnGreenTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC400CC6A /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFF00CC6A /f", "Green Theme");
        }

        private async void BtnRedTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC4E74856 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFFE74856 /f", "Red Theme");
        }

        private async void BtnOrangeTheme_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"ColorizationColor\" /t REG_DWORD /d 0xC4FF8C00 /f & reg add \"HKCU\\Software\\Microsoft\\Windows\\DWM\" /v \"AccentColor\" /t REG_DWORD /d 0xFFFF8C00 /f", "Orange Theme");
        }

        // Other Advanced Tools
        private async void BtnWindowsRecovery_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("RecoveryDrive.exe", "Windows Recovery Drive");
        }

        private async void BtnInstallDotNet_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("dism /online /enable-feature /featurename:NetFX3 /all", "Install .NET Framework 3.5");
        }

        private async void BtnInstallDirectX_Click(object sender, RoutedEventArgs e)
        {
            await Task.Run(() => Process.Start(new ProcessStartInfo { FileName = "https://www.microsoft.com/en-us/download/details.aspx?id=35", UseShellExecute = true }));
        }

        private async void BtnInstallVCRedist_Click(object sender, RoutedEventArgs e)
        {
            await Task.Run(() => Process.Start(new ProcessStartInfo { FileName = "https://aka.ms/vs/17/release/vc_redist.x64.exe", UseShellExecute = true }));
        }

        private async void BtnDiskBenchmark_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("winsat disk -drive c", "Disk Benchmark");
        }

        private async void BtnCMDColors_Click(object sender, RoutedEventArgs e)
        {
            await ShowInfoDialogAsync("CMD Colors", "CMD Color Schemes:\n\nXem thêm tại: https://github.com/mbadolato/iTerm2-Color-Schemes\n\nCác theme phổ biến:\n- Dracula\n- Monokai\n- Solarized\n- Nord\n- One Dark");
        }

        private async void BtnSoundMods_Click(object sender, RoutedEventArgs e)
        {
            await ShowInfoDialogAsync("Sound Mods", "Sound Mods:\n\n- Dolby Atmos for Headphones\n- DTS Sound Unbound\n- Windows Sonic for Headphones\n\nMở Settings > Sound > Spatial Sound để cấu hình");
        }

        private async void BtnUserManagement_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("netplwiz", "User Management");
        }

        private async void BtnWingetManager_Click(object sender, RoutedEventArgs e)
        {
            await ExecuteGhostCommand("start ms-windows-store://pdp/?ProductId=9NBLGGH4NNS1", "Windows Package Manager");
        }

        // Quick Tools Event Handlers
        private void BtnTaskManager_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "taskmgr.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Task Manager", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở Task Manager: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnCMD_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Command Prompt", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở CMD: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnPowerShell_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở PowerShell", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở PowerShell: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnControlPanel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "control.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Control Panel", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở Control Panel: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnSystemInfo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "msinfo32.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở System Information", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở System Info: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnDeviceManager_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "devmgmt.msc",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Device Manager", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở Device Manager: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnOpenRegistry_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "regedit.exe",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Registry Editor", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở Registry: {ex.Message}", ControlAppearance.Caution);
            }
        }

        private void BtnOpenServices_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "services.msc",
                    UseShellExecute = true
                });
                ShowSnackbar("Thành công", "Đã mở Services Manager", ControlAppearance.Success);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Lỗi", $"Không thể mở Services: {ex.Message}", ControlAppearance.Caution);
            }
        }
    }
}

