using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
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
            
            // Load settings
            settings = AppSettings.Load();
            ApplySettings();
            
            CheckAdminPrivileges();
            
            // Navigate to home page by default
            NavigateToPage("home");
            
            // Show intro and check for updates
            Loaded += async (s, e) => await ShowIntroAndCheckUpdate();
            
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
            settings?.Save();
        }

        private async Task ShowIntroAndCheckUpdate()
        {
            try
            {
                // Wait for intro animations to play (1.5 seconds)
                await Task.Delay(1500);
                
                // Show intro for 1 second
                IntroStatusText.Text = "Đang khởi động...";
                await Task.Delay(1000);
                
                // Check for updates
                IntroStatusText.Text = "Kiểm tra cập nhật...";
                await Task.Delay(500);
                
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(10);
                    
                    string versionUrl = "https://raw.githubusercontent.com/AlexHerrySeek/ZibllWindows/refs/heads/main/backend/version";
                    string latestVersion = await httpClient.GetStringAsync(versionUrl);
                    latestVersion = latestVersion.Trim();

                    if (latestVersion != version)
                    {
                        IntroStatusText.Text = $"Phát hiện phiên bản mới: {latestVersion}";
                        await Task.Delay(1000);
                        
                        // Fade out intro with animation
                        await FadeOutIntro();
                        
                        var result = System.Windows.MessageBox.Show(
                            $"Phiên bản hiện tại: {version}\nPhiên bản mới: {latestVersion}\n\nBạn có muốn tải về và cập nhật?",
                            "Cập nhật có sẵn",
                            System.Windows.MessageBoxButton.YesNo,
                            System.Windows.MessageBoxImage.Information);

                        if (result == System.Windows.MessageBoxResult.Yes)
                        {
                            await DownloadAndUpdate();
                            return;
                        }
                    }
                    else
                    {
                        IntroStatusText.Text = "Bạn đang sử dụng phiên bản mới nhất";
                        await Task.Delay(1000);
                    }
                }
                
                // Fade out intro and show welcome message
                await FadeOutIntro();
                ShowSnackbar("Sẵn sàng", "Ứng dụng đã sẵn sàng để kích hoạt Windows và Office", ControlAppearance.Success, 4000);
            }
            catch (Exception ex)
            {
                IntroStatusText.Text = "Lỗi khi kiểm tra cập nhật";
                await Task.Delay(1500);
                await FadeOutIntro();
                
                ShowSnackbar("Lỗi", $"Không thể kiểm tra cập nhật: {ex.Message}\n\nỨng dụng sẽ đóng sau 3 giây...", ControlAppearance.Danger, 3000);
                await Task.Delay(3000);
                System.Windows.Application.Current.Shutdown();
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
            NavigateToPage("windows");
        }

        private void NavigateToOffice_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage("office");
        }

        private void NavigateToPage(string pageTag)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToPage called with: '{pageTag}'");
            
            Dispatcher.Invoke(() =>
            {
                // Hide all pages
                HomePage.Visibility = Visibility.Collapsed;
                WindowsPage.Visibility = Visibility.Collapsed;
                OfficePage.Visibility = Visibility.Collapsed;
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
                    case "windows":
                        WindowsPage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing WindowsPage");
                        break;
                    case "office":
                        OfficePage.Visibility = Visibility.Visible;
                        System.Diagnostics.Debug.WriteLine("Showing OfficePage");
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

            var backdropType = selectedItem.Content.ToString() ?? "None";
            
            try
            {
                switch (backdropType)
                {
                    case "None":
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
                // ...existing code...
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
            await ShowConfirmAndActivate("Windows bằng phương pháp HWID (Vĩnh viễn)", "hwid");
        }

        private async void BtnActivateWindowsKMS38_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Windows bằng phương pháp KMS38 (Đến 2038)", "kms38");
        }

        private async void BtnActivateWindowsOnlineKMS_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Windows bằng Online KMS (180 ngày)", "kms");
        }

        // Office Activation Methods
        private async void BtnActivateOfficeOhook_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Office bằng phương pháp Ohook (Vĩnh viễn)", "ohook");
        }

        private async void BtnActivateOfficeKMS_Click(object sender, RoutedEventArgs e)
        {
            await ShowConfirmAndActivate("Office bằng Online KMS (180 ngày)", "kms_office");
        }

        private async Task ShowConfirmAndActivate(string message, string method)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = $"Bạn có muốn kích hoạt {message}?",
                PrimaryButtonText = "OK",
                // ...existing code...
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
                // LoadingProgress is kept for Home page; BusyOverlay blocks globally
                
                string psCommand = method switch
                {
                    "hwid" => "irm https://massgrave.dev/get | iex; hwid",
                    "kms38" => "irm https://massgrave.dev/get | iex; kms38",
                    "kms" => "irm https://massgrave.dev/get | iex; kms",
                    "ohook" => "irm https://massgrave.dev/get | iex; ohook",
                    "kms_office" => "irm https://massgrave.dev/get | iex; kms /o",
                    _ => "irm https://massgrave.dev/get | iex"
                };
                
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
                        SetStatus("Kích hoạt thành công!", InfoBarSeverity.Success);
                        await ShowInfoDialogAsync(" Kích hoạt thành công", "Quá trình kích hoạt đã hoàn tất.");
                    }
                    else
                    {
                        SetStatus("Kích hoạt thất bại!", InfoBarSeverity.Error);
                        await ShowInfoDialogAsync(" Kích hoạt thất bại", "Có lỗi xảy ra trong quá trình kích hoạt.");
                    }
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

        // Additional Tools
        private async void BtnRemoveActivation_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = "Bạn có muốn gỡ bỏ tất cả kích hoạt KMS/Ohook?\n\nThao tác này sẽ xóa:\n- Ohook activation\n- KMS activation\n- Scheduled tasks",
                PrimaryButtonText = "OK",
                // ...existing code...
            }.ShowDialogAsync();

            if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
            {
                await RemoveAllActivations();
            }
        }

        private async Task RemoveAllActivations()
        {
            try
            {
                ShowBusy("Đang gỡ bỏ kích hoạt...");
                SetStatus("Đang gỡ bỏ kích hoạt...", InfoBarSeverity.Warning);
                DisableButtons();
                

                // AppendOutput("=== Bắt đầu gỡ bỏ kích hoạt ===\n");
                
                // Remove Ohook and KMS
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

                    SetStatus("Đã gỡ bỏ kích hoạt", InfoBarSeverity.Success);
                    // AppendOutput("\n=== Hoàn thành gỡ bỏ ===\n");
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

        private async void BtnInstallKey_Click(object sender, RoutedEventArgs e)
        {
            var inputBox = new Wpf.Ui.Controls.MessageBox
            {
                Title = "Cài đặt Product Key",
                Content = "Nhập Product Key (25 ký tự):\n\nVí dụ: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX",
                PrimaryButtonText = "OK",
                // ...existing code...
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
                // ...existing code...
            }.ShowDialogAsync();
        }

        private async void BtnConvertOffice_Click(object sender, RoutedEventArgs e)
        {
            var result = await new Wpf.Ui.Controls.MessageBox
            {
                Title = "Xác nhận",
                Content = "Chuyển đổi Office từ Retail/C2R sang Volume?\n\nĐiều này cần thiết để kích hoạt KMS.",
                PrimaryButtonText = "OK",
                // ...existing code...
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
                // ...existing code...
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
                // ...existing code...
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
                // ...existing code...
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
            BtnActivateWindowsHWID.IsEnabled = false;
            BtnActivateWindowsKMS38.IsEnabled = false;
            BtnActivateWindowsOnlineKMS.IsEnabled = false;
            BtnActivateOfficeOhook.IsEnabled = false;
            BtnActivateOfficeKMS.IsEnabled = false;
            BtnCheckActivation.IsEnabled = false;
            BtnRemoveActivation.IsEnabled = false;
            BtnInstallKey.IsEnabled = false;
            BtnChangeEdition.IsEnabled = false;
            BtnConvertOffice.IsEnabled = false;
            BtnOfficeToolPlus.IsEnabled = false;
            BtnExtractOEM.IsEnabled = false;
            BtnTroubleshoot.IsEnabled = false;
            BtnResetLicensing.IsEnabled = false;
            BtnAdvancedKMS.IsEnabled = false;
        }

        private void EnableButtons()
        {
            BtnActivateWindowsHWID.IsEnabled = true;
            BtnActivateWindowsKMS38.IsEnabled = true;
            BtnActivateWindowsOnlineKMS.IsEnabled = true;
            BtnActivateOfficeOhook.IsEnabled = true;
            BtnActivateOfficeKMS.IsEnabled = true;
            BtnCheckActivation.IsEnabled = true;
            BtnRemoveActivation.IsEnabled = true;
            BtnInstallKey.IsEnabled = true;
            BtnChangeEdition.IsEnabled = true;
            BtnConvertOffice.IsEnabled = true;
            BtnOfficeToolPlus.IsEnabled = true;
            BtnExtractOEM.IsEnabled = true;
            BtnTroubleshoot.IsEnabled = true;
            BtnResetLicensing.IsEnabled = true;
            BtnAdvancedKMS.IsEnabled = true;
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
        private async Task ExecuteGhostCommand(string command, string featureName)
        {
            try
            {
                var confirmed = await ShowConfirmDialogAsync(featureName, $"Bạn có chắc chắn muốn thực hiện: {featureName}?\n\nLưu ý: Một số thay đổi có thể yêu cầu khởi động lại.");
                if (!confirmed)
                {
                    ShowSnackbar("Đã hủy", $"Đã hủy thao tác: {featureName}", ControlAppearance.Secondary);
                    return;
                }
                
                ShowBusy($"Đang thực hiện: {featureName}...");
                SetStatus($"Đang thực hiện: {featureName}...", InfoBarSeverity.Informational);

                var tempFile = Path.Combine(Path.GetTempPath(), $"ghost_{Guid.NewGuid()}.cmd");
                await File.WriteAllTextAsync(tempFile, $"@echo off\n{command}\necho.\necho Hoan tat: {featureName}\npause");

                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c \"{tempFile}\"",
                        Verb = "runas",
                        UseShellExecute = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    }
                };

                process.Start();
                await process.WaitForExitAsync();

                try { File.Delete(tempFile); } catch { }

                if (process.ExitCode == 0)
                {
                    SetStatus($"Hoàn thành: {featureName}", InfoBarSeverity.Success);
                    ShowSnackbar("Thành công", $"{featureName} đã hoàn tất!", ControlAppearance.Success);
                    await ShowInfoDialogAsync("Thành công", $"{featureName} đã hoàn tất thành công!");
                }
                else
                {
                    SetStatus($"Lỗi: {featureName}", InfoBarSeverity.Warning);
                    ShowSnackbar("Cảnh báo", $"{featureName} hoàn tất với mã lỗi: {process.ExitCode}", ControlAppearance.Caution);
                    await ShowInfoDialogAsync("Cảnh báo", $"{featureName} đã hoàn tất nhưng có mã lỗi: {process.ExitCode}\n\nVui lòng kiểm tra kết quả.");
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // User cancelled UAC prompt
                SetStatus("Đã hủy quyền Admin", InfoBarSeverity.Warning);
                ShowSnackbar("Đã hủy", "Bạn đã từ chối quyền Admin", ControlAppearance.Secondary);
            }
            catch (Exception ex)
            {
                SetStatus($"Lỗi: {ex.Message}", InfoBarSeverity.Error);
                ShowSnackbar("Lỗi", $"Lỗi: {ex.Message}", ControlAppearance.Danger);
                await ShowInfoDialogAsync("Lỗi", $"Lỗi khi thực hiện {featureName}:\n\n{ex.Message}");
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
                var confirmed = await ShowConfirmDialogAsync($"Tải {appName}", $"Tải xuống {appName}?\n\nFile sẽ được lưu vào Desktop.");
                if (!confirmed) return;

                ShowBusy($"Đang tải {appName}...");
                
                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromMinutes(5);
                    
                    var fileBytes = await httpClient.GetByteArrayAsync(url);
                    var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    var filePath = Path.Combine(desktopPath, fileName);
                    
                    File.WriteAllBytes(filePath, fileBytes);
                    
                    HideBusy();
                    ShowSnackbar("Thành công", $"Đã tải {appName} về Desktop", ControlAppearance.Success);
                    
                    // Open folder
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"/select,\"{filePath}\"",
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                HideBusy();
                ShowSnackbar("Lỗi", $"Không thể tải {appName}: {ex.Message}", ControlAppearance.Danger);
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

