using System;
using System.IO;
using System.Text.Json;

namespace ZibllWindows
{
    public class AppSettings
    {
        public string MyCrushName = "Hà Quỳnh Anh";
        public string Theme { get; set; } = "Dark";
        public string AccentColor { get; set; } = "Mặc định";
        public double Opacity { get; set; } = 1.0;
        public string FontFamily { get; set; } = "Segoe UI";
        public double FontSize { get; set; } = 14;
        public bool UseBackgroundImage { get; set; } = false;
        public string BackgroundImagePath { get; set; } = "";
        public string BackgroundTheme { get; set; } = "Không";
        public double BackgroundOpacity { get; set; } = 0.3;
        public double WindowWidth { get; set; } = 1100;
        public double WindowHeight { get; set; } = 500;
        public int WindowState { get; set; } = 0; // 0: Normal, 1: Minimized, 2: Maximized
        public bool UseAnimations { get; set; } = true;
        public double BorderRadius { get; set; } = 8;
        public bool CompactMode { get; set; } = false;
        public string BackdropType { get; set; } = "None";

        private static readonly string SettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ZibllWindows",
            "settings.json"
        );

        private static readonly string BackgroundImageFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ZibllWindows",
            "Backgrounds"
        );

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                string? directory = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        public static string SaveBackgroundImage(string sourceImagePath)
        {
            try
            {
                if (!Directory.Exists(BackgroundImageFolder))
                {
                    Directory.CreateDirectory(BackgroundImageFolder);
                }

                string fileName = $"background_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(sourceImagePath)}";
                string destinationPath = Path.Combine(BackgroundImageFolder, fileName);

                // Delete old background images
                foreach (var oldFile in Directory.GetFiles(BackgroundImageFolder, "background_*"))
                {
                    try { File.Delete(oldFile); } catch { }
                }

                // Copy new image
                File.Copy(sourceImagePath, destinationPath, true);
                return destinationPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving background image: {ex.Message}");
                return sourceImagePath; // Return original path if save fails
            }
        }
    }
}
