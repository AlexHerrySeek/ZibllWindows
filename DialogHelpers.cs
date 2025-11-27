using System.Threading.Tasks;
using System.Windows;

namespace ZibllWindows
{
    public partial class MainWindow
    {
        private async Task ShowInfoDialogAsync(string title, string message)
        {
            // Use MessageBox instead of ContentDialog to avoid DialogHost requirement
            var dialog = new Wpf.Ui.Controls.MessageBox
            {
                Title = title,
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = message,
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    MaxWidth = 520
                },
                PrimaryButtonText = "OK",
                CloseButtonText = "Close"
            };
            await dialog.ShowDialogAsync();
        }

        private async Task<bool> ShowConfirmDialogAsync(string title, string message, string primaryText = "Có", string secondaryText = "Không")
        {
            var dialog = new Wpf.Ui.Controls.MessageBox
            {
                Title = title,
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = message,
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    MaxWidth = 520
                },
                PrimaryButtonText = primaryText,
                CloseButtonText = "Close"
            };
            var result = await dialog.ShowDialogAsync();
            return result == Wpf.Ui.Controls.MessageBoxResult.Primary;
        }
    }
}
