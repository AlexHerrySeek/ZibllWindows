using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace ZibllWindows
{
    public partial class ProductKeyDialog : Wpf.Ui.Controls.FluentWindow
    {
        public string ProductKey { get; private set; } = string.Empty;

        public ProductKeyDialog()
        {
            InitializeComponent();
            KeyTextBox.Focus();
            KeyTextBox.SelectAll();
        }

        private void KeyTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            string text = KeyTextBox.Text.ToUpper();
            
            // Remove any non-alphanumeric characters except hyphens
            text = Regex.Replace(text, "[^A-Z0-9-]", "");
            
            // Auto-format with hyphens (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
            if (text.Length > 0 && !text.Contains("-"))
            {
                // Add hyphens every 5 characters
                string formatted = "";
                for (int i = 0; i < text.Length; i++)
                {
                    if (i > 0 && i % 5 == 0)
                        formatted += "-";
                    formatted += text[i];
                }
                text = formatted;
            }

            // Update text if changed
            if (KeyTextBox.Text != text)
            {
                int caretIndex = KeyTextBox.CaretIndex;
                KeyTextBox.Text = text;
                KeyTextBox.CaretIndex = Math.Min(caretIndex, text.Length);
            }

            // Validate: should be 29 characters (25 alphanumeric + 4 hyphens)
            string keyWithoutHyphens = text.Replace("-", "");
            bool isValid = keyWithoutHyphens.Length == 25 && text.Length == 29;
            
            OkButton.IsEnabled = isValid;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ProductKey = KeyTextBox.Text.Replace("-", "").ToUpper();
            
            if (ProductKey.Length == 25)
            {
                DialogResult = true;
                Close();
            }
            else
            {
                System.Windows.MessageBox.Show("Product Key phải có 25 ký tự!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        protected override void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter && OkButton.IsEnabled)
            {
                OkButton_Click(OkButton, new RoutedEventArgs());
            }
            else if (e.Key == Key.Escape)
            {
                CancelButton_Click(null, new RoutedEventArgs());
            }
            base.OnKeyDown(e);
        }
    }
}
