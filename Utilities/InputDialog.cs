using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    public static class InputDialog
    {
        public static string Show(
            string prompt,
            string title,
            string defaultValue = "",
            Window owner = null,
            string previewImageBase64 = null)
        {
            var dlg = new Window
            {
                Title = title,
                SizeToContent = SizeToContent.WidthAndHeight,
                MinWidth = 380,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                Owner = owner,
                ShowInTaskbar = false
            };

            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // preview (optional)
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // prompt label
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // textbox
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // buttons

            int row = 0;

            BitmapImage previewBmp = TryDecodeBase64Png(previewImageBase64);
            if (previewBmp != null)
            {
                var border = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)),
                    BorderThickness = new Thickness(1),
                    Background = Brushes.White,
                    Margin = new Thickness(0, 0, 0, 10),
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                var img = new Image
                {
                    Source = previewBmp,
                    Stretch = Stretch.Uniform,
                    MaxWidth = 320,
                    MaxHeight = 220
                };
                border.Child = img;
                Grid.SetRow(border, row);
                grid.Children.Add(border);
            }
            row++;

            var label = new TextBlock
            {
                Text = prompt,
                Margin = new Thickness(0, 0, 0, 8),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(label, row++);
            grid.Children.Add(label);

            var textBox = new TextBox { Text = defaultValue ?? "", MinWidth = 320 };
            Grid.SetRow(textBox, row++);
            grid.Children.Add(textBox);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var okButton = new Button { Content = "OK", Width = 75, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var cancelButton = new Button { Content = "Cancel", Width = 75, IsCancel = true };
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            Grid.SetRow(buttons, row);
            grid.Children.Add(buttons);

            dlg.Content = grid;

            bool accepted = false;
            okButton.Click += (s, e) => { accepted = true; dlg.DialogResult = true; };

            dlg.Loaded += (s, e) =>
            {
                textBox.SelectAll();
                textBox.Focus();
            };

            dlg.ShowDialog();
            return accepted ? (textBox.Text ?? "").Trim() : null;
        }

        private static BitmapImage TryDecodeBase64Png(string base64)
        {
            if (string.IsNullOrWhiteSpace(base64)) return null;
            try
            {
                byte[] bytes = Convert.FromBase64String(base64);
                var bmp = new BitmapImage();
                using (var ms = new MemoryStream(bytes))
                {
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                    bmp.Freeze();
                }
                return bmp;
            }
            catch
            {
                return null;
            }
        }
    }
}
