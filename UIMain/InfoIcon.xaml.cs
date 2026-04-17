/*
 InfoIcon is a reusable help icon control for the Settings panel.
 It resolves its HelpKey against the Language CSV, shows the translated text
 as a tooltip on hover and in a popup bubble on click.
*/

using System.Windows;
using System.Windows.Controls;
using L10n = AESCConstruct2026.Localization.Language;

namespace AESCConstruct2026.UIMain
{
    public partial class InfoIcon : UserControl
    {
        public static readonly DependencyProperty HelpKeyProperty =
            DependencyProperty.Register(
                nameof(HelpKey),
                typeof(string),
                typeof(InfoIcon),
                new PropertyMetadata(null, OnHelpKeyChanged));

        public string HelpKey
        {
            get => (string)GetValue(HelpKeyProperty);
            set => SetValue(HelpKeyProperty, value);
        }

        public InfoIcon()
        {
            InitializeComponent();
            Loaded += (_, __) => ApplyLocalization();
        }

        // Re-resolves HelpKey against the current language and updates tooltip + popup text.
        // Called from Language.LocalizeFrameworkElement so the icon refreshes on language change.
        public void ApplyLocalization()
        {
            var key = HelpKey;
            if (string.IsNullOrWhiteSpace(key))
            {
                Visibility = Visibility.Collapsed;
                return;
            }

            var raw = L10n.Translate(key) ?? string.Empty;
            var text = raw.Replace("\\n", "\n");

            InfoButton.ToolTip = text;
            HelpText.Text = text;
            Visibility = Visibility.Visible;
        }

        private static void OnHelpKeyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is InfoIcon icon && icon.IsLoaded)
                icon.ApplyLocalization();
        }

        private void InfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(HelpKey)) return;
            InfoPopup.IsOpen = !InfoPopup.IsOpen;
        }
    }
}
