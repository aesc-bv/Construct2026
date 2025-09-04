using AESCConstruct25.FrameGenerator.UI;
using AESCConstruct25.Licensing;
using AESCConstruct25.UI;
using SpaceClaim.Api.V242;
using System;
using System.Drawing;
using System.Drawing.Imaging;                 // for ImageFormat.Png in GDI → WPF conversion
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;       // for ElementHost
using System.Windows.Media;                   // for ImageSource, Brushes, Transform
using System.Windows.Media.Imaging;           // for BitmapImage
using System.Windows.Threading;
using Application = SpaceClaim.Api.V242.Application;
using Image = System.Drawing.Image;
using Panel = SpaceClaim.Api.V242.Panel;
using WpfWindow = System.Windows.Window;
// Optional WPF aliases to avoid WinForms collisions
using WpfBrushes = System.Windows.Media.Brushes;
using WpfCursors = System.Windows.Input.Cursors;
using WpfControlTemplate = System.Windows.Controls.ControlTemplate;



namespace AESCConstruct25.UIMain
{
    public static class UIManager
    {
        // Tracks whether to float or dock
        private static bool _floatingMode;
        private static string _lastPanelKey;

        // Docked hosts and controls
        private static ElementHost _profileHost, _settingsHost, _plateHost, _fastenerHost,
                                  _ribCutOutHost, _customPropertiesHost, _engravingHost, _connectorHost;
        private static ProfileSelectionControl _profileControl;
        private static SettingsControl _settingsControl;
        private static PlatesControl _plateControl;
        private static FastenersControl _fastenerControl;
        private static RibCutOutControl _ribCutOutControl;
        private static CustomComponentControl _customPropertiesControl;
        private static EngravingControl _engravingControl;
        private static ConnectorControl _connectorControl;

        // Command names
        public const string ProfileCommand = "AESCConstruct25.ProfileSidebar";
        public const string SettingsCommand = "AESCConstruct25.SettingsSidebar";
        public const string PlateCommand = "AESCConstruct25.Plate";
        public const string FastenerCommand = "AESCConstruct25.Fastener";
        public const string RibCutOutCommand = "AESCConstruct25.RibCutOut";
        public const string CustomPropertiesCommand = "AESCConstruct25.CustomProperties";
        public const string EngravingCommand = "AESCConstruct25.EngravingControl";
        public const string ConnectorCommand = "AESCConstruct25.ConnectorSidebar";
        public const string DockToggleCommand = "AESCConstruct25.DockCmd";

        private static Thread _floatingThread;
        private static Dispatcher _floatingDispatcher;
        private static WpfWindow _floatingWindow;
        // Floating window chrome
        private static System.Windows.Controls.ContentControl _floatingContentHost;
        private static System.Windows.Rect? _restoreBounds;   // remembers pre-snap size/pos


        private static Command DockToggleCommandHolder;
        static Bitmap _undock = new Bitmap(new MemoryStream(Resources.Menu_Undock));
        static Bitmap _dock = new Bitmap(new MemoryStream(Resources.Menu_Dock));
        public static void RegisterAll()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            Register(
                ProfileCommand,
                Localization.Language.Translate("Ribbon.Button.FrameGenerator"),
                "Open the profile selection sidebar",
                Resources.FrameGen,
                () => Show(ProfileCommand),
                valid
            );

            Register(
                SettingsCommand,
                Localization.Language.Translate("Ribbon.Button.Settings"),
                "Open the settings sidebar",
                Resources.settings,
                () => Show(SettingsCommand),
                true
            );

            Register(
                PlateCommand,
                Localization.Language.Translate("Ribbon.Button.Plate"),
                "Open the plate-creation pane",
                Resources.InsertPlate,
                () => Show(PlateCommand),
                valid
            );

            Register(
                FastenerCommand,
                Localization.Language.Translate("Ribbon.Button.Fastener"),
                "Open the fastener insertion pane",
                Resources.Fasteners,
                () => Show(FastenerCommand),
                valid
            );

            Register(
                RibCutOutCommand,
                Localization.Language.Translate("Ribbon.Button.RibCutOut"),
                "Open the rib cut-out pane",
                Resources.ribCutout,
                () => Show(RibCutOutCommand),
                valid
            );

            Register(
                CustomPropertiesCommand,
                Localization.Language.Translate("Ribbon.Button.CustomProperties"),
                "Open the custom-properties pane",
                Resources.Custom_Properties,
                () => Show(CustomPropertiesCommand),
                valid
            );

            Register(
                EngravingCommand,
                Localization.Language.Translate("Ribbon.Button.Engraving"),
                "Open the engraving pane",
                Resources.Engraving,
                () => Show(EngravingCommand),
                valid
            );

            Register(
                ConnectorCommand,
                Localization.Language.Translate("Ribbon.Button.Connector"),
                "Open the connector pane",
                Resources.Menu_Connector,
                () => Show(ConnectorCommand),
                valid
            );

            Register(
                DockToggleCommand,
                Localization.Language.Translate("Ribbon.Button.Float"),
                "Toggle between docked panel and floating window",
                Resources.Menu_Undock,
                () => ToggleMode(),
                true
            );
        }

        public static void RefreshLicenseUI()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            SetEnabled(ProfileCommand, valid);
            SetEnabled(PlateCommand, valid);
            SetEnabled(FastenerCommand, valid);
            SetEnabled(RibCutOutCommand, valid);
            SetEnabled(CustomPropertiesCommand, valid);
            SetEnabled(EngravingCommand, valid);
            SetEnabled(ConnectorCommand, valid);

            UpdateCommandTexts();
        }

        private static void SetEnabled(string commandId, bool enabled)
        {
            var c = Command.GetCommand(commandId);
            if (c != null) c.IsEnabled = enabled;
        }

        private static void Register(string name, string text, string hint, byte[] icon, Action execute, bool enabled)
        {
            var cmd = Command.Create(name);

            if (name == DockToggleCommand)
                DockToggleCommandHolder = cmd;

            cmd.Text = text;
            cmd.Hint = hint;
            cmd.Image = LoadImage(icon);
            cmd.IsEnabled = enabled;
            cmd.Executing += (s, e) => execute();
            cmd.KeepAlive(true);
        }

        private static void ToggleMode()
        {
            _floatingMode = !_floatingMode;

            if (DockToggleCommandHolder != null)
            {
                // same icon logic as before
                DockToggleCommandHolder.Image = _floatingMode ? _dock : _undock;

                // translated label: Dock when floating, Float when docked
                DockToggleCommandHolder.Text = Localization.Language.Translate(
                    _floatingMode ? "Ribbon.Button.Float" : "Ribbon.Button.Float"
                );
            }

            if (_floatingMode)
            {
                ShowFloating(_lastPanelKey);
            }
            else
            {
                CloseFloatingWindow();
                ShowDocked(_lastPanelKey);
            }
        }

        private static void Show(string panelKey)
        {
            _lastPanelKey = panelKey;
            if (_floatingMode)
            {
                ShowFloating(panelKey);
            }
            else
            {
                ShowDocked(panelKey);
            }
        }

        private static void CloseFloatingWindow()
        {
            if (_floatingDispatcher != null)
            {
                _floatingDispatcher.BeginInvoke(new Action(() =>
                {
                    _floatingWindow?.Close();
                }));
            }
        }

        private static void ShowFloating(string key)
        {
            Command.Execute("AESC.Construct.SetMode3D");

            if (_floatingThread != null
                && _floatingThread.IsAlive
                && _floatingWindow != null)
            {
                _floatingDispatcher.BeginInvoke(new Action(() =>
                {
                    if (_floatingContentHost != null)
                        _floatingContentHost.Content = CreateControl(key);
                    else
                        _floatingWindow.Content = CreateFloatingRoot(key); // safety fallback

                    _floatingWindow.Tag = key;

                    if (_floatingWindow.WindowState == WindowState.Minimized)
                        _floatingWindow.WindowState = WindowState.Normal;

                    _floatingWindow.Activate();
                }));
                return;
            }

            // First-time creation (or after user closed the window): spin up STA thread + dispatcher
            _floatingThread = new Thread(() =>
            {
                // capture this thread’s dispatcher
                _floatingDispatcher = Dispatcher.CurrentDispatcher;

                // build, show, and run the window
                _floatingWindow = CreateFloatingWindow(key);
                _floatingWindow.Show();
                Dispatcher.Run();
            });

            _floatingThread.SetApartmentState(ApartmentState.STA);
            _floatingThread.IsBackground = true;
            _floatingThread.Start();
        }

        private static WpfWindow CreateFloatingWindow(string key)
        {
            var win = new WpfWindow
            {
                Title = "AESC Construct",
                SizeToContent = SizeToContent.WidthAndHeight,
                Content = CreateFloatingRoot(key),
                Tag = key,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            win.Closed += FloatingWindow_Closed;
            System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(win);

            return win;
        }

        // Root = [Top toolbar with 3 buttons] + [content host]
        // Root = [Top toolbar with 3 buttons] + [content host]
        private static System.Windows.FrameworkElement CreateFloatingRoot(string key)
        {
            var root = new System.Windows.Controls.DockPanel();

            // --- Toolbar (top) ---
            var toolbar = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                Margin = new Thickness(6, 6, 6, 4)
            };

            // Use the same toggle icon as the ribbon command for Left/Right
            var dockToggleIcon = GetDockToggleIconWpf();

            // For Restore, use the requested resource image: Icon_Update
            var restoreIcon = GetResourceImageWpf(Resources.Icon_Update) ?? dockToggleIcon;

            // Left (rotate 180°)
            var btnLeft = MakeIconButton(
                dockToggleIcon,
                tooltip: "Snap left (full height)",
                margin: new Thickness(0, 0, 6, 0),
                onClick: (s, e) => _floatingDispatcher.BeginInvoke(new Action(() =>
                {
                    if (_floatingWindow != null) SnapFloatingToSide(_floatingWindow, Side.Left);
                })),
                size: 22,
                rotationDegrees: 180
            );

            // Right (normal)
            var btnRight = MakeIconButton(
                dockToggleIcon,
                tooltip: "Snap right (full height)",
                margin: new Thickness(0, 0, 6, 0),
                onClick: (s, e) => _floatingDispatcher.BeginInvoke(new Action(() =>
                {
                    if (_floatingWindow != null) SnapFloatingToSide(_floatingWindow, Side.Right);
                })),
                size: 22,
                rotationDegrees: 0
            );

            // Restore (Icon_Update)
            var btnRestore = MakeIconButton(
                restoreIcon,
                tooltip: "Restore original size & position",
                margin: new Thickness(0),
                onClick: (s, e) => _floatingDispatcher.BeginInvoke(new Action(() =>
                {
                    if (_floatingWindow != null) RestoreFloatingBounds(_floatingWindow);
                })),
                size: 22,
                rotationDegrees: 0
            );

            toolbar.Children.Add(btnLeft);
            toolbar.Children.Add(btnRestore);
            toolbar.Children.Add(btnRight);

            System.Windows.Controls.DockPanel.SetDock(toolbar, System.Windows.Controls.Dock.Top);
            root.Children.Add(toolbar);

            // --- Content host (fills remaining area) ---
            _floatingContentHost = new System.Windows.Controls.ContentControl
            {
                Content = CreateControl(key),
                Margin = new Thickness(6, 0, 6, 6)
            };
            root.Children.Add(_floatingContentHost);

            return root;
        }
        private static ImageSource GetResourceImageWpf(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return null;

            try
            {
                using (var ms = new MemoryStream(bytes))
                {
                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.StreamSource = ms;
                    bi.EndInit();
                    bi.Freeze();
                    return bi;
                }
            }
            catch
            {
                return null;
            }
        }


        private enum Side { Left, Right }

        // Creates a no-chrome, image-only WPF button with optional rotation
        private static System.Windows.Controls.Button MakeIconButton(
            ImageSource icon,
            string tooltip,
            Thickness margin,
            RoutedEventHandler onClick,
            double size = 22,
            double rotationDegrees = 0
        )
        {
            var img = new System.Windows.Controls.Image
            {
                Source = icon,
                Width = size,
                Height = size,
                Stretch = Stretch.Uniform,
                SnapsToDevicePixels = true,
                RenderTransformOrigin = new System.Windows.Point(0.5, 0.5),
                RenderTransform = Math.Abs(rotationDegrees % 360) < 0.001
                    ? Transform.Identity
                    : new RotateTransform(rotationDegrees)
            };

            var btn = new System.Windows.Controls.Button
            {
                Content = img,
                ToolTip = tooltip,
                Margin = margin,
                Padding = new Thickness(0),
                Background = WpfBrushes.Transparent,
                BorderBrush = null,
                BorderThickness = new Thickness(0),
                FocusVisualStyle = null,
                Focusable = false,
                Cursor = WpfCursors.Hand,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                VerticalAlignment = System.Windows.VerticalAlignment.Top
            };

            // Strip default WPF button chrome
            var presenter = new FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
            btn.Template = new WpfControlTemplate(typeof(System.Windows.Controls.Button))
            {
                VisualTree = presenter
            };

            if (onClick != null) btn.Click += onClick;
            return btn;
        }

        // Return the same icon used by the Dock/Undock ribbon command as a WPF ImageSource.
        // Tries Command.Image first; falls back to our cached bitmaps; finally to resource bytes.
        private static ImageSource GetDockToggleIconWpf()
        {
            // 1) Try to pull the current Command image (System.Drawing.Image) and convert to WPF
            try
            {
                if (DockToggleCommandHolder?.Image is Image gdi)
                {
                    var src = ToImageSource(gdi);
                    if (src != null) return src;
                }
            }
            catch { /* ignore */ }

            // 2) Fall back to our cached bitmaps (_undock / _dock)
            try
            {
                var gdi = _floatingMode ? _dock : _undock; // mirrors the ribbon toggle state
                var src = ToImageSource(gdi);
                if (src != null) return src;
            }
            catch { /* ignore */ }

            // 3) Last resort: reconstruct from resource bytes
            try
            {
                var bytes = _floatingMode ? Resources.Menu_Dock : Resources.Menu_Undock;
                using (var ms = new MemoryStream(bytes))
                using (var bmp = new Bitmap(ms))
                {
                    return ToImageSource(bmp);
                }
            }
            catch { /* ignore */ }

            return null;
        }

        // Convert System.Drawing.Image → WPF ImageSource (frozen, safe to reuse)
        private static ImageSource ToImageSource(Image gdiImage)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    gdiImage.Save(ms, ImageFormat.Png);
                    ms.Position = 0;

                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.CacheOption = BitmapCacheOption.OnLoad;
                    bi.StreamSource = ms;
                    bi.EndInit();
                    bi.Freeze(); // cross-thread safe
                    return bi;
                }
            }
            catch
            {
                return null;
            }
        }


        private static void SnapFloatingToSide(WpfWindow win, Side side)
        {
            // Save restore bounds (once) before we change anything
            if (_restoreBounds == null)
            {
                double w = double.IsNaN(win.Width) ? (win.ActualWidth > 0 ? win.ActualWidth : 420) : win.Width;
                double h = double.IsNaN(win.Height) ? (win.ActualHeight > 0 ? win.ActualHeight : 600) : win.Height;
                _restoreBounds = new System.Windows.Rect(win.Left, win.Top, w, h);
            }

            // Figure out the current screen’s working area (taskbar-aware)
            var hwnd = new System.Windows.Interop.WindowInteropHelper(win).Handle;
            var screen = (hwnd != IntPtr.Zero)
                ? System.Windows.Forms.Screen.FromHandle(hwnd)
                : System.Windows.Forms.Screen.PrimaryScreen;

            var wa = screen.WorkingArea; // in device pixels

            // Convert to WPF DIPs (handles DPI scaling correctly)
            var source = System.Windows.PresentationSource.FromVisual(win);
            var m = source?.CompositionTarget?.TransformFromDevice ?? System.Windows.Media.Matrix.Identity;
            var tl = m.Transform(new System.Windows.Point(wa.Left, wa.Top));
            var br = m.Transform(new System.Windows.Point(wa.Right, wa.Bottom));
            var wr = new System.Windows.Rect(tl, br);

            // Freeze width to the current window width (or minimum), grow to full-height
            double targetWidth = (win.ActualWidth > 0 ? win.ActualWidth : (double.IsNaN(win.Width) ? 420 : win.Width));
            if (targetWidth < 280) targetWidth = 280;

            win.SizeToContent = SizeToContent.Manual;
            win.Height = wr.Height;
            win.Top = wr.Top;

            win.Width = targetWidth;
            win.Left = (side == Side.Left) ? wr.Left : (wr.Right - targetWidth);

            if (win.WindowState == WindowState.Minimized) win.WindowState = WindowState.Normal;
            win.Activate();
        }

        private static void RestoreFloatingBounds(WpfWindow win)
        {
            if (_restoreBounds is System.Windows.Rect r)
            {
                win.SizeToContent = SizeToContent.Manual;
                win.Left = r.Left;
                win.Top = r.Top;
                win.Width = r.Width;
                win.Height = r.Height;
                _restoreBounds = null;
            }

            if (win.WindowState == WindowState.Minimized) win.WindowState = WindowState.Normal;
            win.Activate();
        }

        private static void FloatingWindow_Closed(object sender, EventArgs e)
        {
            var win = (WpfWindow)sender;
            win.Closed -= FloatingWindow_Closed;
            _floatingWindow = null;
            _floatingContentHost = null;
            _restoreBounds = null;
        }

        private static void LogAndEnableWpfTextBoxes(DependencyObject parent)
        {
            if (parent == null) return;
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);

                // Fully qualify WPF TextBox to avoid ambiguity
                if (child is System.Windows.Controls.TextBox wpfTb)
                {
                    wpfTb.IsEnabled = true;
                    wpfTb.IsReadOnly = false;
                }

                // Recurse into children
                LogAndEnableWpfTextBoxes(child);
            }
        }

        private static void ShowDocked(string key)
        {
            Command.Execute("AESC.Construct.SetMode3D");
            switch (key)
            {
                case ProfileCommand:
                    EnsureProfile(); Application.AddPanelContent(Command.GetCommand(key), _profileHost, Panel.Options); break;
                case SettingsCommand:
                    EnsureSettings(); Application.AddPanelContent(Command.GetCommand(key), _settingsHost, Panel.Options); break;
                case PlateCommand:
                    EnsurePlate(); Application.AddPanelContent(Command.GetCommand(key), _plateHost, Panel.Options); break;
                case FastenerCommand:
                    EnsureFastener(); Application.AddPanelContent(Command.GetCommand(key), _fastenerHost, Panel.Options); break;
                case RibCutOutCommand:
                    EnsureRibCutOut(); Application.AddPanelContent(Command.GetCommand(key), _ribCutOutHost, Panel.Options); break;
                case CustomPropertiesCommand:
                    EnsureCustomProperties(); Application.AddPanelContent(Command.GetCommand(key), _customPropertiesHost, Panel.Options); break;
                case EngravingCommand:
                    EnsureEngraving(); Application.AddPanelContent(Command.GetCommand(key), _engravingHost, Panel.Options); break;
                case ConnectorCommand:
                    EnsureConnector(); Application.AddPanelContent(Command.GetCommand(key), _connectorHost, Panel.Options); break;
            }
        }

        /// <summary>
        /// Always new instance for floating to prevent parent conflicts.
        /// </summary>
        private static System.Windows.Controls.UserControl CreateControl(string key)
        {
            switch (key)
            {
                case ProfileCommand: return new ProfileSelectionControl();
                case SettingsCommand: return new SettingsControl();
                case PlateCommand: return new PlatesControl();
                case FastenerCommand: return new FastenersControl();
                case RibCutOutCommand: return new RibCutOutControl();
                case CustomPropertiesCommand: return new CustomComponentControl();
                case EngravingCommand: return new EngravingControl();
                case ConnectorCommand: return new ConnectorControl();
                default: return null;
            }
        }

        public static void UpdateCommandTexts()
        {
            // Sidebar buttons
            var c = Command.GetCommand(ProfileCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.FrameGenerator");

            c = Command.GetCommand(SettingsCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Settings");

            c = Command.GetCommand(PlateCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Plate");

            c = Command.GetCommand(FastenerCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Fastener");

            c = Command.GetCommand(RibCutOutCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.RibCutOut");

            c = Command.GetCommand(CustomPropertiesCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.CustomProperties");

            c = Command.GetCommand(EngravingCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Engraving");

            c = Command.GetCommand(ConnectorCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Connector");

            // Float/Dock toggle
            if (DockToggleCommandHolder != null)
            {
                DockToggleCommandHolder.Text = Localization.Language.Translate("Ribbon.Button.Float");
            }
        }

        // Lazy initializers for docked panels
        private static void EnsureProfile() { if (_profileControl == null) { _profileControl = new ProfileSelectionControl(); _profileHost = new ElementHost { Dock = DockStyle.Fill, Child = _profileControl }; } }
        private static void EnsureSettings() { if (_settingsControl == null) { _settingsControl = new SettingsControl(); _settingsHost = new ElementHost { Dock = DockStyle.Fill, Child = _settingsControl }; } }
        private static void EnsurePlate() { if (_plateControl == null) { _plateControl = new PlatesControl(); _plateHost = new ElementHost { Dock = DockStyle.Fill, Child = _plateControl }; } }
        private static void EnsureFastener() { if (_fastenerControl == null) { _fastenerControl = new FastenersControl(); _fastenerHost = new ElementHost { Dock = DockStyle.Fill, Child = _fastenerControl }; } }
        private static void EnsureRibCutOut() { if (_ribCutOutControl == null) { _ribCutOutControl = new RibCutOutControl(); _ribCutOutHost = new ElementHost { Dock = DockStyle.Fill, Child = _ribCutOutControl }; } }
        private static void EnsureCustomProperties() { if (_customPropertiesControl == null) { _customPropertiesControl = new CustomComponentControl(); _customPropertiesHost = new ElementHost { Dock = DockStyle.Fill, Child = _customPropertiesControl }; } }
        private static void EnsureEngraving() { if (_engravingControl == null) { _engravingControl = new EngravingControl(); _engravingHost = new ElementHost { Dock = DockStyle.Fill, Child = _engravingControl }; } }
        private static void EnsureConnector() { if (_connectorControl == null) { _connectorControl = new ConnectorControl(); _connectorHost = new ElementHost { Dock = DockStyle.Fill, Child = _connectorControl }; } }

        private static Image LoadImage(byte[] bytes)
        {
            try { return new Bitmap(new MemoryStream(bytes)); }
            catch { return null; }
        }
    }
}
