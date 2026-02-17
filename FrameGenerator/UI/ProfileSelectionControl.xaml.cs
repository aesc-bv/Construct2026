/*
 ProfileSelectionControl is the main WPF UI for the FrameGenerator add-in.

 It lets the user:
 - Choose a built-in profile (Rectangular, Circular, H, L, U, T) with CSV-driven sizes.
 - Import/convert DXF profiles, save them, and select user-defined profiles (with preview images).
 - Configure placement offsets, hollow/solid options, rotation and BOM update flags.
 - Generate profile solids via ExtrudeProfileCommand and optionally rotate new components.
 - Configure and execute joints (miter, straight, T, cutout, trim) via ExecuteJointCommand.
 - Delete generated Construct profile components and restore their original driving curves.
*/

using AESCConstruct2026.Commands;
using AESCConstruct2026.FrameGenerator.Commands;
using AESCConstruct2026.FrameGenerator.Utilities;  // alias our DXFProfile class
using SpaceClaim.Api.V242;                        // for SpaceClaim API (Document, Window.ActiveWindow)
using SpaceClaim.Api.V242.Geometry;               // for ITrimmedCurve, Point, etc.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;                             // for WPF Window, MessageBox, etc.
using System.Windows.Controls;                    // for WPF UserControl, TextBox, RadioButton, etc.
using System.Windows.Media;
using System.Windows.Media.Imaging;               // if you ever need WPF BitmapImage
using Application = SpaceClaim.Api.V242.Application;
using Document = SpaceClaim.Api.V242.Document;
using DXFProfile = AESCConstruct2026.FrameGenerator.Utilities.DXFProfile;
using Image = System.Windows.Controls.Image;
using Matrix = SpaceClaim.Api.V242.Geometry.Matrix;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Orientation = System.Windows.Controls.Orientation;
using Path = System.IO.Path;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Settings = AESCConstruct2026.Properties.Settings;
using UserControl = System.Windows.Controls.UserControl;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026.FrameGenerator.UI
{
    public partial class ProfileSelectionControl : UserControl
    {
        private string selectedProfile = "";
        private string selectedProfileString = "";
        private string selectedProfileImage = "";

        private List<string> csvFieldNames = new List<string>();
        private List<string[]> csvDataRows = new List<string[]>();
        private List<string> csvRowNames = new List<string>();
        private Dictionary<string, TextBox> inputFieldMap
            = new Dictionary<string, TextBox>();
        private string selectedDXFPath = "";
        private List<ITrimmedCurve> dxfContours = null;

        private readonly Dictionary<string, string> _lastSizeByProfile = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public double rotationAngle = 0.0;

        // Initializes the UI control, applies localization, loads user profiles and wires event handlers.
        public ProfileSelectionControl()
        {
            try
            {
                InitializeComponent();
                Localization.Language.Translate("ConstructGroup");
                LocalizeUI();
                LoadUserProfiles();
                WireProfileButtonHandlers();
                WireJointButtonHandlers();
                UpdateGenerateButtons();
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to initialize ProfileSelectionControl:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Returns the last rotation angle entered in the UI for use by external callers.
        public double GetRotationAngle()
        {
            return rotationAngle;
        }

        // Applies localization to all framework elements in this control.
        private void LocalizeUI()
        {
            Localization.Language.LocalizeFrameworkElement(this);
        }

        // Subscribes built-in profile radio buttons to update the generate button state.
        private void WireProfileButtonHandlers()
        {
            foreach (var rb in ProfileGroupContainer.Children.OfType<RadioButton>())
            {
                rb.Checked += (s, e) => UpdateGenerateButtons();
                rb.Unchecked += (s, e) => UpdateGenerateButtons();
            }
        }

        // Subscribes joint type radio buttons to update the joint generate button state.
        private void WireJointButtonHandlers()
        {
            foreach (var rb in JointGroupContainer.Children.OfType<RadioButton>())
            {
                rb.Checked += (s, e) => UpdateGenerateButtons();
                rb.Unchecked += (s, e) => UpdateGenerateButtons();
            }
        }

        // Enables or disables generate buttons and updates their icons and text styles based on selection.
        private void UpdateGenerateButtons()
        {
            bool profileSelected =
                 ProfileGroupContainer.Children.OfType<RadioButton>().Any(rb => rb.IsChecked == true) ||
                 UserProfilesGrid.Children.OfType<Grid>()
                     .SelectMany(g => g.Children.OfType<RadioButton>())
                     .Any(rb => rb.IsChecked == true);

            GenerateButton.IsEnabled = profileSelected;

            // Swap image
            if (GenerateProfileButtonIcon != null)
            {
                var uriString = profileSelected
                    ? "/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Generate_Active.png"
                    : "/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Generate.png";

                GenerateProfileButtonIcon.Source =
                    new BitmapImage(new Uri(uriString, UriKind.RelativeOrAbsolute));
            }

            // Update text style
            if (GenerateProfileButtonText != null)
            {
                GenerateProfileButtonText.Foreground = profileSelected
                    ? Brushes.White : (Brush)FindResource("TextDark");
                GenerateProfileButtonText.FontWeight = FontWeights.Bold;
            }

            // Joint logic
            bool jointSelected = JointGroupContainer.Children.OfType<RadioButton>().Any(rb => rb.IsChecked == true);
            GenerateJoint.IsEnabled = jointSelected;

            if (GenerateJointButtonIcon != null)
            {
                var uriString = jointSelected
                    ? "/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Generate_Active.png"
                    : "/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Generate.png";
                GenerateJointButtonIcon.Source =
                    new BitmapImage(new Uri(uriString, UriKind.RelativeOrAbsolute));
            }

            if (GenerateJointButtonText != null)
            {
                GenerateJointButtonText.Foreground = jointSelected
                    ? Brushes.White : (Brush)FindResource("TextDark");
                GenerateJointButtonText.FontWeight = FontWeights.Bold;
            }
        }

        // Recursively finds the first Image descendant in a visual subtree, or null if none is found.
        private static Image FindFirstImageChild(DependencyObject parent)
        {
            if (parent == null)
                return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is Image img)
                    return img;

                var found = FindFirstImageChild(child);
                if (found != null)
                    return found;
            }
            return null;
        }

        // Handles selecting a profile radio button and configures the UI for built-in, DXF placeholder, or user profile.
        private void ProfileButton_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is RadioButton selectedRb))
                return;

            var btn = GenerateButton;
            btn.IsEnabled = true;

            PlacementFrame.Visibility = Visibility.Visible;

            var imgb = FindFirstImageChild(btn);
            if (imgb?.Source is BitmapImage bi)
            {
                var uri = bi.UriSource?.OriginalString;
                if (string.IsNullOrEmpty(uri)) return;
                string newUri = uri.EndsWith(".png") && !uri.Contains("_Active")
                      ? uri.Replace(".png", "_Active.png") : uri;

                imgb.Source = new BitmapImage(new Uri(newUri, UriKind.RelativeOrAbsolute));
            }

            foreach (var rb in ProfileGroupContainer.Children.OfType<RadioButton>())
            {
                Image img = FindFirstImageChild(rb);
                if (img == null) continue;

                var uriStr = img.Source?.ToString();
                if (string.IsNullOrEmpty(uriStr))
                    continue;

                if (rb == selectedRb)
                {
                    if (uriStr.EndsWith(".png") && !uriStr.Contains("_Active"))
                    {
                        var activeUri = uriStr.Replace(".png", "_Active.png");
                        img.Source = new BitmapImage(new Uri(activeUri, UriKind.RelativeOrAbsolute));
                    }
                }
                else
                {
                    if (uriStr.Contains("_Active.png"))
                    {
                        var normalUri = uriStr.Replace("_Active.png", ".png");
                        img.Source = new BitmapImage(new Uri(normalUri, UriKind.RelativeOrAbsolute));
                    }
                }
            }

            // then your existing UI logic:
            selectedDXFPath = "";
            dxfContours = null;
            selectedProfileString = "";
            selectedProfileImage = "";
            ProfilePreviewImage.Visibility = Visibility.Collapsed;

            bool isPlaceholder = selectedRb.Name == "DXFProfileButton";
            bool isUserProfile = selectedRb.Tag is DXFProfile;
            bool isBuiltIn = !isPlaceholder && !isUserProfile;

            if (isBuiltIn)
            {
                var builtInTag = (string)selectedRb.Tag;
                selectedProfile = builtInTag;
                SizeSelect.Visibility = Visibility.Visible;
                DynamicFieldsGrid.Visibility = Visibility.Visible;
                ProfilePreviewImage.Visibility = Visibility.Visible;
                UserProfilesGridScrollView.Visibility = Visibility.Collapsed;
                ConvertDXFButton.Visibility = Visibility.Collapsed;

                HollowCheckBox.Visibility = (builtInTag == "Rectangular" || builtInTag == "Circular")
                                            ? Visibility.Visible : Visibility.Collapsed;

                LoadPresetSizes(selectedProfile);
                UpdateDynamicFields();

                string imgKey = builtInTag switch
                {
                    "Rectangular" => "Rect",
                    "Circular" => "Circular",
                    "H" => "H",
                    "L" => "L",
                    "U" => "U",
                    "T" => "T",
                    _ => null
                };
                if (imgKey != null)
                    ShowProfileImage(imgKey);
                else
                    ProfilePreviewImage.Visibility = Visibility.Collapsed;
            }
            else if (isPlaceholder)
            {
                SizeSelect.Visibility = Visibility.Collapsed;
                DynamicFieldsGrid.Visibility = Visibility.Collapsed;
                UserProfilesGridScrollView.Visibility = Visibility.Visible;
                ConvertDXFButton.Visibility = Visibility.Visible;
                ProfilePreviewImage.Visibility = Visibility.Collapsed;
            }
            else if (isUserProfile)
            {
                var userProf = (DXFProfile)selectedRb.Tag;
                selectedProfile = userProf.Name;
                selectedProfileString = userProf.ProfileString;
                selectedProfileImage = userProf.ImgString;
                ProfilePreviewImage.Visibility = Visibility.Collapsed;
            }
        }

        // Resolves the CSV path for a given built-in profile type based on application settings.
        private string GetProfileCsvPathFromSettings(string profileType)
        {
            string relPath = null;

            switch (profileType)
            {
                case "Circular":
                    relPath = Settings.Default.Profiles_Circular;
                    break;
                case "H":
                    relPath = Settings.Default.Profiles_H;
                    break;
                case "L":
                    relPath = Settings.Default.Profiles_L;
                    break;
                case "Rectangular":
                    relPath = Settings.Default.Profiles_Rectangular;
                    break;
                case "T":
                    relPath = Settings.Default.Profiles_T;
                    break;
                case "U":
                    relPath = Settings.Default.Profiles_U;
                    break;
                default:
                    return null; // unknown profile
            }

            if (string.IsNullOrWhiteSpace(relPath))
                return null;

            return Path.IsPathRooted(relPath)
                ? relPath
                : Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "AESCConstruct",
                    relPath
                );
        }

        // Loads preset sizes for the selected built-in profile from CSV into combo and backing lists.
        private void LoadPresetSizes(string profileType)
        {
            string filePath = GetProfileCsvPathFromSettings(profileType);
            if (filePath == null || !File.Exists(filePath))
            {
                Application.ReportStatus($"Could not find CSV for profile type '{profileType}'.", StatusMessageType.Error, null);
                return;
            }

            csvRowNames.Clear();
            SizeComboBox.Items.Clear();
            csvFieldNames.Clear();
            csvDataRows.Clear();

            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    var headerLine = reader.ReadLine();
                    if (!string.IsNullOrEmpty(headerLine))
                    {
                        csvFieldNames = headerLine
                            .Split(';')
                            .Skip(1)
                            .Select(f => f.Trim().Replace(" ", ""))
                            .ToList();
                    }

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (string.IsNullOrEmpty(line))
                            continue;

                        var values = line
                            .Split(';')
                            .Select(v => v.Trim())
                            .ToArray();

                        if (values.Length == csvFieldNames.Count + 1)
                        {
                            csvRowNames.Add(values[0]);
                            csvDataRows.Add(values.Skip(1).ToArray());
                            SizeComboBox.Items.Add(values[0]);
                        }
                        else
                        {
                            Application.ReportStatus($"Mismatched CSV row:\n{line}", StatusMessageType.Error, null);
                        }
                    }

                    int idx = 0;
                    if (_lastSizeByProfile.TryGetValue(profileType, out var lastName))
                    {
                        var found = csvRowNames.FindIndex(n => string.Equals(n, lastName, StringComparison.OrdinalIgnoreCase));
                        if (found >= 0) idx = found;
                    }

                    if (SizeComboBox.Items.Count > 0)
                        SizeComboBox.SelectedIndex = idx;
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to load CSV:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Updates the dimension TextBox fields when a size is selected and remembers last used size per profile.
        private void SizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SizeComboBox.SelectedIndex >= 0 && SizeComboBox.SelectedIndex < csvDataRows.Count)
            {
                string[] selectedValues = csvDataRows[SizeComboBox.SelectedIndex];
                // Logger.Log($"AESCConstruct2026: Loading Size Values: {string.Join(", ", selectedValues)}\n");

                for (int i = 0; i < csvFieldNames.Count; i++)
                {
                    if (inputFieldMap.ContainsKey(csvFieldNames[i]) && i < selectedValues.Length)
                    {
                        inputFieldMap[csvFieldNames[i]].Text = selectedValues[i];
                    }
                }

                if (!string.IsNullOrEmpty(selectedProfile) &&
                SizeComboBox.SelectedIndex >= 0 &&
                SizeComboBox.SelectedIndex < csvRowNames.Count)
                {
                    _lastSizeByProfile[selectedProfile] = csvRowNames[SizeComboBox.SelectedIndex];
                }
            }
        }

        // Rebuilds the dynamic dimension field grid based on csvFieldNames and seeds values from current size selection.
        private void UpdateDynamicFields()
        {
            DynamicFieldsGrid.Children.Clear();
            DynamicFieldsGrid.ColumnDefinitions.Clear();
            DynamicFieldsGrid.RowDefinitions.Clear();
            inputFieldMap.Clear();

            if (csvFieldNames.Count == 0)
            {
                // Logger.Log($"AESCConstruct2026: ERROR - No valid fields found for {selectedProfile}!\n");
                return;
            }

            // make two equal-width columns
            DynamicFieldsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            DynamicFieldsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int rowIndex = 0;
            for (int i = 0; i < csvFieldNames.Count; i++)
            {
                // add a new row for every two fields
                if (i % 2 == 0)
                    DynamicFieldsGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                int columnIndex = i % 2;

                // create sub-grid for one label+input pair, with top margin 5px
                var cellGrid = new Grid
                {
                    Margin = new Thickness(0, 5, 0, 0)
                };
                cellGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) }); // fixed px for label
                cellGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(37) }); // rest

                // the small “one-letter” label
                var label = new TextBlock
                {
                    Text = $"{csvFieldNames[i]}:", // e.g. "w:"
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                };
                Grid.SetColumn(label, 0);

                // the input fills remaining space
                var input = new TextBox
                {
                    Name = $"{csvFieldNames[i]}Input",
                    Margin = new Thickness(5, 0, 5, 0),
                    VerticalAlignment = System.Windows.VerticalAlignment.Center
                };
                Grid.SetColumn(input, 1);

                // assemble
                cellGrid.Children.Add(label);
                cellGrid.Children.Add(input);

                // map for later reads
                inputFieldMap[csvFieldNames[i]] = input;

                // place sub-grid into the parent
                Grid.SetRow(cellGrid, rowIndex);
                Grid.SetColumn(cellGrid, columnIndex);
                DynamicFieldsGrid.Children.Add(cellGrid);

                if (i % 2 != 0)
                    rowIndex++;
            }

            // trigger default selection logic
            SizeComboBox_SelectionChanged(null, null);
        }

        // Toggles visibility of profile size edit fields panel and refreshes dynamic fields.
        private void EditProfileSizes_Click(object sender, RoutedEventArgs e)
        {
            bool currentlyVisible = ProfileSizeEditFields.Visibility == Visibility.Visible;
            ProfileSizeEditFields.Visibility = currentlyVisible
                ? Visibility.Collapsed
                : Visibility.Visible;

            UpdateDynamicFields();
        }

        // Displays the correct profile and placement images for the given built-in profile key.
        void ShowProfileImage(string imgKey)
        {
            if (string.IsNullOrEmpty(imgKey))
            {
                ProfilePreviewImage.Visibility = Visibility.Collapsed;
                return;
            }
            var uri = new Uri($"/AESCConstruct2026;component/FrameGenerator/UI/Images/Img_Measures_Frame_{imgKey}.png", UriKind.Relative);
            ProfilePreviewImage.Source = new BitmapImage(uri);
            ProfilePreviewImage.Visibility = Visibility.Visible;

            var uri2 = new Uri($"/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Frame_{imgKey}_BG.png", UriKind.Relative);
            PlacementFrame.Source = new BitmapImage(uri2);
        }

        // Opens a DXF file, converts it into a DXFProfile, persists profile and preview, and updates the user profiles list.
        private void ConvertDXFButton_Click(object sender, RoutedEventArgs e)
        {
            // 1) Ask the user to select a DXF file
            var dlg = new OpenFileDialog
            {
                Filter = "DXF Files (*.dxf)|*.dxf",
                Title = "Select a DXF File"
            };
            bool? fileOk = dlg.ShowDialog();
            if (fileOk != true)
                return;

            string dxfPath = dlg.FileName;

            // Capture the original document path before any design operations
            string originalDocPath = Window.ActiveWindow?.Document?.Path;

            DXFProfile profile = null;
            Window winDXF = null;

            // ── Phase 1: Design operations INSIDE WriteBlock ──
            WriteBlock.ExecuteTask("Convert DXF Profile", () =>
            {
                Window currentWindow = Window.ActiveWindow;
                Document doc = currentWindow.Document;
                doc.Save();

                try
                {
                    Document.Open(dxfPath, null);
                    winDXF = Window.ActiveWindow;
                }
                catch (Exception ex)
                {
                    Application.ReportStatus($"Failed to open DXF:\n{ex.Message}", StatusMessageType.Information, null);
                    return;
                }

                // Build the profile + preview image (no nested WriteBlock)
                profile = DXFImportHelper.DXFtoProfile(insideWriteBlock: true);

                // Close the DXF window by reference
                try { winDXF?.Close(); } catch { }
            });

            // ── Phase 2: Post-processing OUTSIDE WriteBlock ──
            if (profile == null)
                return;

            // Copy the profile‐string to the clipboard (safe outside WriteBlock)
            try { Clipboard.SetText(profile.ProfileString); } catch { }

            Application.ReportStatus($"DXF → Profile succeeded.\n\nName = {profile.Name}\n(Profile string copied to clipboard.)", StatusMessageType.Information, null);

            // Create UserDXFProfiles folder if needed
            string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string userFolder = Path.Combine(programData, "AESCConstruct", "UserDXFProfiles");

            try
            {
                Directory.CreateDirectory(userFolder);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Warning: Could not create folder:\n{userFolder}\n\n{ex.Message}", StatusMessageType.Error, null);
                return;
            }

            // Decode Base64 preview and save it as a PNG
            string safeName = string.Concat(profile.Name
                .Where(c => !Path.GetInvalidFileNameChars().Contains(c)))
                .Replace(' ', '_');

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string imageFileName = $"{safeName}_{timestamp}.png";
            string imageFullPath = Path.Combine(userFolder, imageFileName);

            try
            {
                byte[] pngBytes = Convert.FromBase64String(profile.ImgString);
                File.WriteAllBytes(imageFullPath, pngBytes);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to save preview image to disk:\n{ex.Message}", StatusMessageType.Error, null);
                imageFileName = "";
            }

            // Append (or create) profiles.csv
            string csvPath = Settings.Default.profiles;

            if (!Path.IsPathRooted(csvPath))
            {
                string baseDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "AESCConstruct"
                );
                csvPath = Path.Combine(baseDir, csvPath);
            }
            bool writeHeader = !File.Exists(csvPath);

            try
            {
                using (var sw = new StreamWriter(csvPath, append: true, encoding: System.Text.Encoding.UTF8))
                {
                    if (writeHeader)
                    {
                        sw.WriteLine("Name;ProfileString;ImageRelativePath");
                    }

                    string escapedProfile = profile.ProfileString.Replace(";", "\\;");
                    string imageRel = string.IsNullOrEmpty(imageFileName) ? "" : imageFileName;

                    sw.WriteLine($"{profile.Name};{escapedProfile};{imageRel}");
                }

                LoadUserProfiles();

                // Auto-check the newest radio button
                var newestRb = UserProfilesGrid.Children
                                 .OfType<RadioButton>()
                                 .LastOrDefault();
                if (newestRb != null)
                    newestRb.IsChecked = true;
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to update CSV:\n{ex.Message}", StatusMessageType.Error, null);
            }

            // Show the saved PNG in a preview dialog (safe outside WriteBlock)
            if (!string.IsNullOrEmpty(imageFullPath) && File.Exists(imageFullPath))
            {
                try
                {
                    using var bmp = new System.Drawing.Bitmap(imageFullPath);
                    var previewForm = new System.Windows.Forms.Form
                    {
                        Text = $"DXF Preview: {profile.Name}",
                        ClientSize = new System.Drawing.Size(bmp.Width, bmp.Height)
                    };
                    var pictureBox = new System.Windows.Forms.PictureBox
                    {
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Image = new System.Drawing.Bitmap(bmp),
                        SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                    };
                    previewForm.Controls.Add(pictureBox);
                    previewForm.ShowDialog();
                }
                catch (Exception ex)
                {
                    Application.ReportStatus($"Failed to render saved preview image:\n{ex.Message}", StatusMessageType.Error, null);
                }
            }

            // Smart reopen: only if the original document isn't already the active one
            if (!string.IsNullOrEmpty(originalDocPath))
            {
                var activeDoc = Window.ActiveWindow?.Document;
                if (activeDoc == null || !string.Equals(activeDoc.Path, originalDocPath, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        Document.Open(originalDocPath, null);
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Failed to reopen document:\n{ex.Message}", StatusMessageType.Error, null);
                    }
                }
            }
        }

        // Loads user-defined DXF profiles from CSV and builds the UI list with preview images and delete buttons.
        private void LoadUserProfiles()
        {
            UserProfilesGrid.Children.Clear();

            try
            {
                var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                var userFolder = Path.Combine(programData, "AESCConstruct", "UserDXFProfiles");
                //var csvPath = Path.Combine(userFolder, "profiles.csv");
                var csvPath = Settings.Default.profiles;
                if (!File.Exists(csvPath))
                    return;

                foreach (var line in File.ReadAllLines(csvPath).Skip(1))
                {
                    var raw = line.Split(';');
                    if (raw.Length < 3) continue;

                    var prof = new DXFProfile
                    {
                        Name = raw[0],
                        ProfileString = raw[1].Replace("\\;", ";"),
                        ImgString = raw[2]
                    };

                    var container = new Grid
                    {
                        Width = 230,
                        Height = 70,
                        Margin = new Thickness(0)
                    };

                    // Create RadioButton with stackpanel content
                    var contentStack = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Left
                    };

                    // Image
                    var imageControl = new Image
                    {
                        Width = 60,
                        Height = 60,
                        Margin = new Thickness(5, 0, 10, 0)
                    };

                    if (!string.IsNullOrEmpty(prof.ImgString))
                    {
                        var imgPath = Path.Combine(userFolder, prof.ImgString);
                        if (File.Exists(imgPath))
                        {
                            try
                            {
                                var bitmap = new BitmapImage();
                                bitmap.BeginInit();
                                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                bitmap.UriSource = new Uri(imgPath, UriKind.Absolute);
                                bitmap.EndInit();
                                bitmap.Freeze();
                                imageControl.Source = bitmap;
                            }
                            catch (Exception ex)
                            {
                                Application.ReportStatus($"Failed to load image for profile \"{prof.Name}\":\n{ex.Message}", StatusMessageType.Error, null);
                            }
                        }
                    }

                    // Text label
                    var label = new TextBlock
                    {
                        Text = prof.Name,
                        FontSize = 12,
                        Width = 130,
                        TextAlignment = System.Windows.TextAlignment.Center,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        TextWrapping = TextWrapping.Wrap
                    };

                    contentStack.Children.Add(imageControl);
                    contentStack.Children.Add(label);

                    var radioButton = new RadioButton
                    {
                        GroupName = "ProfileGroup",
                        Tag = prof,
                        Width = 230,
                        Height = 70,
                        Margin = new Thickness(0, 0, 0, 0),
                        Content = contentStack,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Left
                    };
                    radioButton.Checked += ProfileButton_Checked;

                    // Delete button (✕)
                    var delBtn = new Button
                    {
                        Width = 16,
                        Height = 16,
                        BorderThickness = new Thickness(0),
                        VerticalAlignment = System.Windows.VerticalAlignment.Top,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                        Margin = new Thickness(0, 8, 12, 0),
                        ToolTip = "Remove this profile"
                    };

                    // Load the image
                    var img = new Image
                    {
                        Width = 12,
                        Height = 12,
                        Source = new BitmapImage(new Uri("/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Delete.png", UriKind.Relative))
                    };

                    // Set image as content
                    delBtn.Content = img;
                    delBtn.Click += (s, e) =>
                    {
                        var confirm = MessageBox.Show(
                            $"Are you sure you want to delete the custom profile \"{prof.Name}\"?",
                            "Confirm Deletion",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Warning
                        );

                        if (confirm != MessageBoxResult.Yes)
                            return;

                        try
                        {
                            var updated = File.ReadAllLines(csvPath)
                                              .Where(l => !l.StartsWith(prof.Name + ";"))
                                              .ToArray();
                            File.WriteAllLines(csvPath, updated);
                        }
                        catch (Exception ex)
                        {
                            Application.ReportStatus($"Failed to update CSV when deleting profile:\n{ex.Message}", StatusMessageType.Error, null);
                        }

                        if (!string.IsNullOrEmpty(prof.ImgString))
                        {
                            string imgPath = Path.Combine(userFolder, prof.ImgString);
                            if (File.Exists(imgPath))
                            {
                                try
                                {
                                    File.Delete(imgPath);
                                }
                                catch (Exception ex)
                                {
                                    Application.ReportStatus($"Failed to delete image for profile \"{prof.Name}\":\n{ex.Message}", StatusMessageType.Error, null);
                                }
                            }
                        }

                        LoadUserProfiles();

                        if (selectedProfileString == prof.ProfileString)
                            selectedProfileString = "";
                    };

                    container.Children.Add(radioButton);
                    container.Children.Add(delBtn);

                    UserProfilesGrid.Children.Add(container);
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Error loading user profiles:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }


        // Generates the selected profile (user-saved, DXF or built-in) and optionally rotates new components.
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            WriteBlock.ExecuteTask("Generate Profile", () =>
            {
                var win = Window.ActiveWindow;
                if (win == null)
                    return;

                var existingComps = win.Document.MainPart
                  .GetChildren<Component>()
                  .ToHashSet();

                double.TryParse(
                    RotationAngleTextBox.Text,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out rotationAngle
                );

                var oldOri = Application.UserOptions.WorldOrientation;
                Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
                try
                {
                    // If nothing has been chosen, bail out:
                    if (string.IsNullOrEmpty(selectedDXFPath)
                     && inputFieldMap.Count == 0
                     && string.IsNullOrEmpty(selectedProfileString))
                    {
                        Application.ReportStatus("Please select a profile (built-in or user-saved) or load a DXF file before generating.", StatusMessageType.Warning, null);
                        return;
                    }

                    bool isHollow = HollowCheckBox.IsChecked == false;
                    bool updateBOM = UpdateBOM.IsChecked == true;
                    //
                    // ─── 1) USER‐SAVED PROFILE (selectedProfileString != "") ───────────────────────────────
                    //
                    // Log exactly the raw string (with no trailing “.”):
                    if (!string.IsNullOrEmpty(selectedProfileString))
                    {
                        // ─── 1a) SPLIT INTO INDIVIDUAL CURVE STRINGS ───────────────────────────────────────
                        var curves = new List<ITrimmedCurve>();

                        // First split on '&' → each loop
                        string[] loopStrings = selectedProfileString.Split('&');

                        for (int loopIndex = 0; loopIndex < loopStrings.Length; loopIndex++)
                        {
                            string loopStr = loopStrings[loopIndex].Trim();

                            string[] curveChunks = loopStr.Split(' ');

                            for (int i = 0; i < curveChunks.Length; i++)
                            {
                                string curveStr = curveChunks[i].Trim();
                                if (string.IsNullOrEmpty(curveStr))
                                    continue;

                                try
                                {
                                    ITrimmedCurve c = DXFImportHelper.CurveFromString(curveStr);
                                    if (c != null)
                                    {
                                        curves.Add(c);

                                        // Log the geometry type and key points:
                                        if (c is CurveSegment seg && seg.Geometry is Line)
                                        {
                                            var ps = seg.StartPoint;
                                            var pe = seg.EndPoint;
                                        }
                                        else if (c is CurveSegment arcSeg && arcSeg.Geometry is Circle cir)
                                        {
                                            var center = cir.Axis.Origin;
                                            double radius = cir.Radius;
                                        }
                                    }
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }

                        var firstSeg = (CurveSegment)curves[0];
                        var lastSeg = (CurveSegment)curves[curves.Count - 1];
                        var diff = (firstSeg.StartPoint - lastSeg.EndPoint).Magnitude;

                        if (curves.Count == 0)
                        {
                            Application.ReportStatus("Could not reconstruct any curves from the saved profile string.", StatusMessageType.Warning, null);
                            return;
                        }

                        // ─── 1b) COMPUTE BOUNDING‐BOX (width, height) ─────────────────────────────────────────
                        var (profW, profH) = DXFImportHelper.GetDXFSize(curves);
                        double w = profW;
                        double h = profH;

                        // ─── 1c) FIGURE OUT OFFSET RADIO CHOICE ──────────────────────────────────────────────
                        double offsetX = 0, offsetY = 0;
                        if (Offset_TopLeft.IsChecked == true) { offsetX = w / 2; offsetY = -h / 2; }
                        else if (Offset_TopCenter.IsChecked == true) { offsetX = 0; offsetY = -h / 2; }
                        else if (Offset_TopRight.IsChecked == true) { offsetX = -w / 2; offsetY = -h / 2; }
                        else if (Offset_MiddleLeft.IsChecked == true) { offsetX = w / 2; offsetY = 0; }
                        else if (Offset_Center.IsChecked == true) { offsetX = 0; offsetY = 0; }
                        else if (Offset_MiddleRight.IsChecked == true) { offsetX = -w / 2; offsetY = 0; }
                        else if (Offset_BottomLeft.IsChecked == true) { offsetX = w / 2; offsetY = h / 2; }
                        else if (Offset_BottomCenter.IsChecked == true) { offsetX = 0; offsetY = h / 2; }
                        else if (Offset_BottomRight.IsChecked == true) { offsetX = -w / 2; offsetY = h / 2; }

                        foreach (var c in curves)
                        {
                            if (c is CurveSegment seg)
                            {
                                var ps = seg.StartPoint;
                                var pe = seg.EndPoint;
                            }
                        }
                        // ─── 1d) EXTRUDE THE CURVES ───────────────────────────────────────────────────────────
                        try
                        {
                            ExtrudeProfileCommand.ExecuteExtrusion(
                                "CSV",              // e.g. “square”
                                curves,            // List<ITrimmedCurve>
                                false,             // always solid for user-saved
                                offsetX,
                                offsetY,
                                "",       // no DXF file path in this branch
                                updateBOM,
                                selectedProfileString
                            );
                            //});
                        }
                        catch (Exception)
                        {
                            Application.ReportStatus("An error occurred while extruding the user-saved profile.\nSee log in addin folder for details.", StatusMessageType.Error, null);
                        }
                        return;
                    }

                    //
                    // ─── 2) LOADED DXF (selectedDXFPath != "" && dxfContours != null) ───────────────────────────────────
                    //
                    if (!string.IsNullOrEmpty(selectedDXFPath) && dxfContours != null)
                    {
                        selectedProfile = "DXF";
                        var (DXFwidth, DXFheight) = DXFImportHelper.GetDXFSize(dxfContours);
                        double w = DXFwidth;
                        double h = DXFheight;

                        double offsetX = 0, offsetY = 0;
                        if (Offset_TopLeft.IsChecked == true) { offsetX = w / 2; offsetY = -h / 2; }
                        else if (Offset_TopCenter.IsChecked == true) { offsetX = 0; offsetY = -h / 2; }
                        else if (Offset_TopRight.IsChecked == true) { offsetX = -w / 2; offsetY = -h / 2; }
                        else if (Offset_MiddleLeft.IsChecked == true) { offsetX = w / 2; offsetY = 0; }
                        else if (Offset_Center.IsChecked == true) { offsetX = 0; offsetY = 0; }
                        else if (Offset_MiddleRight.IsChecked == true) { offsetX = -w / 2; offsetY = 0; }
                        else if (Offset_BottomLeft.IsChecked == true) { offsetX = w / 2; offsetY = h / 2; }
                        else if (Offset_BottomCenter.IsChecked == true) { offsetX = 0; offsetY = h / 2; }
                        else if (Offset_BottomRight.IsChecked == true) { offsetX = -w / 2; offsetY = h / 2; }

                        try
                        {
                            ExtrudeProfileCommand.ExecuteExtrusion(
                                selectedProfile,   // “DXF”
                                dxfContours,       // List<ITrimmedCurve>
                                false,             // always solid for imported DXF
                                offsetX,
                                offsetY,
                                selectedDXFPath,    // pass the original .dxf path
                                updateBOM,
                                ""
                            );
                        }
                        catch (Exception)
                        {
                            Application.ReportStatus("An error occurred while extruding the DXF contours.\nSee log in addin folder for details.", StatusMessageType.Error, null);
                        }
                        return;
                    }

                    //
                    // ─── 3) BUILT‐IN SHAPES (Rectangular, Circular, H, L, U, T) ───────────────────────────────
                    //
                    double wBuilt = 0, hBuilt = 0;
                    if (selectedProfile == "Circular")
                    {
                        wBuilt = inputFieldMap.ContainsKey("D")
                            ? double.Parse(inputFieldMap["D"].Text.Replace(',', '.'), CultureInfo.InvariantCulture) / 1000
                            : 0.0;
                        hBuilt = wBuilt;
                    }
                    else if (selectedProfile == "L")
                    {
                        wBuilt = inputFieldMap.ContainsKey("b")
                            ? double.Parse(inputFieldMap["b"].Text.Replace(',', '.'), CultureInfo.InvariantCulture) / 1000
                            : 0.0;
                        hBuilt = inputFieldMap.ContainsKey("a")
                            ? double.Parse(inputFieldMap["a"].Text.Replace(',', '.'), CultureInfo.InvariantCulture) / 1000
                            : 0.0;
                    }
                    else
                    {
                        wBuilt = inputFieldMap.ContainsKey("w")
                            ? double.Parse(inputFieldMap["w"].Text.Replace(',', '.'), CultureInfo.InvariantCulture) / 1000
                            : 0.0;
                        hBuilt = inputFieldMap.ContainsKey("h")
                            ? double.Parse(inputFieldMap["h"].Text.Replace(',', '.'), CultureInfo.InvariantCulture) / 1000
                            : 0.0;
                    }

                    double offsetXB = 0, offsetYB = 0;
                    if (Offset_TopLeft.IsChecked == true) { offsetXB = wBuilt / 2; offsetYB = -hBuilt / 2; }
                    else if (Offset_TopCenter.IsChecked == true) { offsetXB = 0; offsetYB = -hBuilt / 2; }
                    else if (Offset_TopRight.IsChecked == true) { offsetXB = -wBuilt / 2; offsetYB = -hBuilt / 2; }
                    else if (Offset_MiddleLeft.IsChecked == true) { offsetXB = wBuilt / 2; offsetYB = 0; }
                    else if (Offset_Center.IsChecked == true) { offsetXB = 0; offsetYB = 0; }
                    else if (Offset_MiddleRight.IsChecked == true) { offsetXB = -wBuilt / 2; offsetYB = 0; }
                    else if (Offset_BottomLeft.IsChecked == true) { offsetXB = wBuilt / 2; offsetYB = hBuilt / 2; }
                    else if (Offset_BottomCenter.IsChecked == true) { offsetXB = 0; offsetYB = hBuilt / 2; }
                    else if (Offset_BottomRight.IsChecked == true) { offsetXB = -wBuilt / 2; offsetYB = hBuilt / 2; }

                    bool isHollowBuilt = HollowCheckBox.IsChecked == false;

                    try
                    {
                        var dataDict = inputFieldMap
                        .ToDictionary(
                            kvp => kvp.Key,                // e.g. "w", "h", "t", etc.
                            kvp => kvp.Value.Text.Trim()   // the string value the user typed
                        );

                        dataDict["Name"] = csvRowNames[SizeComboBox.SelectedIndex];

                        ExtrudeProfileCommand.ExecuteExtrusion(
                            selectedProfile,  // e.g. "Rectangular", "H", etc.
                            dataDict,    // numeric‐based sizes
                            isHollowBuilt,
                            offsetXB,
                            offsetYB,
                            "",
                            updateBOM,
                            ""
                        );
                    }
                    catch (Exception)
                    {
                        Application.ReportStatus("An error occurred while extruding the built‐in profile.", StatusMessageType.Error, null);
                    }
                }
                finally
                {
                    var currentComps = win.Document.MainPart
                          .GetChildren<Component>()
                          .ToList();

                    var newComps = currentComps
                          .Where(c => !existingComps.Contains(c))
                          .ToList();

                    // 3) if we have an angle and new components, rotate them
                    if (rotationAngle != 0.0 && newComps.Count > 0)
                    {
                        RotateComponentCommand.ApplyRotation(win, newComps, rotationAngle);
                    }

                    // restore whatever the user had selected
                    Application.UserOptions.WorldOrientation = oldOri;
                }
            });
        }

        // Placeholder for handling hollow checkbox state changes (currently unused).
        private void HollowCheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }


        ////////////// joint 

        // Parses joint options from the UI and executes the selected joint type via ExecuteJointCommand.
        private void GenerateJoint_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                // parse gap (mm → m)
                double spacing = 0;
                if (double.TryParse(Gap.Text, out var mm))
                    spacing = mm / 1000.0;

                // pick joint type
                string jointType =
                    MiterJoint.IsChecked == true ? "Miter" :
                    StraightJoint.IsChecked == true ? "Straight" :
                    StraightJoint2.IsChecked == true ? "Straight2" :
                    TJoint.IsChecked == true ? "T" :
                    CutOut.IsChecked == true ? "CutOut" :
                    Trim.IsChecked == true ? "Trim" :
                    "None"
                ;

                bool updateBOM = UpdateBOMJoint.IsChecked == true;

                // do it inside a write‐block
                WriteBlock.ExecuteTask("Execute Joint", () =>
                {
                    ExecuteJointCommand.ExecuteJoint(window, spacing, jointType, updateBOM);
                });
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        // Calls ExecuteJointCommand.RestoreGeometry to undo joint geometry modifications for the active document.
        private void RestoreGeometry_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                WriteBlock.ExecuteTask("Restore Geometry", () =>
                {
                    ExecuteJointCommand.RestoreGeometry(window);
                });
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        // Calls ExecuteJointCommand.RestoreJoint to restore joint helper data in the active document.
        private void RestoreJoint_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                WriteBlock.ExecuteTask("Restore Joint", () =>
                {
                    ExecuteJointCommand.RestoreJoint(window);
                });
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        // Updates the placement icons so that the selected placement radio button shows an active image.
        private void PlacementRadio_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is RadioButton selectedRb))
                return;

            foreach (var rb in PlacementGrid.Children.OfType<RadioButton>())
            {
                var img = FindDescendant<Image>(rb);
                if (img == null) continue;

                var uri = img.Source.ToString();
                if (rb == selectedRb)
                {
                    if (uri.EndsWith(".png") && !uri.Contains("_Active.png"))
                    {
                        var active = uri.Replace(".png", "_Active.png");
                        img.Source = new BitmapImage(new Uri(active));
                    }
                }
                else
                {
                    if (uri.Contains("_Active.png"))
                    {
                        var normal = uri.Replace("_Active.png", ".png");
                        img.Source = new BitmapImage(new Uri(normal));
                    }
                }
            }
        }

        // Recursively searches the visual tree starting at parent for a descendant of type T.
        private static T FindDescendant<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) return t;
                var found = FindDescendant<T>(child);
                if (found != null) return found;
            }
            return null;
        }

        // Filters the displayed user profiles in the grid based on the text entered in the search box.
        private void ProfileSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var filter = ProfileSearchBox.Text.Trim();
            if (string.IsNullOrEmpty(filter))
            {
                // show all
                foreach (UIElement child in UserProfilesGrid.Children)
                    child.Visibility = Visibility.Visible;
            }
            else
            {
                var lower = filter.ToLowerInvariant();
                foreach (UIElement child in UserProfilesGrid.Children)
                {
                    // Each child is a Container Grid, whose first child is the RadioButton
                    if (child is Grid container && container.Children.OfType<RadioButton>().FirstOrDefault() is RadioButton rb)
                    {
                        // Tag is a DXFProfile for user-loaded profiles
                        if (rb.Tag is DXFProfile prof)
                        {
                            container.Visibility = prof.Name
                                                     .ToLowerInvariant()
                                                     .Contains(lower)
                                ? Visibility.Visible
                                : Visibility.Collapsed;
                        }
                        else
                        {
                            // if somehow other kinds of buttons got in here, leave them visible
                            container.Visibility = Visibility.Visible;
                        }
                    }
                }
            }
        }

        // Deletes a selected Construct profile component and restores its original driving curves where possible.
        private void DeleteProfileButton_Click(object sender, RoutedEventArgs e)
        {
            {
                var win = Window.ActiveWindow;
                if (win == null)
                    return;

                WriteBlock.ExecuteTask("Delete Construct Profile", () =>
                {
                    try
                    {
                        var sel = win.ActiveContext?.Selection;
                        if (sel == null || sel.Count == 0)
                        {
                            Application.ReportStatus("Select a line, edge, face or curve that belongs to a profile component.", StatusMessageType.Warning, null);
                            return;
                        }

                        Logger.Log($"sel.Count : {sel.Count}");
                        // Try every picked item; delete the first eligible component we can find.
                        foreach (var picked in sel)
                        {
                            Logger.Log($"Picked item type: {picked?.GetType().Name}");
                            var comp = TryGetOwningComponent(win.Document, picked);
                            Logger.Log($"Picked comp: {comp.Name}");
                            if (comp == null) continue;

                            var part = comp.Template; // Component.Template is the owning Part (not .Master)  ← fixes CS1061
                            Logger.Log($"Picked part: {part}");
                            if (part == null) continue;

                            // Presence of AESC_Construct marks our "Construct" components (profiles & plates). 
                            // Profiles set it to 'true', plates set "Plates" — key presence is enough. 
                            var props = part.CustomProperties;
                            bool isConstruct = props.ContainsKey("AESC_Construct");
                            if (!isConstruct)
                                continue;

                            // Best-effort: restore hidden original (top-level) curves; ConstructCurve lives inside the component
                            // and will be removed with the component.
                            TryUnhideOriginalCurves(win.Document, comp);
                            comp.Delete();

                            Application.ReportStatus(
                                $"Deleted profile component \"{part.DisplayName}\" and restored original curve visibility (where applicable).",
                                StatusMessageType.Information, null);
                            //return;
                        }

                        //Application.ReportStatus("No Construct profile component found on the current selection.", StatusMessageType.Warning, null);
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Delete Profile failed: {ex.Message}", StatusMessageType.Error, null);
                    }
                });
            }
        }

        /// <summary>
        /// Find the Component under MainPart that owns the picked item.
        /// Works for IDesign* wrappers and raw Design* objects (curve/edge/face/body),
        /// and for selecting the Component itself.
        /// </summary>
        // Resolves the owning Component in MainPart for a picked geometry object or wrapper.
        private static Component TryGetOwningComponent(Document doc, object picked)
        {
            if (picked == null || doc == null) return null;

            // Unwrap common selection wrappers (IDesign*) to their Design* masters
            DesignCurve dc = null;
            DesignEdge de = null;
            DesignFace df = null;
            DesignBody db = null;

            if (picked is Component c0) return c0;

            if (picked is IDesignCurve idc) dc = idc.Master;
            else if (picked is DesignCurve dc0) dc = dc0;

            if (picked is IDesignEdge ide) de = ide.Master;
            else if (picked is DesignEdge de0) de = de0;

            if (picked is IDesignFace idf) df = idf.Master;
            else if (picked is DesignFace df0) df = df0;

            if (picked is IDesignBody idb) db = idb.Master;
            else if (picked is DesignBody db0) db = db0;

            // If we have an owning DesignCurve (e.g., selecting ConstructCurve or any curve inside a Part),
            // find the component whose Template contains that exact DesignCurve instance.
            foreach (var comp in doc.MainPart.GetChildren<Component>())
            {
                var part = comp.Template;
                if (part == null) continue;

                if (dc != null)
                {
                    foreach (var c in part.GetChildren<DesignCurve>())
                        if (ReferenceEquals(c, dc))
                            return comp;
                }

                // For face/edge: compare geometry Body to each DesignBody.Shape in the Part.
                // (DesignFace.Shape.Body / DesignEdge.Shape.Body is a Modeler.Body we can match)
                if (de != null)
                {
                    var geomBody = de.Shape.Body;
                    if (part.Bodies.Any(b => b.Shape == geomBody))
                        return comp;
                }
                if (df != null)
                {
                    var geomBody = df.Shape.Body;
                    if (part.Bodies.Any(b => b.Shape == geomBody))
                        return comp;
                }

                // For DesignBody directly
                if (db != null && part.Bodies.Any(b => ReferenceEquals(b, db)))
                    return comp;
            }

            return null;
        }

        /// <summary>
        /// Make original (top-level) design curves visible again in all open views.
        /// Skips any curve named "ConstructCurve" (those live inside the component and are deleted with it).
        /// </summary>
        // Restores visibility of original top-level curves that correspond to the component's ConstructCurve helpers.
        private static void TryUnhideOriginalCurves(Document doc, Component comp)
        {
            try
            {
                if (doc == null || comp == null || comp.Template == null)
                    return;

                var viewContexts = Window.GetWindows(doc)
                                         .Where(w => w != null && w.Document == doc)
                                         .Select(w => w.ActiveContext as IAppearanceContext)
                                         .Where(ctx => ctx != null)
                                         .ToList();

                // Collect world-space endpoints from all ConstructCurve(s) inside the component
                Matrix placement = comp.Placement; // component local → document/world
                var constructWorldEnds = new List<(Point A, Point B)>();

                foreach (var c in comp.Template.GetChildren<DesignCurve>())
                {
                    if (!string.Equals(c.Name, "ConstructCurve", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (TryGetCurveEndpointsLocal(c, out var aLocal, out var bLocal))
                    {
                        var aWorld = placement * aLocal;
                        var bWorld = placement * bLocal;
                        constructWorldEnds.Add((aWorld, bWorld));
                    }
                }

                if (constructWorldEnds.Count == 0)
                    return;

                const double tol = 1e-5;

                bool MatchesAnyConstructEnds(Point p0, Point p1)
                {
                    foreach (var pair in constructWorldEnds)
                    {
                        var A = pair.A;
                        var B = pair.B;

                        // same orientation
                        if (PointsEqual(p0, A, tol) && PointsEqual(p1, B, tol))
                            return true;
                        // reversed
                        if (PointsEqual(p0, B, tol) && PointsEqual(p1, A, tol))
                            return true;
                    }
                    return false;
                }

                // Unhide only the original curves in MainPart that match those endpoints
                foreach (var curve in doc.MainPart.GetChildren<DesignCurve>())
                {
                    if (string.Equals(curve.Name, "ConstructCurve", StringComparison.OrdinalIgnoreCase))
                        continue; // top-level helpers should not exist; skip defensively

                    if (!TryGetCurveEndpointsLocal(curve, out var p0, out var p1))
                        continue;

                    if (!MatchesAnyConstructEnds(p0, p1))
                        continue;

                    curve.SetVisibility(null, true);
                    foreach (var ctx in viewContexts)
                        curve.SetVisibility(ctx, true);
                }
            }
            catch
            {
            }
        }

        // Reads endpoints in the curve's own (local) part coordinates for either CurveSegment or ITrimmedCurve.
        private static bool TryGetCurveEndpointsLocal(DesignCurve dc, out Point start, out Point end)
        {
            start = default(Point);
            end = default(Point);

            var shape = dc.Shape;

            var seg = shape as CurveSegment;
            if (seg != null)
            {
                start = seg.StartPoint;
                end = seg.EndPoint;
                return true;
            }

            var trimmed = shape as ITrimmedCurve;
            if (trimmed != null)
            {
                start = trimmed.StartPoint;
                end = trimmed.EndPoint;
                return true;
            }

            return false;
        }

        // Compares two geometry points with a given distance tolerance.
        private static bool PointsEqual(Point a, Point b, double tol)
        {
            return (a - b).Magnitude <= tol;
        }

        // Placeholder handler for rotation angle text changes; rotation is read when generating profiles.
        private void RotationAngleTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
