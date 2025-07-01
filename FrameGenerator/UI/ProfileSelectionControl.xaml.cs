using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;                             // for WPF Window, MessageBox, etc.
using System.Windows.Controls;                    // for WPF UserControl, TextBox, RadioButton, etc.
using SpaceClaim.Api.V242;                        // for SpaceClaim API (Document, Window.ActiveWindow)
using SpaceClaim.Api.V242.Geometry;               // for ITrimmedCurve, Point, etc.
using System.Windows.Media.Imaging;               // if you ever need WPF BitmapImage
using DXFProfile = AESCConstruct25.FrameGenerator.Utilities.DXFProfile;
using UserControl = System.Windows.Controls.UserControl;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using AESCConstruct25.FrameGenerator.Commands;
using Orientation = System.Windows.Controls.Orientation;
using Application = SpaceClaim.Api.V242.Application;
using AESCConstruct25.FrameGenerator.Utilities;  // alias our DXFProfile class

namespace AESCConstruct25.FrameGenerator.UI
{
    public partial class ProfileSelectionControl : UserControl
    {
        private static string logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "AESCConstruct25_Log.txt"
        );

        private string selectedProfile = "";
        private string selectedProfileString = "";
        private string selectedProfileImage = "";

        private List<string> csvFieldNames = new List<string>();
        private List<string[]> csvDataRows = new List<string[]>();
        private Dictionary<string, System.Windows.Controls.TextBox> inputFieldMap
            = new Dictionary<string, System.Windows.Controls.TextBox>();
        private string selectedDXFPath = "";
        private List<ITrimmedCurve> dxfContours = null;

        public ProfileSelectionControl()
        {
            InitializeComponent();
            RectangularProfileButton.IsChecked = true;
            LoadUserProfiles();
        }

        //private void ProfileButton_Checked(object sender, RoutedEventArgs e)
        //{
        //    if (sender is System.Windows.Controls.RadioButton rb && rb.Tag is string profileName)
        //    {
        //        selectedDXFPath = "";
        //        dxfContours = null;

        //        selectedProfile = profileName;
        //        File.AppendAllText(logPath, $"AESCConstruct25: Profile selected in UI: {selectedProfile}\n");

        //        LoadPresetSizes(selectedProfile);

        //        DynamicFieldsGrid.Visibility = Visibility.Visible;
        //        SizeComboBox.Visibility = Visibility.Visible;
        //        HollowCheckBox.Visibility = (selectedProfile == "Rectangular" || selectedProfile == "Circular")
        //                                     ? Visibility.Visible
        //                                     : Visibility.Collapsed;

        //        UpdateDynamicFields();
        //    }
        //}
        private void ProfileButton_Checked(object sender, RoutedEventArgs e)
        {
            if (!(sender is RadioButton rb))
                return;

            // clear any previous DXF selections
            selectedDXFPath = "";
            dxfContours = null;
            selectedProfileString = "";
            selectedProfileImage = "";

            // use the RadioButton’s Name to distinguish the “DXF” placeholder
            bool isPlaceholder = rb.Name == "DXFProfileButton";
            bool isUserProfile = rb.Tag is DXFProfile;
            bool isBuiltIn = !isPlaceholder && !isUserProfile;

            // 1) Built-in shapes: show SizeSelect + DynamicFieldsGrid, hide all DXF UI
            if (isBuiltIn)
            {
                var builtInTag = (string)rb.Tag;
                selectedProfile = builtInTag;

                SizeSelect.Visibility = Visibility.Visible;
                DynamicFieldsGrid.Visibility = Visibility.Visible;

                UserProfilesGridScrollView.Visibility = Visibility.Collapsed;
                LoadDXFButton.Visibility = Visibility.Collapsed;
                ConvertDXFButton.Visibility = Visibility.Collapsed;

                HollowCheckBox.Visibility =
                    (builtInTag == "Rectangular" || builtInTag == "Circular")
                      ? Visibility.Visible
                      : Visibility.Collapsed;

                LoadPresetSizes(selectedProfile);
                UpdateDynamicFields();
            }
            // 2) The “DXF” placeholder button: show the folder/grid of available DXF profiles
            else if (isPlaceholder)
            {
                SizeSelect.Visibility = Visibility.Collapsed;
                DynamicFieldsGrid.Visibility = Visibility.Collapsed;

                UserProfilesGridScrollView.Visibility = Visibility.Visible;
                LoadDXFButton.Visibility = Visibility.Visible;
                ConvertDXFButton.Visibility = Visibility.Visible;
            }
            //// 3) An actual user-loaded DXFProfile: hide the placeholder UI, show DynamicFieldsGrid
            else if (isUserProfile)
            {
                var userProf = (DXFProfile)rb.Tag;
                selectedProfile = userProf.Name;
                selectedProfileString = userProf.ProfileString;
                selectedProfileImage = userProf.ImgString;
            }
        }

        private void LoadPresetSizes(string profileType)
        {
            string resourceName = $"AESCConstruct25.FrameGenerator.Resources.Presets.Profiles_{profileType}.csv";
            SizeComboBox.Items.Clear();
            csvFieldNames.Clear();
            csvDataRows.Clear();

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                using (StreamReader reader = new StreamReader(stream))
                {
                    string headerLine = reader.ReadLine();
                    if (!string.IsNullOrEmpty(headerLine))
                    {
                        csvFieldNames = headerLine.Split(';')
                                                 .Skip(1)
                                                 .Select(f => f.Trim().Replace(" ", ""))
                                                 .ToList();
                        File.AppendAllText(logPath, $"AESCConstruct25: Loaded CSV fields: {string.Join(", ", csvFieldNames)}\n");
                    }

                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (!string.IsNullOrEmpty(line))
                        {
                            string[] values = line.Split(';').Select(v => v.Trim()).ToArray();
                            if (values.Length == csvFieldNames.Count + 1)
                            {
                                csvDataRows.Add(values.Skip(1).ToArray());
                                SizeComboBox.Items.Add(values[0]);
                            }
                            else
                            {
                                File.AppendAllText(logPath, $"AESCConstruct25: WARNING - Mismatched CSV row: {line}\n");
                            }
                        }
                    }
                }

                if (SizeComboBox.Items.Count > 0)
                    SizeComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"AESCConstruct25: ERROR - Failed to load CSV: {resourceName} ({ex.Message})\n");
            }
        }

        private void SizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SizeComboBox.SelectedIndex >= 0 && SizeComboBox.SelectedIndex < csvDataRows.Count)
            {
                string[] selectedValues = csvDataRows[SizeComboBox.SelectedIndex];
                File.AppendAllText(logPath, $"AESCConstruct25: Loading Size Values: {string.Join(", ", selectedValues)}\n");

                for (int i = 0; i < csvFieldNames.Count; i++)
                {
                    if (inputFieldMap.ContainsKey(csvFieldNames[i]) && i < selectedValues.Length)
                    {
                        inputFieldMap[csvFieldNames[i]].Text = selectedValues[i];
                    }
                }
            }
        }

        private void UpdateDynamicFields()
        {
            DynamicFieldsGrid.Children.Clear();
            DynamicFieldsGrid.ColumnDefinitions.Clear();
            DynamicFieldsGrid.RowDefinitions.Clear();
            inputFieldMap.Clear();

            if (csvFieldNames.Count == 0)
            {
                File.AppendAllText(logPath, $"AESCConstruct25: ERROR - No valid fields found for {selectedProfile}!\n");
                return;
            }

            DynamicFieldsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            DynamicFieldsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            int rowIndex = 0;
            for (int i = 0; i < csvFieldNames.Count; i++)
            {
                if (i % 2 == 0)
                    DynamicFieldsGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                int columnIndex = i % 2;

                var label = new TextBlock
                {
                    Text = $"{csvFieldNames[i]}:",
                    Margin = new Thickness(5)
                };
                Grid.SetRow(label, rowIndex);
                Grid.SetColumn(label, columnIndex);

                var input = new System.Windows.Controls.TextBox
                {
                    Name = $"{csvFieldNames[i]}Input",
                    Margin = new Thickness(5)
                };
                Grid.SetRow(input, rowIndex);
                Grid.SetColumn(input, columnIndex);

                inputFieldMap[csvFieldNames[i]] = input;
                DynamicFieldsGrid.Children.Add(label);
                DynamicFieldsGrid.Children.Add(input);

                if (i % 2 != 0)
                    rowIndex++;
            }

            SizeComboBox_SelectionChanged(null, null);
        }

        private void LoadDXFButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "DXF Files (*.dxf)|*.dxf",
                Title = "Select a DXF File"
            };

            bool? result = dlg.ShowDialog();
            if (result != true)
                return;

            selectedDXFPath = dlg.FileName;
            File.AppendAllText(logPath, $"AESCConstruct25: Selected DXF File - {selectedDXFPath}\n");

            if (!DXFImportHelper.ImportDXFContours(selectedDXFPath, out dxfContours))
            {
                System.Windows.MessageBox.Show(
                    "Failed to load DXF contours.",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                selectedDXFPath = "";
            }
            else
            {
                System.Windows.MessageBox.Show(
                    "DXF file loaded successfully. Click 'Generate' to process.",
                    "DXF Loaded",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                DynamicFieldsGrid.Visibility = Visibility.Collapsed;
                SizeComboBox.Visibility = Visibility.Collapsed;
                PlacementGrid.Visibility = Visibility.Visible;
                GenerateButton.Visibility = Visibility.Visible;
            }
        }

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

            WriteBlock.ExecuteTask("Convert DXF Profile", () =>
            {
                // 2) Open the DXF in SpaceClaim so that Window.ActiveWindow is valid
                try
                {
                    SpaceClaim.Api.V242.Document.Open(dxfPath, null);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Failed to open DXF:\n{ex.Message}",
                        "DXF → Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                // 3) Build the profile + preview image
                DXFProfile profile = DXFImportHelper.DXFtoProfile();

                // 4) Close that DXF window
                SpaceClaim.Api.V242.Window.ActiveWindow?.Close();

                if (profile == null)
                {
                    // DXFtoProfile already showed an error 
                    return;
                }

                // 5) Copy the profile‐string to the clipboard
                System.Windows.Clipboard.SetText(profile.ProfileString);

                // 6) Let the user know that the string was copied
                System.Windows.MessageBox.Show(
                    $"DXF → Profile succeeded.\n\nName = {profile.Name}\n(Profile string copied to clipboard.)",
                    "DXF → Profile",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                //
                // 7) Create “C:\ProgramData\AESC_Construct\UserDXFProfiles” if it doesn’t exist
                //
                string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                // CommonApplicationData typically resolves to "C:\ProgramData"
                string userFolder = Path.Combine(programData, "AESC_Construct", "UserDXFProfiles");

                try
                {
                    Directory.CreateDirectory(userFolder);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Warning: Could not create folder:\n{userFolder}\n\n{ex.Message}",
                        "DXF → Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                //
                // 8) Decode Base64 preview and save it as a PNG in that folder
                //
                // Sanitize profile.Name for a valid filename
                string safeName = string.Join("_",
                    profile.Name.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)
                                .Select(tok => tok.Trim()));
                // Append a timestamp so multiple exports don’t collide
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
                    System.Windows.MessageBox.Show(
                        $"Failed to save preview image to disk:\n{ex.Message}",
                        "DXF → Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    // If saving PNG fails, omit the image path in CSV
                    imageFileName = "";
                }

                //
                // 9) Append (or create) “profiles.csv” in C:\ProgramData\AESC_Construct\UserDXFProfiles
                //
                string csvPath = Path.Combine(userFolder, "profiles.csv");
                bool writeHeader = !File.Exists(csvPath);

                try
                {
                    using (var sw = new StreamWriter(csvPath, append: true, encoding: System.Text.Encoding.UTF8))
                    {
                        if (writeHeader)
                        {
                            sw.WriteLine("Name;ProfileString;ImageRelativePath");
                        }

                        // Escape semicolons in the profile string so they don’t break our CSV
                        string escapedProfile = profile.ProfileString.Replace(";", "\\;");
                        string imageRel = string.IsNullOrEmpty(imageFileName) ? "" : imageFileName;

                        sw.WriteLine($"{profile.Name};{escapedProfile};{imageRel}");
                    }

                    //System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    //{
                        // Re-populate all user profiles from the updated CSV
                        LoadUserProfiles();

                        // Auto-check the newest one (so ProfileButton_Checked fires immediately)
                        var newestRb = UserProfilesGrid.Children
                                         .OfType<System.Windows.Controls.RadioButton>()
                                         .LastOrDefault();
                        if (newestRb != null)
                            newestRb.IsChecked = true;
                    //});
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Failed to update CSV:\n{ex.Message}",
                        "DXF → Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }

                //
                // 10) Finally, show the saved PNG in a WinForms window
                //
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
                        System.Windows.MessageBox.Show(
                            $"Failed to render saved preview image:\n{ex.Message}",
                            "DXF Preview Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
            });
        }

        private void LoadUserProfiles()
        {
            UserProfilesGrid.Children.Clear();

            try
            {
                // 1) Find folder & CSV
                var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                var userFolder = Path.Combine(programData, "AESC_Construct", "UserDXFProfiles");
                var csvPath = Path.Combine(userFolder, "profiles.csv");
                if (!File.Exists(csvPath))
                    return;

                // 2) Read lines, skip header
                var lines = File.ReadAllLines(csvPath).Skip(1);
                foreach (var line in lines)
                {
                    var raw = line.Split(';');
                    if (raw.Length < 3) continue;
                    var prof = new DXFProfile
                    {
                        Name = raw[0],
                        ProfileString = raw[1].Replace("\\;", ";"),
                        ImgString = raw[2]
                    };

                    // 3) Build the standard radio-content StackPanel
                    var contentStack = new StackPanel
                    {
                        Orientation = Orientation.Vertical,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                    };

                    if (!string.IsNullOrEmpty(prof.ImgString))
                    {
                        var imgPath = Path.Combine(userFolder, prof.ImgString);
                        if (File.Exists(imgPath))
                        {
                            var bmp = new BitmapImage(new Uri(imgPath));
                            contentStack.Children.Add(new System.Windows.Controls.Image
                            {
                                Source = bmp,
                                Width = 50,
                                Height = 50,
                                Margin = new Thickness(2)
                            });
                        }
                    }

                    contentStack.Children.Add(new TextBlock
                    {
                        Text = prof.Name,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        Margin = new Thickness(0, 2, 0, 0)
                    });

                    // 4) Declare the container Grid first so the delete handler can reference it
                    var container = new Grid
                    {
                        Width = 100,
                        Height = 100
                    };
                    container.ColumnDefinitions.Add(new ColumnDefinition());
                    container.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    // 5) Create the RadioButton as before
                    var userRb = new RadioButton
                    {
                        GroupName = "ProfileGroup",
                        Tag = prof,
                        Content = contentStack,
                        Width = 100,
                        Height = 100,
                        Margin = new Thickness(5)
                    };
                    userRb.Checked += ProfileButton_Checked;
                    Grid.SetColumn(userRb, 0);
                    container.Children.Add(userRb);

                    // 6) Create the delete “X” button
                    var delBtn = new Button
                    {
                        Content = "✕",
                        Width = 16,
                        Height = 16,
                        FontSize = 10,
                        VerticalAlignment = System.Windows.VerticalAlignment.Top,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                        Margin = new Thickness(0, 4, 4, 0),
                        BorderThickness = new Thickness(0),
                        ToolTip = "Remove this profile"
                    };
                    delBtn.Click += (s, e) =>
                    {
                        // remove from CSV
                        var updated = File.ReadAllLines(csvPath)
                                          .Where(l => !l.StartsWith(prof.Name + ";"))
                                          .ToArray();
                        File.WriteAllLines(csvPath, updated);

                        // remove from UI
                        UserProfilesGrid.Children.Remove(container);

                        // clear if it was selected
                        if (selectedProfileString == prof.ProfileString)
                            selectedProfileString = "";
                    };
                    Grid.SetColumn(delBtn, 1);
                    container.Children.Add(delBtn);

                    // 7) Add the container to your UniformGrid
                    UserProfilesGrid.Children.Add(container);
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"Error in LoadUserProfiles(): {ex}\n");
            }
        }



        /// <summary>
        /// Called when one of the “user‐profile” buttons is clicked:
        /// stores its ProfileString + ImgString, hides the numeric sliders,
        /// and shows the placement & generate buttons.
        /// </summary>
        private void UserProfileButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn
                && btn.Tag is DXFProfile prof)
            {
                // Clear any DXFContours (we’re using the baked profile string)
                selectedDXFPath = "";
                dxfContours = null;

                // Store the chosen profile‐string + image
                selectedProfileString = prof.ProfileString;
                selectedProfileImage = prof.ImgString;
                selectedProfile = prof.Name;

                File.AppendAllText(logPath, $"User profile clicked: {prof.Name}\n");

                // Hide the CSV / numeric fields entirely:
                DynamicFieldsGrid.Visibility = Visibility.Collapsed;
                SizeComboBox.Visibility = Visibility.Collapsed;
                HollowCheckBox.Visibility = Visibility.Collapsed;

                // Show the placement grid + Generate button:
                PlacementGrid.Visibility = Visibility.Visible;
                GenerateButton.Visibility = Visibility.Visible;
            }
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                // If nothing has been chosen, bail out:
                if (string.IsNullOrEmpty(selectedDXFPath)
                 && inputFieldMap.Count == 0
                 && string.IsNullOrEmpty(selectedProfileString))
                {
                    System.Windows.MessageBox.Show(
                        "Please select a profile (built-in or user-saved) or load a DXF file before generating.",
                        "Selection Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                bool isHollow = HollowCheckBox.IsChecked == false;
                bool updateBOM = UpdateBOM.IsChecked == true;
                Logger.Log($"updateBOM {updateBOM}");
                File.AppendAllText(logPath, $"AESCConstruct25: Generate Button Clicked.\n");

                //
                // ─── 1) USER‐SAVED PROFILE (selectedProfileString != "") ───────────────────────────────
                //
                // Log exactly the raw string (with no trailing “.”):
                File.AppendAllText(logPath, $"AESCConstruct25 custom: {selectedProfileString}\n");
                if (!string.IsNullOrEmpty(selectedProfileString))
                {
                    // ─── 1a) SPLIT INTO INDIVIDUAL CURVE STRINGS ───────────────────────────────────────
                    var curves = new List<ITrimmedCurve>();

                    // First split on '&' → each loop
                    string[] loopStrings = selectedProfileString.Split('&');
                    File.AppendAllText(logPath,
                        $"AESCConstruct25: Number of loops found = {loopStrings.Length}\n");

                    for (int loopIndex = 0; loopIndex < loopStrings.Length; loopIndex++)
                    {
                        string loopStr = loopStrings[loopIndex].Trim();
                        File.AppendAllText(logPath,
                            $"AESCConstruct25: Loop[{loopIndex}] = \"{loopStr}\"\n");

                        // Now split that loop on spaces → each “S…E…” or “S…E…M…” piece
                        string[] curveChunks = loopStr.Split(' ');
                        File.AppendAllText(logPath,
                            $"AESCConstruct25:   → Found {curveChunks.Length} curve‐chunks in this loop.\n");

                        for (int i = 0; i < curveChunks.Length; i++)
                        {
                            string curveStr = curveChunks[i].Trim();
                            if (string.IsNullOrEmpty(curveStr))
                                continue;

                            File.AppendAllText(logPath,
                                $"AESCConstruct25:     • curveStr[{i}] = \"{curveStr}\"\n");

                            try
                            {
                                ITrimmedCurve c = DXFImportHelper.CurveFromString(curveStr);
                                if (c == null)
                                {
                                    File.AppendAllText(logPath,
                                        $"AESCConstruct25 WARNING: CurveFromString returned NULL for “{curveStr}”\n");
                                }
                                else
                                {
                                    curves.Add(c);

                                    // Log the geometry type and key points:
                                    if (c is CurveSegment seg && seg.Geometry is Line)
                                    {
                                        var ps = seg.StartPoint;
                                        var pe = seg.EndPoint;
                                        File.AppendAllText(logPath,
                                            $"AESCConstruct25 LOG:     → Parsed LINE: Start=({ps.X:0.###},{ps.Y:0.###})  End=({pe.X:0.###},{pe.Y:0.###})\n");
                                    }
                                    else if (c is CurveSegment arcSeg && arcSeg.Geometry is Circle cir)
                                    {
                                        var center = cir.Axis.Origin;
                                        double radius = cir.Radius;
                                        File.AppendAllText(logPath,
                                            $"AESCConstruct25 LOG:     → Parsed CIRCLE/ARC: Center=({center.X:0.###},{center.Y:0.###})  Radius={radius:0.###}\n");
                                    }
                                    else
                                    {
                                        File.AppendAllText(logPath,
                                            $"AESCConstruct25 LOG:     → Parsed curve of unknown type (neither LINE nor CIRCLE).\n");
                                    }
                                }
                            }
                            catch (Exception exParse)
                            {
                                File.AppendAllText(logPath,
                                    $"AESCConstruct25 ERROR: Exception in CurveFromString(\"{curveStr}\") : {exParse}\n");
                            }
                        }
                    }

                    var firstSeg = (CurveSegment)curves[0];
                    File.AppendAllText(logPath, $"firstseg {firstSeg.StartPoint}, {firstSeg.EndPoint}\n");
                    var lastSeg = (CurveSegment)curves[curves.Count - 1];
                    File.AppendAllText(logPath, $"lastSeg {lastSeg.StartPoint}, {lastSeg.EndPoint}\n");
                    var diff = (firstSeg.StartPoint - lastSeg.EndPoint).Magnitude;
                    File.AppendAllText(logPath, $"  Loop closure gap = {diff:0.########} m\n");

                    File.AppendAllText(logPath,
                        $"AESCConstruct25: Total curves parsed = {curves.Count}\n");

                    if (curves.Count == 0)
                    {
                        System.Windows.MessageBox.Show(
                            "Could not reconstruct any curves from the saved profile string.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        return;
                    }

                    // ─── 1b) COMPUTE BOUNDING‐BOX (width, height) ─────────────────────────────────────────
                    var (profW, profH) = DXFImportHelper.GetDXFSize(curves);
                    double w = profW;
                    double h = profH;
                    File.AppendAllText(logPath,
                        $"AESCConstruct25: Bounding box from GetDXFSize → width = {w:0.####}, height = {h:0.####}\n");

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

                    File.AppendAllText(logPath,
                        $"AESCConstruct25: (User-saved) Applying Offset X={offsetX:0.####}, Y={offsetY:0.####}\n");

                    foreach (var c in curves)
                    {
                        if (c is CurveSegment seg)
                        {
                            var ps = seg.StartPoint;
                            var pe = seg.EndPoint;
                            File.AppendAllText(logPath,
                                $"  → CSV curve: Start=({ps.X:0.######},{ps.Y:0.######},{ps.Z:0.######})  " +
                                $"End=({pe.X:0.######},{pe.Y:0.######},{pe.Z:0.######})\n");
                        }
                    }
                    // ─── 1d) EXTRUDE THE CURVES ───────────────────────────────────────────────────────────
                    try
                    {
                        WriteBlock.ExecuteTask("Generate Profile (user-saved)", () =>
                        {
                            ExtrudeProfileCommand.ExecuteExtrusion(
                                "CSV",      // e.g. “square”
                                curves,            // List<ITrimmedCurve>
                                false,             // always solid for user-saved
                                offsetX,
                                offsetY,
                                "",       // no DXF file path in this branch
                                updateBOM
                            );
                        });
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(logPath,
                            $"AESCConstruct25 ERROR extruding user-saved “{selectedProfile}”: {ex}\n");
                        System.Windows.MessageBox.Show(
                            "An error occurred while extruding the user-saved profile.\nSee log on desktop for details.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
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

                    File.AppendAllText(logPath,
                        $"AESCConstruct25: (DXF) Applying Offset X={offsetX:0.####}, Y={offsetY:0.####}\n");

                    try
                    {
                        WriteBlock.ExecuteTask("Generate Profile (DXF)", () =>
                        {
                            ExtrudeProfileCommand.ExecuteExtrusion(
                                selectedProfile,   // “DXF”
                                dxfContours,       // List<ITrimmedCurve>
                                false,             // always solid for imported DXF
                                offsetX,
                                offsetY,
                                selectedDXFPath,    // pass the original .dxf path
                                updateBOM
                            );
                        });
                    }
                    catch (Exception exDxf)
                    {
                        File.AppendAllText(logPath,
                            $"AESCConstruct25 ERROR extruding DXF contours: {exDxf}\n");
                        System.Windows.MessageBox.Show(
                            "An error occurred while extruding the DXF contours.\nSee log on desktop for details.",
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                    return;
                }

                //
                // ─── 3) BUILT‐IN SHAPES (Rectangular, Circular, H, L, U, T) ───────────────────────────────
                //
                double wBuilt = 0, hBuilt = 0;
                if (selectedProfile == "Circular")
                {
                    wBuilt = inputFieldMap.ContainsKey("D") ? double.Parse(inputFieldMap["D"].Text) / 1000 : 0.0;
                    hBuilt = wBuilt;
                }
                else if (selectedProfile == "L")
                {
                    wBuilt = inputFieldMap.ContainsKey("b") ? double.Parse(inputFieldMap["b"].Text) / 1000 : 0.0;
                    hBuilt = inputFieldMap.ContainsKey("a") ? double.Parse(inputFieldMap["a"].Text) / 1000 : 0.0;
                }
                else
                {
                    wBuilt = inputFieldMap.ContainsKey("w") ? double.Parse(inputFieldMap["w"].Text) / 1000 : 0.0;
                    hBuilt = inputFieldMap.ContainsKey("h") ? double.Parse(inputFieldMap["h"].Text) / 1000 : 0.0;
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

                File.AppendAllText(logPath,
                    $"AESCConstruct25: (Built-in) Applying Offset X={offsetXB:0.####}, Y={offsetYB:0.####}\n");
                bool isHollowBuilt = HollowCheckBox.IsChecked == false;

                try
                {
                    var dataDict = inputFieldMap
                    .ToDictionary(
                        kvp => kvp.Key,                // e.g. "w", "h", "t", etc.
                        kvp => kvp.Value.Text.Trim()   // the string value the user typed
                    );
                
                    WriteBlock.ExecuteTask("Generate Profile (built-in)", () =>
                    {
                        ExtrudeProfileCommand.ExecuteExtrusion(
                            selectedProfile,  // e.g. "Rectangular", "H", etc.
                            dataDict,    // numeric‐based sizes
                            isHollowBuilt,
                            offsetXB,
                            offsetYB,
                            "",
                            updateBOM
                        );
                    });
                }
                catch (Exception exBuiltIn)
                {
                    File.AppendAllText(logPath,
                        $"AESCConstruct25 ERROR extruding built‐in “{selectedProfile}”: {exBuiltIn}\n");
                    System.Windows.MessageBox.Show(
                        "An error occurred while extruding the built‐in profile.\nSee log on desktop for details.",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            finally
            {
                // restore whatever the user had selected
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        private void HollowCheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }
    }
}
