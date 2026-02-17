using LicenseSpot.Framework;
using SpaceClaim.Api.V242; // for Command (network toggle UI)
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows;
using AESCConstruct2026.FrameGenerator.Utilities;

namespace AESCConstruct2026.Licensing
{
    /// <summary>
    /// Central license integration. Network licenses are considered "valid" only when checked-out.
    /// </summary>
    public static class ConstructLicenseSpot
    {
        private sealed class Construct2026Anchor { }
        public static bool IsDebugMode { get; set; } = false;
        private static void Debug(string m) { if (IsDebugMode) MessageBox.Show(m, "License Debug"); }

        private static readonly Type AnchorType = typeof(Construct2026Anchor);
        private static readonly object _lock = new object();
        private static ExtendedLicense _license;
        public static ExtendedLicense CurrentLicense => _license;

        // Put your real RSA public key here
        private static string _keyPair =
            "<RSAKeyValue><Modulus>zx7FQo4ZybNKG3wZfxkMwMJ2PRlejQ4CjDELNGc5BiE4Wk30MO84r/oYhfzZBNPU5iHVAxsy/PLWzWm3f0Fl+0T95UJ5m3Xgm8SrM9U3O/omDUJDH9U8+ve8lJEXYGQQz+ga3ydOfQ0qMqUOwjqhVX8REuXOF7xPbUcBPWLhCLlwNZWfbPmCzpZeCBfxTguGQbHLD+LcHAfFh9QMVeYIwi4kUHtd/r6cUNmb/2/b2pt11Bt8iyq9DkBBuMrWNpHI9VP+ALp7s+x5RzDlVPoi6QUVhHi1EJuKAtYmFIrLUw08AGVJuVUHHdzhXNi/zIZG/K/2vwNuWNDD3SHBcz3yMQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        // Optional product filter (empty = disabled)
        private static string LicenseDescription = "";

        // Preferred license file
        private static readonly string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        private static readonly string BaseDir = Path.Combine(ProgramData, "IPManager");
        public static readonly string LicenseFilePath = Path.Combine(BaseDir, "AESC_License_Construct2026.lic");

        // Streams must stay alive for the lifetime of the Bitmap (GDI+ requirement).
        // These are static app-lifetime objects, so no disposal is needed.
        private static readonly MemoryStream _activeStream = new MemoryStream(Resources.Network_Active);
        private static readonly MemoryStream _inactiveStream = new MemoryStream(Resources.Network_Inactive);
        static Bitmap _active = new Bitmap(_activeStream);
        static Bitmap _inactive = new Bitmap(_inactiveStream);

        // Public state for UI gating
        public static bool IsValid { get; private set; } = false;     // true only if: (node-locked genuine) OR (network genuine + checked-out)
        public static bool IsNetwork { get; private set; } = false;
        public static string Status { get; private set; } = "License is not valid.";

        // ─────────────────────────────────────────────────────────────────────

        public static bool CheckLicense()
        {
            lock (_lock)
            {
                var sw = Stopwatch.StartNew();
                Logger.Log("[License] CheckLicense start");
                try
                {
                    EnsureProgramDataFolder();

                    Logger.Log($"[License] BaseDir='{BaseDir}'");
                    Logger.Log($"[License] LicenseFilePath='{LicenseFilePath}'");

                    if (!File.Exists(LicenseFilePath))
                    {
                        IsValid = false; IsNetwork = false;
                        Status = "License file not found. Please activate.";
                        Logger.Log("[License] File not found");
                        return false;
                    }

                    try
                    {
                        var fi = new FileInfo(LicenseFilePath);
                        Logger.Log($"[License] File size={fi.Length} bytes, LastWriteUtc={fi.LastWriteTimeUtc:o}");
                        if (fi.Length < 16)
                        {
                            IsValid = false; IsNetwork = false;
                            Status = "License file appears incomplete. Try activating again.";
                            Logger.Log("[License] File too small");
                            return false;
                        }
                        //string sha = ComputeSha256(LicenseFilePath);
                        //Logger.Log($"[License] File SHA256={sha}");
                    }
                    catch (Exception exMeta)
                    {
                        Logger.Log($"[License] File meta read error: {exMeta.Message}");
                    }

                    _license = OpenStrict(LicenseFilePath);
                    if (_license == null)
                    {
                        IsValid = false; IsNetwork = false;
                        Status = "No license found or invalid license file.";
                        Logger.Log("[License] OpenStrict returned null");
                        return false;
                    }

                    Logger.Log("[License] Refresh()");
                    _license.Refresh();
                    IsNetwork = _license.IsNetwork;
                    Logger.Log($"[License] IsNetwork={IsNetwork}");

                    bool g1 = false, g2 = false;
                    try { g1 = _license.IsGenuineEx(0) == GenuineResult.Genuine; } catch (Exception ex) { Logger.Log($"[License] IsGenuineEx error: {ex.Message}"); }
                    try { g2 = _license.IsGenuine(true, _keyPair); } catch (Exception ex) { Logger.Log($"[License] IsGenuine error: {ex.Message}"); }
                    Logger.Log($"[License] Genuine checks g1={g1}, g2={g2}");

                    bool descOk = true;
                    if (!string.IsNullOrEmpty(LicenseDescription))
                    {
                        try
                        {
                            var desc = _license.GetProperty("Description")?.ToString() ?? "";
                            descOk = string.Equals(desc, LicenseDescription, StringComparison.OrdinalIgnoreCase);
                            Logger.Log($"[License] Description stored='{desc}', expected='{LicenseDescription}', match={descOk}");
                        }
                        catch (Exception ex)
                        {
                            descOk = false;
                            Logger.Log($"[License] Description read error: {ex.Message}");
                        }
                    }

                    if (g1 && g2 && descOk)
                    {
                        if (IsNetwork)
                        {
                            bool connected = false;
                            try { connected = _license.IsValidConnection(); } catch (Exception ex) { Logger.Log($"[License] IsValidConnection error: {ex.Message}"); }
                            IsValid = connected;
                            Status = connected ? "Valid license." : "Valid license, not activated.";
                            Logger.Log($"[License] Network connection valid={connected}");
                        }
                        else
                        {
                            IsValid = true;
                            Status = "Valid license.";
                            Logger.Log("[License] Local license valid");
                        }
                    }
                    else
                    {
                        IsValid = false;
                        Status = "License is not valid.";
                        Logger.Log("[License] Validity failed");
                    }

                    UpdateNetworkButtonUI();
                    Logger.Log($"[License] CheckLicense end result IsValid={IsValid} in {sw.ElapsedMilliseconds} ms");
                    return IsValid;
                }
                catch (Exception ex)
                {
                    IsValid = false; IsNetwork = false;
                    Status = "CheckLicense error: " + ex.Message;
                    Logger.Log($"[License] CheckLicense exception: {ex}");
                    return false;
                }
            }
        }

        public static bool TryActivate(string serialNumber, out string message)
        {
            lock (_lock)
            {
                var sw = Stopwatch.StartNew();
                Logger.Log("[License] TryActivate start");
                if (string.IsNullOrWhiteSpace(serialNumber))
                {
                    message = "No activation code provided.";
                    Logger.Log("[License] Activation aborted empty serial");
                    return false;
                }

                try
                {
                    EnsureProgramDataFolder();

                    var handle = ExtendedLicenseManager.GetLicense(AnchorType, _keyPair) as ExtendedLicense;
                    if (handle == null)
                    {
                        message = "Could not create license handle.";
                        Logger.Log("[License] GetLicense returned null");
                        return false;
                    }

                    try { Logger.Log("[License] Deactivate existing if any"); handle.Deactivate(); } catch (Exception ex) { Logger.Log($"[License] Deactivate error: {ex.Message}"); }

                    var stem = Path.GetFileNameWithoutExtension(LicenseFilePath);
                    Logger.Log($"[License] Activate to stem='{stem}', path dir='{Path.GetDirectoryName(LicenseFilePath)}'");
                    handle.Activate(serialNumber.Trim(), saveFile: true, fileName: stem);

                    Logger.Log("[License] Waiting for file to appear and stabilize");
                    WaitUntilFileLooksWritten();

                    _license = OpenStrict(LicenseFilePath);

                    var ok = CheckLicense(); // also logs — re-entrant lock is OK
                    message = Status;
                    Logger.Log($"[License] TryActivate end ok={ok} in {sw.ElapsedMilliseconds} ms, status='{Status}'");
                    return ok;
                }
                catch (Exception ex)
                {
                    IsValid = false;
                    IsNetwork = false;
                    Status = "Activation error: " + ex.Message;
                    message = Status;
                    Logger.Log($"[License] TryActivate exception: {ex}");
                    return false;
                }
            }
        }

        // Call this right after startup to ensure a network license is not consuming a seat until user clicks the button.
        public static void EnsureNetworkDeactivatedOnStartup()
        {
            lock (_lock)
            {
                Logger.Log("[License] EnsureNetworkDeactivatedOnStartup");
                try
                {
                    if (_license != null && _license.IsNetwork && _license.IsValidConnection())
                    {
                        Logger.Log("[License] Network was connected at startup, checking in");
                        _license.CheckIn();
                        IsValid = false;
                        Status = "Valid license, not activated.";
                        UpdateNetworkButtonUI();
                        Debug("[Startup] Network license was connected – checked in to keep it deactivated until user action.");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[License] EnsureNetworkDeactivatedOnStartup error: {ex.Message}");
                    // ignore; if check in fails we leave state as is
                }
            }
        }

        // ── Network toggles wired to the ribbon ──────────────────────────────

        public static void LicenseCheckOut()
        {
            lock (_lock)
            {
                Logger.Log("[License] CheckOut requested");
                if (_license == null || !_license.IsNetwork)
                {
                    Logger.Log("[License] CheckOut aborted no network license handle");
                    return;
                }

                try
                {
                    _license.CheckOut();
                    bool connected = _license.IsValidConnection();
                    IsValid = connected;
                    Status = connected ? "Network license activated." : "Could not activate the network license (already connected?).";
                    Logger.Log($"[License] CheckOut result connected={connected}");
                }
                catch (Exception ex)
                {
                    IsValid = false;
                    Status = "CheckOut error: " + ex.Message;
                    Logger.Log($"[License] CheckOut exception: {ex}");
                }
                UpdateNetworkButtonUI();
            }
        }

        public static void LicenseCheckIn()
        {
            lock (_lock)
            {
                if (_license == null || !_license.IsNetwork) return;

                try
                {
                    if (_license.IsValidConnection())
                    {
                        _license.CheckIn();
                    }
                    IsValid = false;
                    Status = "Valid license, not activated.";
                }
                catch (Exception ex)
                {
                    Status = "CheckIn error: " + ex.Message;
                }
                UpdateNetworkButtonUI();
            }
        }

        // ── Ribbon button state ──────────────────────────────────────────────

        public static void UpdateNetworkButtonUI()
        {
            try
            {
                var cmd = Command.GetCommand("AESCConstruct2026.ActivateNetwork");
                if (cmd == null) return;

                var lic = _license;
                bool hasNet = lic != null && lic.IsNetwork;
                bool connected = hasNet && lic.IsValidConnection();

                cmd.IsEnabled = hasNet;
                cmd.Text = connected ? "Deactivate network license" : "Activate network license";

                // Optional: swap icons if you prefer (keep or remove as needed)
                 cmd.Image = connected ? _active
                                       : _inactive;
            }
            catch (Exception ex) { Logger.Log("[License] UpdateNetworkButtonUI failed: " + ex.Message); }
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static void EnsureProgramDataFolder()
        {
            try { Directory.CreateDirectory(BaseDir); } catch (Exception ex) { Logger.Log("[License] EnsureProgramDataFolder failed: " + ex.Message); }
        }

        private static ExtendedLicense OpenStrict(string filePath)
        {
            try
            {
                var vinfo = new LicenseValidationInfo { LicenseFile = new LicenseFile(filePath) };
                return ExtendedLicenseManager.GetLicense(AnchorType, instance: null, info: vinfo, publicKey: _keyPair) as ExtendedLicense;
            }
            catch (Exception ex)
            {
                Logger.Log("[License] OpenStrict failed: " + ex.Message);
                return null;
            }
        }

        private static void WaitUntilFileLooksWritten(int attempts = 10, int delayMs = 100)
        {
            for (int i = 0; i < attempts; i++)
            {
                try
                {
                    if (File.Exists(LicenseFilePath) && new FileInfo(LicenseFilePath).Length > 32) return;
                }
                catch { }
                System.Threading.Thread.Sleep(delayMs);
            }
        }
    }
}
