//private static string _keyPair =
//            "<RSAKeyValue><Modulus>pR07DDlVi/rcY1XeNkdFHdXEtbkk9zOBNx9MA+PwGMOMHfeA6c3cqFizdt/pcjR+p7SPwTP6L/K6DO6asU0KeSEoBwZZaTQ/UJyp4T/xJtrHpThJX5XaIf35ebp9zV1ETiYik+C0HbPVpZVrCZjRnf9waLRsO5UtTnZDQn8yvY0vtz7OlWBdHtTBO0EKmd+gXvR1K2pML/R2LLu4mpwbpv7L3p4O6/xaEEPvGuHyKHeK6B3+IW5bo94DuzG6OCi9r10xQ4N/ZT6/3WJVp6CfrzZtUd8/vYh7EiBqeYvKZm9jcgEfiG1nUY7Y1ntX6dtjSTFT9k6Ne2ZcybVUPWMJGw==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

using LicenseSpot.Framework;
using SpaceClaim.Api.V242; // for Command (network toggle UI)
using System;
using System.Drawing;
using System.IO;
using System.Windows;

namespace AESCConstruct25.Licensing
{
    /// <summary>
    /// Central license integration. Network licenses are considered "valid" only when checked-out.
    /// </summary>
    public static class ConstructLicenseSpot
    {
        private sealed class Construct25Anchor { }
        public static bool IsDebugMode { get; set; } = false;
        private static void Debug(string m) { if (IsDebugMode) MessageBox.Show(m, "License Debug"); }

        private static readonly Type AnchorType = typeof(Construct25Anchor);
        private static ExtendedLicense _license;
        public static ExtendedLicense CurrentLicense => _license;

        // Put your real RSA public key here
        private static string _keyPair =
            "<RSAKeyValue><Modulus>pR07DDlVi/rcY1XeNkdFHdXEtbkk9zOBNx9MA+PwGMOMHfeA6c3cqFizdt/pcjR+p7SPwTP6L/K6DO6asU0KeSEoBwZZaTQ/UJyp4T/xJtrHpThJX5XaIf35ebp9zV1ETiYik+C0HbPVpZVrCZjRnf9waLRsO5UtTnZDQn8yvY0vtz7OlWBdHtTBO0EKmd+gXvR1K2pML/R2LLu4mpwbpv7L3p4O6/xaEEPvGuHyKHeK6B3+IW5bo94DuzG6OCi9r10xQ4N/ZT6/3WJVp6CfrzZtUd8/vYh7EiBqeYvKZm9jcgEfiG1nUY7Y1ntX6dtjSTFT9k6Ne2ZcybVUPWMJGw==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        // Optional product filter (empty = disabled)
        private static string LicenseDescription = "";

        // Preferred license file
        private static readonly string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        private static readonly string BaseDir = Path.Combine(ProgramData, "IPManager");
        public static readonly string LicenseFilePath = Path.Combine(BaseDir, "AESC_License_Construct25.lic");

        static Bitmap _active = new Bitmap(new MemoryStream(Resources.Network_Active));
        static Bitmap _inactive = new Bitmap(new MemoryStream(Resources.Network_Inactive));

        // Public state for UI gating
        public static bool IsValid { get; private set; } = false;     // true only if: (node-locked genuine) OR (network genuine + checked-out)
        public static bool IsNetwork { get; private set; } = false;
        public static string Status { get; private set; } = "License is not valid.";

        // ─────────────────────────────────────────────────────────────────────

        public static bool CheckLicense()
        {
            try
            {
                EnsureProgramDataFolder();

                if (!File.Exists(LicenseFilePath))
                {
                    IsValid = false; IsNetwork = false;
                    Status = "License file not found. Please activate.";
                    return false;
                }

                var fi = new FileInfo(LicenseFilePath);
                if (fi.Length < 16)
                {
                    IsValid = false; IsNetwork = false;
                    Status = "License file appears incomplete. Try activating again.";
                    return false;
                }

                _license = OpenStrict(LicenseFilePath);
                if (_license == null)
                {
                    IsValid = false; IsNetwork = false;
                    Status = "No license found or invalid license file.";
                    return false;
                }

                _license.Refresh();
                IsNetwork = _license.IsNetwork;

                bool g1 = _license.IsGenuineEx(0) == GenuineResult.Genuine;
                bool g2 = _license.IsGenuine(true, _keyPair);

                bool descOk = true;
                if (!string.IsNullOrEmpty(LicenseDescription))
                {
                    try
                    {
                        var desc = _license.GetProperty("Description")?.ToString() ?? "";
                        descOk = string.Equals(desc, LicenseDescription, StringComparison.OrdinalIgnoreCase);
                    }
                    catch { descOk = false; }
                }

                if (g1 && g2 && descOk)
                {
                    if (IsNetwork)
                    {
                        bool connected = _license.IsValidConnection();
                        IsValid = connected;
                        Status = connected ? "Valid license." : "Valid license, not activated.";
                    }
                    else
                    {
                        IsValid = true;
                        Status = "Valid license.";
                    }
                }
                else
                {
                    IsValid = false;
                    Status = "License is not valid.";
                }

                if (!IsValid)
                {
                    //var exp = _license.GetTimeLimit()?.EndDate;
                    //string SafeProp(string k) { try { return _license.GetProperty(k)?.ToString() ?? ""; } catch { return ""; } }
                }

                UpdateNetworkButtonUI();
                return IsValid;
            }
            catch (Exception ex)
            {
                IsValid = false; IsNetwork = false;
                Status = "CheckLicense error: " + ex.Message;
                return false;
            }
        }

        public static bool TryActivate(string serialNumber, out string message)
        {
            if (string.IsNullOrWhiteSpace(serialNumber))
            {
                message = "No activation code provided.";
                return false;
            }

            try
            {
                EnsureProgramDataFolder();

                // Create a clean handle bound only by key for activation.
                var handle = ExtendedLicenseManager.GetLicense(AnchorType, _keyPair) as ExtendedLicense;
                if (handle == null)
                {
                    message = "Could not create license handle.";
                    return false;
                }

                try { handle.Deactivate(); } catch { /* ignore */ }

                // LicenseSpot saves to ProgramData\IPManager with this stem.
                var stem = Path.GetFileNameWithoutExtension(LicenseFilePath);
                handle.Activate(serialNumber.Trim(), saveFile: true, fileName: stem);

                WaitUntilFileLooksWritten();

                // Now reopen STRICTLY from our exact file so other .lic files cannot interfere.
                _license = OpenStrict(LicenseFilePath);
                var ok = CheckLicense();
                message = Status;
                return ok;
            }
            catch (Exception ex)
            {
                IsValid = false;
                IsNetwork = false;
                Status = "Activation error: " + ex.Message;
                message = Status;
                return false;
            }
        }



        // Call this right after startup to ensure a network license is not consuming a seat until user clicks the button.
        public static void EnsureNetworkDeactivatedOnStartup()
        {
            try
            {
                if (_license != null && _license.IsNetwork && _license.IsValidConnection())
                {
                    _license.CheckIn();
                    IsValid = false;
                    Status = "Valid license, not activated.";
                    UpdateNetworkButtonUI();
                    Debug("[Startup] Network license was connected – checked in to keep it deactivated until user action.");
                }
            }
            catch
            {
                // ignore; if check-in fails we leave state as-is
            }
        }

        // ── Network toggles wired to the ribbon ──────────────────────────────

        public static void licenseCheckOut()
        {
            if (_license == null || !_license.IsNetwork) return;

            try
            {
                _license.CheckOut();
                if (_license.IsValidConnection())
                {
                    IsValid = true;
                    Status = "Network license activated.";
                }
                else
                {
                    IsValid = false;
                    Status = "Could not activate the network license (already connected?).";
                }
            }
            catch (Exception ex)
            {
                IsValid = false;
                Status = "CheckOut error: " + ex.Message;
            }
            UpdateNetworkButtonUI();
        }

        public static void licenseCheckIn()
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

        // ── Ribbon button state ──────────────────────────────────────────────

        public static void UpdateNetworkButtonUI()
        {
            try
            {
                var cmd = Command.GetCommand("AESCConstruct25.ActivateNetwork");
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
            catch { /* ignore */ }
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static void EnsureProgramDataFolder()
        {
            try { Directory.CreateDirectory(BaseDir); } catch { }
        }

        private static ExtendedLicense OpenStrict(string filePath)
        {
            try
            {
                var vinfo = new LicenseValidationInfo { LicenseFile = new LicenseFile(filePath) };
                return ExtendedLicenseManager.GetLicense(AnchorType, instance: null, info: vinfo, publicKey: _keyPair) as ExtendedLicense;
            }
            catch
            {
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
