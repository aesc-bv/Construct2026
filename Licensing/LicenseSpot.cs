//using LicenseSpot.Framework;
//using SpaceClaim.Api.V242;
//using System;
//using System.IO;
//using System.Windows;
//using AESCConstruct25.FrameGenerator.Utilities;
//using AESCConstruct25.FrameGenerator.UI;

//namespace AESCConstruct25.LicenseSpot
//{
//    internal class LicenseSpot
//    {
//        public static ExtendedLicense license;
//        public static string _addinPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\SpaceClaim\\Addins\\AESCConstruct";
//        public static string keyPair = "<RSAKeyValue><Modulus>y1bZeGeXfpp5NEm4S34PmT5F0uAouchu30MnWUYNxhqjrs2BNWvMeYDrT0Fqw6JhUchv7OFduwTxzOS7V8yajkCgvbo0OnxYF4Qgf46xoTE0kYBDeuXLnASPfoeo+JfQ5M81IhfwYsMD1i6pvIlODl5KK/8ydqNs3bcknUF/DGDeOa3eEyCoxnFMzu+t2rdENKUK5P403l6K3R2WpfzN+xHAyqs6OUi79zdzPoSkU/X5+ZFO5ZQiO3L4Qn/Yc2Pgn4GrqbpQ6QJVpd9n3kioC9LFs8lOKqHwlz5G+Ob96quhn7aswFKbFZQAXiHF+vKhUENzWnMywtEhp+bn4rqOaQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";
//        public static string licenseFolder = _addinPath + "\\AESCCalculateLicense.lic";
//        private static DateTime LicenseExpiration = DateTime.Parse("01/09/2025");

//        public static class _License
//        {
//            public static string status = "License is not valid.";
//            public static Boolean valid = false;
//            public static Boolean genuine = false;
//            public static Boolean isNetwork = false;
//        }

//        public static void activateLicense()
//        {
//           // Logger.Log("activate");
//            var info = new LicenseValidationInfo
//            {
//                LicenseFile = new LicenseFile(LicenseSpot.LicenseSpot.licenseFolder),
//                SerialNumber = Settings.Default.SerialNumber
//            };
//            LicenseSpot.LicenseSpot.license =
//                ExtendedLicenseManager.GetLicense(typeof(SettingsControl),
//                                                 null, info, LicenseSpot.LicenseSpot.keyPair);
//            LicenseSpot.LicenseSpot._License.valid = LicenseSpot.LicenseSpot.license.Validate();
//        }

//        //public static void licenseCheckIn()
//        //{
//        //    MessageBox.Show("licenseCheckIn\nlicense.IsNetwork: " + license.IsNetwork.ToString() + "\nlicense.IsValidConnection(): " + license.IsValidConnection().ToString());
//        //    if (license.IsNetwork)// && license.IsValidConnection())
//        //    {
//        //        if (license.IsValidConnection())
//        //            license.CheckIn();
//        //        LicenseSpot._License.valid = false;
//        //        _License.status = "Valid license, not activated.";
//        //        //if (Properties.Settings.Default.NetworkLogUser_Enabled)
//        //        //{
//        //        //Logger.Log(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "\tCheckIn  \t" + license.HardwareID);
//        //        //MessageBox.Show("Check licenseCheckIn: Succes");
//        //        //}

//        //        //update button activate
//        //        try
//        //        {
//        //            //Command com = Command.GetCommand("AESC.Calculate.LicenseNetworkCmd");
//        //            //com.Text = Language.tl("LicenseNetwork_Activate");
//        //            //com.Hint = Language.tl("LicenseNetwork_Activate_Tooltip");
//        //            //com.Image = Properties.Resources.icons8_next_36px;
//        //        }
//        //        catch { }

//        //    }
//        //}


//        //public static string GetHardWareID()
//        //{
//        //    return license.HardwareID;
//        //}

//        //public static void licenseCheckOut()
//        //{
//        //    try
//        //    {
//        //        license.CheckOut();
//        //        if (!license.IsValidConnection())
//        //        {
//        //            MessageBox.Show("Could not activate the Calculate license, it is already connected.");
//        //            return;
//        //        }

//        //        LicenseSpot._License.valid = true;
//        //        _License.status = "Valid license";

//        //        //checkModules();

//        //        //if (Properties.Settings.Default.NetworkLogUser_Enabled)
//        //        //{
//        //            //Logger.Log(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "\tCheckout\t" + license.HardwareID);
//        //        //}
//        //        //Functions.ShowCustomMessageBox("License Check Out Succes");

//        //        try
//        //        {
//        //            //update button activate
//        //            Command com = Command.GetCommand("AESC.Calculate.LicenseNetworkCmd");
//        //            //com.Text = Language.tl("LicenseNetwork_Deactivate");
//        //            //com.Hint = Language.tl("LicenseNetwork_Deactivate_Tooltip");
//        //            //com.Image = Properties.Resources.icons8_pause_squared_36px;
//        //        }
//        //        catch { }


//        //    }
//        //    catch
//        //    {
//        //        //Functions.ShowCustomMessageBox("Could not activate the Calculate license, please check if it is already connected.");
//        //        LicenseSpot._License.valid = false;
//        //    }

//        //}

//        public static void checkLicense()
//        {
//            //LicenseValidationInfo info = new LicenseValidationInfo();
//            //info.LicenseFile = new LicenseFile(licenseFolder);
//            //license = ExtendedLicenseManager.GetLicense(typeof(AESCConstruct25.UIMain.UIManager), null, info, keyPair);
//            //LicenseSpot._License.isNetwork = license.IsNetwork;
//           // Logger.Log("// Logger.Log(\"AESCConstruct25:Licensing started\");");
//            //try
//            //{
//            //    license.Refresh();
//            //    //check valid license
//            //    bool genuine = license.IsGenuineEx(0) == GenuineResult.Genuine && license.IsGenuine(true, keyPair);
//            //    LicenseSpot._License.genuine = genuine;
//            //    if (!license.IsNetwork)
//            //        LicenseSpot._License.valid = genuine;
//            //    MessageBox.Show("license.IsGenuineEx(0)\t" + license.IsGenuineEx(0).ToString() + "\nGenuineResult.Genuine\t" + GenuineResult.Genuine.ToString() + "\nlicense.IsGenuine(true, keyPair)\t" + license.IsGenuine(true, keyPair).ToString() + "\n_License.valid\t" + LicenseSpot._License.valid.ToString()) ;
//            //}
//            //catch
//            //{
//            //    _License.status = "Could not connect with the internet";
//            //    LicenseSpot._License.valid = false;
//            //}



//            //// Write to log if needed
//            //if (license.IsNetwork && license.IsValidConnection())
//            //    licenseCheckIn();


//            //if (license.IsNetwork && LicenseSpot._License.genuine)
//            //{
//            //    //if (Properties.Settings.Default.AutomaticCheckIn_FloatingLicense)
//            //    //{
//            //        licenseCheckOut();
//            //    //}
//            //    //else
//            //    //    _License.status = "Valid license, not activated.";
//            //}


//            //if (_License.valid)
//            //{
//            //    _License.status = "Valid license";
//            //    //checkModules();
//            //}
//            //else if (DateTime.Compare(LicenseExpiration, DateTime.Now) > 0)
//            //{
//            //   // Logger.Log("Trial version expires in " + (LicenseExpiration - DateTime.Now).Days.ToString() + "days");
//            //    LicenseSpot._License.valid = true;
//            //}
//        }

//    }
//}
using AESCConstruct25.Properties;
using LicenseSpot.Framework;
using System;

namespace AESCConstruct25.LicenseSpot
{
    internal static class LicenseSpot
    {
        public static ExtendedLicense License;
        public static class State
        {
            public static bool Valid { get; internal set; }
            public static bool IsNetwork { get; internal set; }
            public static string Status { get; internal set; }
        }

        private static readonly string KeyPair = "<RSAKeyValue><Modulus>pR07DDlVi/rcY1XeNkdFHdXEtbkk9zOBNx9MA+PwGMOMHfeA6c3cqFizdt/pcjR+p7SPwTP6L/K6DO6asU0KeSEoBwZZaTQ/UJyp4T/xJtrHpThJX5XaIf35ebp9zV1ETiYik+C0HbPVpZVrCZjRnf9waLRsO5UtTnZDQn8yvY0vtz7OlWBdHtTBO0EKmd+gXvR1K2pML/R2LLu4mpwbpv7L3p4O6/xaEEPvGuHyKHeK6B3+IW5bo94DuzG6OCi9r10xQ4N/ZT6/3WJVp6CfrzZtUd8/vYh7EiBqeYvKZm9jcgEfiG1nUY7Y1ntX6dtjSTFT9k6Ne2ZcybVUPWMJGw==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        /// <summary>
        /// Reads the saved serial number from settings, tries to activate or validate,
        /// and populates State.Valid/Status.
        /// </summary>
        public static void Initialize()
        {
            // Logger.Log("LicenseSpot: Initializing license");
            var serial = Settings.Default.SerialNumber?.Trim();
            if (string.IsNullOrEmpty(serial))
            {
                State.Valid = false;
                State.Status = "No serial entered.";
                return;
            }

            try
            {
                // This will look in ProgramData\IPManager\license.lic (default) or App_Data, etc.
                License = ExtendedLicenseManager.GetLicense(typeof(LicenseSpot), null, KeyPair);
                State.IsNetwork = License.IsNetwork;
                License.Refresh();

                var result = License.IsGenuineEx(0);
                State.Valid = result == GenuineResult.Genuine;
                State.Status = State.Valid
                  ? "License is valid."
                  : $"License invalid: {License.InvalidReason}";
            }
            catch (Exception ex)
            {
                State.Valid = false;
                State.Status = "License check failed: " + ex.Message;
                // Logger.Log("LicenseSpot: exception in Initialize(): " + ex);
            }
        }



        /// <summary>
        /// Called from Settings “Activate” button to save & re-check.
        /// </summary>
        public static void Activate(string serial)
        {
            // 1) save to user settings
            Settings.Default.SerialNumber = serial;
            Settings.Default.Save();
            // Logger.Log($"LicenseSpot: activating serial “{serial}”");

            try
            {
                // 2) get a fresh ExtendedLicense instance (no file must exist yet)
                License = ExtendedLicenseManager.GetLicense(typeof(LicenseSpot), null);

                // 3) actually activate ⇒ downloads & writes the .lic for us
                //    the return value is the raw xml, if you ever need it
                string licenseXml = License.Activate(
                  serial,                // serial number
                  saveFile: true,      // write it out
                  fileName: "AESCCalculateLicense.lic"
                );

                // 4) now re-validate
                Initialize();

                // Logger.Log($"LicenseSpot: Activation result: {State.Status} (Valid={State.Valid})");
            }
            catch (Exception ex)
            {
                State.Valid = false;
                State.Status = "Activation failed: " + ex.Message;
                // Logger.Log("LicenseSpot: exception in Activate(): " + ex);
            }
        }

    }
}
