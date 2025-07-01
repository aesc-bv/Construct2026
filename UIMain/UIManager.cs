using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using AESCConstruct25.FrameGenerator.Utilities;   // for Logger
using AESCConstruct25.FrameGenerator.UI;
using SpaceClaim.Api.V242.Extensibility;
using Panel = System.Windows.Forms.Panel;
using SpaceClaim.Api.V242;
using AESCConstruct25.UI;

namespace AESCConstruct25.UIMain
{
	public static class UIManager
	{
		public const string ProfileCommand = "AESCConstruct25.ProfileSidebar";
		public const string JointCommand = "AESCConstruct25.JointSidebar";
        public const string SettingsCommand = "AESCConstruct25.SettingsSidebar";
        public const string PlateCommand = "AESCConstruct25.Plate";
        public const string FastenerCommand = "AESCConstruct25.Fastener";

        // PROFILE
        static PanelTab _profileTab;
		static Panel _profilePanel;
		static ElementHost _profileHost;
		static ProfileSelectionControl _profileControl;

        // JOINT
        static PanelTab _jointTab;
        static Panel _jointPanel;
        static ElementHost _jointHost;
        static JointSelectionControl _jointControl;

        // SETTINGS
        static PanelTab _settingsTab;
        static Panel _settingsPanel;
        static ElementHost _settingsHost;
        static SettingsControl _settingsControl;

        static PanelTab _plateTab;
        static Panel _platePanel;
        static ElementHost _plateHost;
        static PlatesControl _plateControl;

        static PanelTab _fastenerTab;
        static Panel _fastenerPanel;
        static ElementHost _fastenerHost;
        static FastenersControl _fastenerControl;

        public static bool IncludeMaterialInExcel { get; set; }
        public static bool IncludeMaterialInBOM { get; set; }

        public static void RegisterAll()
		{
			var pCmd = Command.Create(ProfileCommand);
			pCmd.Text = "Profile Selector";
			pCmd.Hint = "Open the profile selection sidebar";
			pCmd.Executing += OnProfileToggle;

			var jCmd = Command.Create(JointCommand);
			jCmd.Text = "Joint Selector";
			jCmd.Hint = "Open the joint selection sidebar";
			jCmd.Executing += OnJointToggle;

            var sCmd = Command.Create(SettingsCommand);
            sCmd.Text = "Settings";
            sCmd.Hint = "Open the settings window";
            sCmd.Executing += OnSettingsToggle;

            var plCmd = Command.Create(PlateCommand);
            plCmd.Text = "Plate";
            plCmd.Hint = "Open the plate‐creation pane";
            plCmd.Executing += OnPlateToggle;

            var fCmd = Command.Create(FastenerCommand);
            fCmd.Text = "Fastener";
            fCmd.Hint = "Open the fastener insertion pane";
            fCmd.Executing += OnFastenerToggle;

            //AddMaterialCheckboxExcel.Checked += (s, e) => UIManager.IncludeMaterialInExcel = true;
            //AddMaterialCheckboxExcel.Unchecked += (s, e) => UIManager.IncludeMaterialInExcel = false;

            //AddMaterialCheckboxBOM.Checked += (s, e) => UIManager.IncludeMaterialInBOM = true;
            //AddMaterialCheckboxBOM.Unchecked += (s, e) => UIManager.IncludeMaterialInBOM = false;
        }

        static void OnSettingsToggle(object sender, EventArgs e)
        {
            var cmd = (Command)sender;

            // close any other pane
            CloseProfile();
            CloseJoint();

            // if already open, just close it
            if (_settingsTab != null)
            {
                CloseSettings();
                return;
            }

            // build/rebuild if needed
            if (_settingsPanel == null || _settingsPanel.IsDisposed)
            {
                _settingsControl = new SettingsControl();
                _settingsHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _settingsControl
                };
                _settingsPanel = new Panel { Dock = DockStyle.Fill };
                _settingsPanel.Controls.Add(_settingsHost);
            }

            // show it
            _settingsTab = PanelTab.Create(cmd, _settingsPanel, DockLocation.Right, 500, false);
            _settingsTab?.Activate();
        }

        public static void CloseSettings()
        {
            if (_settingsTab != null)
            {
                _settingsTab.Close();
                _settingsTab = null;
            }
            if (_settingsPanel != null && !_settingsPanel.IsDisposed)
                _settingsPanel.Dispose();

            _settingsPanel = null;
            _settingsHost = null;
            _settingsControl = null;
        }

        static void OnProfileToggle(object sender, EventArgs e)
		{
			var cmd = (Command)sender;
			// always close the other one first
			CloseJoint();

			// if it’s already open, just close it
			if (_profileTab != null)
			{
				CloseProfile();
				return;
			}

			// otherwise build/rebuild
			if (_profilePanel == null || _profilePanel.IsDisposed)
			{
				_profileHost = null;
				_profileControl = null;
				_profilePanel = null;

				_profileControl = new ProfileSelectionControl();
				_profileHost = new ElementHost
				{
					Dock = DockStyle.Fill,
					Child = _profileControl
				};
				_profilePanel = new Panel
				{
					Dock = DockStyle.Fill
				};
				_profilePanel.Controls.Add(_profileHost);
			}

			// show it
			_profileTab = PanelTab.Create(cmd, _profilePanel, DockLocation.Right, 500, false);
			_profileTab?.Activate();
		}

		static void OnJointToggle(object sender, EventArgs e)
		{
			var cmd = (Command)sender;
			// always close the other one first
			CloseProfile();

			// if it’s already open, just close it
			if (_jointTab != null)
			{
				CloseJoint();
				return;
			}

			// otherwise build/rebuild
			if (_jointPanel == null || _jointPanel.IsDisposed)
			{
				_jointHost = null;
				_jointControl = null;
				_jointPanel = null;

				_jointControl = new JointSelectionControl();
				_jointHost = new ElementHost
				{
					Dock = DockStyle.Fill,
					Child = _jointControl
				};
				_jointPanel = new Panel
				{
					Dock = DockStyle.Fill
				};
				_jointPanel.Controls.Add(_jointHost);
			}

			// show it
			_jointTab = PanelTab.Create(cmd, _jointPanel, DockLocation.Right, 500, false);
			_jointTab?.Activate();
		}

        static void OnPlateToggle(object sender, EventArgs e)
        {
            var cmd = (Command)sender;

            // 1) close anything else
            CloseProfile();
            CloseJoint();
            CloseSettings();

            // 2) if already open, just close it
            if (_plateTab != null)
            {
                ClosePlate();
                return;
            }

            // 3) otherwise (re)build the panel
            if (_platePanel == null || _platePanel.IsDisposed)
            {
                _plateControl = new PlatesControl();
                _plateHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _plateControl
                };
                _platePanel = new Panel
                {
                    Dock = DockStyle.Fill
                };
                _platePanel.Controls.Add(_plateHost);
            }

            // 4) show it on the right
            _plateTab = PanelTab.Create(cmd, _platePanel, DockLocation.Right, 500, false);
            _plateTab?.Activate();
        }

        static void OnFastenerToggle(object sender, EventArgs e)
        {
            var cmd = (Command)sender;

            // close everything else
            CloseProfile();
            CloseJoint();
            CloseSettings();
            ClosePlate();

            // if already open, close it
            if (_fastenerTab != null)
            {
                CloseFastener();
                return;
            }

            // build/rebuild
            if (_fastenerPanel == null || _fastenerPanel.IsDisposed)
            {
                _fastenerControl = new FastenersControl();
                _fastenerHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _fastenerControl
                };
                _fastenerPanel = new Panel { Dock = DockStyle.Fill };
                _fastenerPanel.Controls.Add(_fastenerHost);
            }

            // show on the right
            _fastenerTab = PanelTab.Create(cmd, _fastenerPanel, DockLocation.Right, 500, false);
            _fastenerTab?.Activate();
        }

        public static void CloseFastener()
        {
            if (_fastenerTab != null)
            {
                _fastenerTab.Close();
                _fastenerTab = null;
            }
            if (_fastenerPanel != null && !_fastenerPanel.IsDisposed)
                _fastenerPanel.Dispose();

            _fastenerPanel = null;
            _fastenerHost = null;
            _fastenerControl = null;
        }


        public static void ClosePlate()
        {
            if (_plateTab != null)
            {
                _plateTab.Close();
                _plateTab = null;
            }

            if (_platePanel != null && !_platePanel.IsDisposed)
                _platePanel.Dispose();

            _platePanel = null;
            _plateHost = null;
            _plateControl = null;
        }

        /// <summary>
        /// Close the Profile panel if it’s open.
        /// </summary>
        public static void CloseProfile()
		{
			if (_profileTab != null)
			{
				_profileTab.Close();
				_profileTab = null;
			}
			if (_profilePanel != null && !_profilePanel.IsDisposed)
				_profilePanel.Dispose();
			_profilePanel = null;
			_profileHost = null;
			_profileControl = null;
		}

		/// <summary>
		/// Close the Joint panel if it’s open.
		/// </summary>
		public static void CloseJoint()
		{
			if (_jointTab != null)
			{
				_jointTab.Close();
				_jointTab = null;
			}
			if (_jointPanel != null && !_jointPanel.IsDisposed)
				_jointPanel.Dispose();
			_jointPanel = null;
			_jointHost = null;
			_jointControl = null;
		}
	}
}
