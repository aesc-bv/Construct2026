using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using AESCConstruct25.FrameGenerator.Utilities;   // for Logger
using AESCConstruct25.FrameGenerator.UI;
using SpaceClaim.Api.V251.Extensibility;
using Panel = System.Windows.Forms.Panel;
using SpaceClaim.Api.V251;

namespace AESCConstruct25.UIMain
{
	public static class UIManager
	{
		public const string ProfileCommand = "AESCConstruct25.ProfileSidebar";
		public const string JointCommand = "AESCConstruct25.JointSidebar";
        public const string SettingsCommand = "AESCConstruct25.SettingsSidebar";

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
			//CloseJoint();

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
			//CloseProfile();

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
