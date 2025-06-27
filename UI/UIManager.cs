using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using AESCConstruct25.Utilities;   // for Logger
using AESCConstruct25.UI;
using SpaceClaim.Api.V251.Extensibility;
using Panel = System.Windows.Forms.Panel;
using SpaceClaim.Api.V251;

namespace AESCConstruct25.Commands
{
	public static class UIManager
	{
		public const string ProfileCommand = "AESCConstruct25.ProfileSidebar";
		public const string JointCommand = "AESCConstruct25.JointSidebar";

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
