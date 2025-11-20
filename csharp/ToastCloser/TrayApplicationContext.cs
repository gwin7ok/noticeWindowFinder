using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ToastCloser
{
    public class TrayApplicationContext : ApplicationContext
    {
        private NotifyIcon _trayIcon;
        private Config _config;
        private SettingsForm? _settingsForm;
        private ConsoleForm? _consoleForm;

        public TrayApplicationContext(Config cfg)
        {
            _config = cfg;

            var iconPath = Path.Combine(AppContext.BaseDirectory, "ToastCloser.ico");
            Icon icon = SystemIcons.Application;
            try { if (File.Exists(iconPath)) icon = new Icon(iconPath); } catch { }

            _trayIcon = new NotifyIcon()
            {
                Icon = icon,
                Text = "ToastCloser",
                Visible = true
            };

            var menu = new ContextMenuStrip();
            var settingsItem = new ToolStripMenuItem("設定...");
            settingsItem.Click += (s, e) => ShowSettings();
            var consoleItem = new ToolStripMenuItem("コンソールを表示");
            consoleItem.Click += (s, e) => ToggleConsole();
            var reloadItem = new ToolStripMenuItem("設定を再読み込み");
            reloadItem.Click += (s, e) => ReloadConfig();
            var exitItem = new ToolStripMenuItem("終了");
            exitItem.Click += (s, e) => ExitApplication();

            menu.Items.Add(settingsItem);
            menu.Items.Add(consoleItem);
            menu.Items.Add(reloadItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            _trayIcon.ContextMenuStrip = menu;
            _trayIcon.DoubleClick += (s, e) => ToggleConsole();

            // Show a balloon tip on first run
            try { _trayIcon.ShowBalloonTip(2000, "ToastCloser", "Tray mode: 右クリックで設定・コンソールを開けます", ToolTipIcon.Info); } catch { }
        }

        private void ShowSettings()
        {
            if (_settingsForm == null || _settingsForm.IsDisposed)
            {
                _settingsForm = new SettingsForm(_config);
                _settingsForm.ConfigSaved += (s, cfg) =>
                {
                    _config = cfg;
                    _config.Save();
                    // Optionally notify running scanner to apply new config (future)
                };
                _settingsForm.Show();
            }
            else
            {
                _settingsForm.BringToFront();
            }
        }

        private void ToggleConsole()
        {
            if (_consoleForm == null || _consoleForm.IsDisposed)
            {
                _consoleForm = new ConsoleForm();
                _consoleForm.Show();
            }
            else
            {
                if (_consoleForm.Visible) _consoleForm.Hide(); else _consoleForm.Show();
            }
        }

        private void ReloadConfig()
        {
            try
            {
                _config = Config.Load();
                MessageBox.Show("設定を再読み込みしました。", "ToastCloser", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("設定の再読み込みに失敗しました: " + ex.Message);
            }
        }

        private void ExitApplication()
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            Application.Exit();
        }
    }
}
