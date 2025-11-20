using System;
using System.Windows.Forms;

namespace ToastCloser
{
    public partial class SettingsForm : Form
    {
        private Config _config;
        public event EventHandler<Config>? ConfigSaved;

        public SettingsForm(Config cfg)
        {
            _config = cfg;
            InitializeComponent();
            LoadValues();
        }

        private void LoadValues()
        {
            txtDisplayLimit.Text = _config.DisplayLimitSeconds.ToString();
            txtPollInterval.Text = _config.PollIntervalSeconds.ToString();
            chkDetectOnly.Checked = _config.DetectOnly;
            chkPreserveHistory.Checked = _config.PreserveHistory;
            txtIdleMs.Text = _config.PreserveHistoryIdleMs.ToString();
            chkVerbose.Checked = _config.VerboseLog;
        }

        private void SaveValues()
        {
            double.TryParse(txtDisplayLimit.Text, out var d); _config.DisplayLimitSeconds = d;
            double.TryParse(txtPollInterval.Text, out var p); _config.PollIntervalSeconds = p;
            _config.DetectOnly = chkDetectOnly.Checked;
            _config.PreserveHistory = chkPreserveHistory.Checked;
            int.TryParse(txtIdleMs.Text, out var im); _config.PreserveHistoryIdleMs = im;
            _config.VerboseLog = chkVerbose.Checked;
        }

        private void btnSave_Click(object? sender, EventArgs e)
        {
            SaveValues();
            try { _config.Save(); } catch { }
            ConfigSaved?.Invoke(this, _config);
            this.Close();
        }

        private void btnCancel_Click(object? sender, EventArgs e)
        {
            this.Close();
        }

        #region Designer
        private TextBox txtDisplayLimit = null!;
        private TextBox txtPollInterval = null!;
        private TextBox txtIdleMs = null!;
        private CheckBox chkDetectOnly = null!;
        private CheckBox chkPreserveHistory = null!;
        private CheckBox chkVerbose = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;

        private void InitializeComponent()
        {
            this.txtDisplayLimit = new TextBox() { Left = 140, Top = 12, Width = 80 };
            this.txtPollInterval = new TextBox() { Left = 140, Top = 42, Width = 80 };
            this.chkDetectOnly = new CheckBox() { Left = 12, Top = 72, Text = "検出のみ (DetectOnly)" };
            this.chkPreserveHistory = new CheckBox() { Left = 12, Top = 102, Text = "PreserveHistory" };
            this.txtIdleMs = new TextBox() { Left = 140, Top = 132, Width = 80 };
            this.chkVerbose = new CheckBox() { Left = 12, Top = 162, Text = "VerboseLog" };
            this.btnSave = new Button() { Text = "保存", Left = 50, Width = 80, Top = 200 };
            this.btnCancel = new Button() { Text = "キャンセル", Left = 150, Width = 80, Top = 200 };

            var lbl1 = new Label() { Left = 12, Top = 12, Text = "DisplayLimitSeconds:" };
            var lbl2 = new Label() { Left = 12, Top = 42, Text = "PollIntervalSeconds:" };
            var lbl3 = new Label() { Left = 12, Top = 132, Text = "PreserveHistoryIdleMs:" };

            this.ClientSize = new System.Drawing.Size(260, 240);
            this.Controls.AddRange(new Control[] { lbl1, lbl2, lbl3, txtDisplayLimit, txtPollInterval, txtIdleMs, chkDetectOnly, chkPreserveHistory, chkVerbose, btnSave, btnCancel });
            this.Text = "ToastCloser 設定";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            btnSave.Click += btnSave_Click;
            btnCancel.Click += btnCancel_Click;
        }
        #endregion
    }
}
