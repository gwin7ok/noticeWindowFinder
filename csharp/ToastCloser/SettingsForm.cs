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
            // Larger layout so labels and inputs are not truncated
            this.txtDisplayLimit = new TextBox() { Left = 280, Top = 20, Width = 120 };
            this.txtPollInterval = new TextBox() { Left = 280, Top = 60, Width = 120 };
            this.chkDetectOnly = new CheckBox() { Left = 20, Top = 100, Text = "検出のみ (DetectOnly)", AutoSize = true };
            this.chkPreserveHistory = new CheckBox() { Left = 20, Top = 140, Text = "PreserveHistory", AutoSize = true };
            this.txtIdleMs = new TextBox() { Left = 280, Top = 180, Width = 120 };
            this.chkVerbose = new CheckBox() { Left = 20, Top = 220, Text = "VerboseLog", AutoSize = true };
            this.btnSave = new Button() { Text = "保存", Left = 140, Width = 100, Top = 260 };
            this.btnCancel = new Button() { Text = "キャンセル", Left = 260, Width = 100, Top = 260 };

            var lbl1 = new Label() { Left = 20, Top = 22, Width = 250, Text = "DisplayLimitSeconds:" , AutoSize = false};
            var lbl2 = new Label() { Left = 20, Top = 62, Width = 250, Text = "PollIntervalSeconds:", AutoSize = false };
            var lbl3 = new Label() { Left = 20, Top = 182, Width = 250, Text = "PreserveHistoryIdleMs:", AutoSize = false };

            this.ClientSize = new System.Drawing.Size(520, 340);
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
