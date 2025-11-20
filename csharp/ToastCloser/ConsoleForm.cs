using System;
using System.Windows.Forms;

namespace ToastCloser
{
    public partial class ConsoleForm : Form
    {
        private RichTextBox _rtb = null!;

        public ConsoleForm()
        {
            InitializeComponent();
            SubscribeLogger();
        }

        private void InitializeComponent()
        {
            this._rtb = new RichTextBox() { Dock = DockStyle.Fill, ReadOnly = true }; 
            this.ClientSize = new System.Drawing.Size(800, 400);
            this.Controls.Add(_rtb);
            this.Text = "ToastCloser Console";
        }

        private void SubscribeLogger()
        {
            try
            {
                if (Program.Logger.Instance != null)
                {
                    Program.Logger.Instance.OnLogLine += Instance_OnLogLine;
                }
            }
            catch { }
        }

        private void Instance_OnLogLine(string line)
        {
            if (this.IsDisposed) return;
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action(() => AppendLine(line))); } catch { }
            }
            else AppendLine(line);
        }

        private void AppendLine(string line)
        {
            try
            {
                _rtb.AppendText(line + Environment.NewLine);
                if (_rtb.Lines.Length > 5000)
                {
                    var lines = _rtb.Lines;
                    var keep = new string[4000];
                    Array.Copy(lines, lines.Length - 4000, keep, 0, 4000);
                    _rtb.Lines = keep;
                }
                _rtb.SelectionStart = _rtb.Text.Length;
                _rtb.ScrollToCaret();
            }
            catch { }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try { if (Program.Logger.Instance != null) Program.Logger.Instance.OnLogLine -= Instance_OnLogLine; } catch { }
            base.OnFormClosed(e);
        }
    }
}
