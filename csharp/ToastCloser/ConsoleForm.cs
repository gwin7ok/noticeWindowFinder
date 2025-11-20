using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ToastCloser
{
    public partial class ConsoleForm : Form
    {
        private RichTextBox _rtb = null!;
        private bool _autoScroll = true;

        public ConsoleForm()
        {
            InitializeComponent();
            try
            {
                Debug.WriteLine("ConsoleForm: constructor start");
            }
            catch { }

            // Load past logs first, then subscribe for live updates
            try
            {
                Debug.WriteLine("ConsoleForm: starting LoadPastLogs()");
                LoadPastLogs();
                Debug.WriteLine("ConsoleForm: LoadPastLogs() completed");
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ConsoleForm: LoadPastLogs() failed: " + ex.Message); } catch { }
            }

            try
            {
                Debug.WriteLine("ConsoleForm: subscribing to logger events");
                SubscribeLogger();
                Debug.WriteLine("ConsoleForm: subscribed to logger events");
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ConsoleForm: SubscribeLogger() failed: " + ex.Message); } catch { }
            }

            try
            {
                Debug.WriteLine("ConsoleForm: wiring scroll detection");
                WireScrollDetection();
                Debug.WriteLine("ConsoleForm: wired scroll detection");
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ConsoleForm: WireScrollDetection() failed: " + ex.Message); } catch { }
            }

            try
            {
                Debug.WriteLine("ConsoleForm: constructor done");
            }
            catch { }
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
                Debug.WriteLine("ConsoleForm.SubscribeLogger: attempting subscribe");
                if (Program.Logger.Instance != null)
                {
                    Program.Logger.Instance.OnLogLine += Instance_OnLogLine;
                    Debug.WriteLine("ConsoleForm.SubscribeLogger: subscribe OK");
                }
                else
                {
                    // logger not available
                    try { Debug.WriteLine("ConsoleForm.SubscribeLogger: Logger.Instance was null"); } catch { }
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
                Debug.WriteLine($"ConsoleForm.AppendLine: appending line length={line?.Length ?? 0}");
                _rtb.AppendText(line + Environment.NewLine);
                if (_rtb.Lines.Length > 5000)
                {
                    var lines = _rtb.Lines;
                    var keep = new string[4000];
                    Array.Copy(lines, lines.Length - 4000, keep, 0, 4000);
                    _rtb.Lines = keep;
                }
                if (_autoScroll)
                {
                    Debug.WriteLine("ConsoleForm.AppendLine: autoscroll true -> scrolling to caret");
                    _rtb.SelectionStart = _rtb.Text.Length;
                    _rtb.ScrollToCaret();
                }
                else
                {
                    Debug.WriteLine("ConsoleForm.AppendLine: autoscroll false -> not scrolling");
                }
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ConsoleForm.AppendLine exception: " + ex.Message); } catch { }
            }
        }

        private void WireScrollDetection()
        {
            try
            {
                Debug.WriteLine("ConsoleForm.WireScrollDetection: wiring events");
                _rtb.MouseWheel += (s, e) => CheckAutoScrollState();
                _rtb.MouseDown += (s, e) => CheckAutoScrollState();
                _rtb.KeyDown += (s, e) => CheckAutoScrollState();
                Debug.WriteLine("ConsoleForm.WireScrollDetection: events wired");
            }
            catch { }
        }

        private void CheckAutoScrollState()
        {
            try
            {
                Debug.WriteLine("ConsoleForm.CheckAutoScrollState: checking scroll state");
                bool atBottom = IsRichTextBoxAtBottom(_rtb);
                Debug.WriteLine($"ConsoleForm.CheckAutoScrollState: atBottom={atBottom}");
                // If user scrolled away from bottom, stop autoscroll. If user scrolled back to bottom, resume.
                _autoScroll = atBottom;
                Debug.WriteLine($"ConsoleForm.CheckAutoScrollState: _autoScroll set to {_autoScroll}");
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ConsoleForm.CheckAutoScrollState exception: " + ex.Message); } catch { }
            }
        }

        private static bool IsRichTextBoxAtBottom(RichTextBox rtb)
        {
            try
            {
                var si = new SCROLLINFO();
                si.cbSize = (uint)Marshal.SizeOf(si);
                si.fMask = (uint)(SIF_PAGE | SIF_POS | SIF_RANGE);
                if (GetScrollInfo(rtb.Handle, SB_VERT, ref si))
                {
                    // nPos + nPage >= nMax indicates bottom
                    bool atBottom = si.nPos + (int)si.nPage >= si.nMax;
                    try { Debug.WriteLine($"ConsoleForm.IsRichTextBoxAtBottom: nPos={si.nPos} nPage={si.nPage} nMax={si.nMax} atBottom={atBottom}"); } catch { }
                    return atBottom;
                }
            }
            catch { }
            return true;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SCROLLINFO
        {
            public uint cbSize;
            public uint fMask;
            public int nMin;
            public int nMax;
            public uint nPage;
            public int nPos;
            public int nTrackPos;
        }

        private const int SB_VERT = 1;
        private const int SIF_RANGE = 0x1;
        private const int SIF_PAGE = 0x2;
        private const int SIF_POS = 0x4;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetScrollInfo(IntPtr hWnd, int nBar, ref SCROLLINFO lpScrollInfo);

        private void LoadPastLogs()
        {
            // Load only the tail of log files to avoid huge memory/CPU usage
            const int maxTotalLines = 5000;
            const int perFileTail = 2000;
            try
            {
                var logsDir = Path.Combine(AppContext.BaseDirectory, "logs");
                if (!Directory.Exists(logsDir)) return;

                var files = Directory.GetFiles(logsDir, "auto_closer*").OrderBy(f => File.GetCreationTime(f)).ToArray();
                var combined = new System.Collections.Generic.List<string>();
                foreach (var f in files)
                {
                    try
                    {
                        var tail = ReadLastLines(f, perFileTail);
                        if (tail != null && tail.Length > 0) combined.AddRange(tail);
                        // keep combined bounded
                        if (combined.Count > maxTotalLines)
                        {
                            combined = combined.Skip(Math.Max(0, combined.Count - maxTotalLines)).ToList();
                        }
                    }
                    catch (Exception ex)
                    {
                        try { Debug.WriteLine($"ConsoleForm.LoadPastLogs: error reading {f}: {ex.Message}"); } catch { }
                    }
                }

                // Append lines to UI in a single BeginInvoke batch to avoid blocking
                if (combined.Count > 0)
                {
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            foreach (var line in combined)
                            {
                                try { AppendLine(line); } catch { }
                            }
                        }));
                    }
                    catch (Exception ex)
                    {
                        try { Debug.WriteLine("ConsoleForm.LoadPastLogs: BeginInvoke failed: " + ex.Message); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                try { AppendLine($"(ログ読み取りエラー: {ex.Message})"); } catch { }
            }
            finally { _autoScroll = true; }
        }

        private static string[] ReadLastLines(string path, int maxLines)
        {
            try
            {
                // Efficient streaming to keep only last N lines
                var q = new System.Collections.Generic.Queue<string>();
                foreach (var line in File.ReadLines(path))
                {
                    q.Enqueue(line);
                    if (q.Count > maxLines) q.Dequeue();
                }
                return q.ToArray();
            }
            catch (Exception ex)
            {
                try { Debug.WriteLine("ReadLastLines error: " + ex.Message); } catch { }
                return Array.Empty<string>();
            }
        }

        private DateTime? ParseLogLineTimestamp(string line)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(line) || line.Length < 19) return null;
                var prefix = line.Substring(0, 19);
                if (DateTime.TryParseExact(prefix, "yyyy/MM/dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out var dt)) return dt;
            }
            catch { }
            return null;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try { if (Program.Logger.Instance != null) Program.Logger.Instance.OnLogLine -= Instance_OnLogLine; } catch { }
            base.OnFormClosed(e);
        }
    }
}
