using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

namespace ToastCloser
{
    public class Config
    {
        public double DisplayLimitSeconds { get; set; } = 10.0;
        public double PollIntervalSeconds { get; set; } = 1.0;
        public bool DetectOnly { get; set; } = false;
        public bool PreserveHistory { get; set; } = true;
        public string PreserveHistoryMode { get; set; } = "noticecenter";
        public int PreserveHistoryIdleMs { get; set; } = 2000;
        public int PreserveHistoryMaxMonitorSeconds { get; set; } = 15;
        public int DetectionTimeoutMs { get; set; } = 2000;
        public int WinShortcutKeyDelayMs { get; set; } = 300;
        public bool VerboseLog { get; set; } = false;

        public static string ConfigFileName => Path.Combine(AppContext.BaseDirectory, "ToastCloser.ini");

        public void Save()
        {
            try
            {
                var lines = new List<string>();
                lines.Add("[General]");
                lines.Add($"DisplayLimitSeconds={DisplayLimitSeconds.ToString(CultureInfo.InvariantCulture)}");
                lines.Add($"PollIntervalSeconds={PollIntervalSeconds.ToString(CultureInfo.InvariantCulture)}");
                lines.Add($"DetectOnly={DetectOnly}");
                lines.Add($"PreserveHistory={PreserveHistory}");
                lines.Add($"PreserveHistoryMode={PreserveHistoryMode}");
                lines.Add($"PreserveHistoryIdleMs={PreserveHistoryIdleMs}");
                lines.Add($"PreserveHistoryMaxMonitorSeconds={PreserveHistoryMaxMonitorSeconds}");
                lines.Add($"DetectionTimeoutMs={DetectionTimeoutMs}");
                lines.Add($"WinShortcutKeyDelayMs={WinShortcutKeyDelayMs}");
                lines.Add($"VerboseLog={VerboseLog}");
                File.WriteAllLines(ConfigFileName, lines);
            }
            catch { }
        }

        public static Config Load()
        {
            var cfg = new Config();
            try
            {
                if (!File.Exists(ConfigFileName)) return cfg;
                var lines = File.ReadAllLines(ConfigFileName);
                foreach (var raw in lines)
                {
                    var line = raw.Trim();
                    if (string.IsNullOrEmpty(line) || line.StartsWith("#") || line.StartsWith(";") || line.StartsWith("[")) continue;
                    var idx = line.IndexOf('=');
                    if (idx <= 0) continue;
                    var key = line.Substring(0, idx).Trim();
                    var val = line.Substring(idx + 1).Trim();
                    switch (key)
                    {
                        case "DisplayLimitSeconds": double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var v1); cfg.DisplayLimitSeconds = v1; break;
                        case "PollIntervalSeconds": double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var v2); cfg.PollIntervalSeconds = v2; break;
                        case "DetectOnly": bool.TryParse(val, out var b1); cfg.DetectOnly = b1; break;
                        case "PreserveHistory": bool.TryParse(val, out var b2); cfg.PreserveHistory = b2; break;
                        case "PreserveHistoryMode": cfg.PreserveHistoryMode = val; break;
                        case "PreserveHistoryIdleMs": int.TryParse(val, out var i1); cfg.PreserveHistoryIdleMs = i1; break;
                        case "PreserveHistoryMaxMonitorSeconds": int.TryParse(val, out var i2); cfg.PreserveHistoryMaxMonitorSeconds = i2; break;
                        case "DetectionTimeoutMs": int.TryParse(val, out var i3); cfg.DetectionTimeoutMs = i3; break;
                        case "WinShortcutKeyDelayMs": int.TryParse(val, out var i4); cfg.WinShortcutKeyDelayMs = i4; break;
                        case "VerboseLog": bool.TryParse(val, out var b3); cfg.VerboseLog = b3; break;
                    }
                }
            }
            catch { }
            return cfg;
        }
    }
}
