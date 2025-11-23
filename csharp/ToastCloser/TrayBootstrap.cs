using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToastCloser
{
    static class TrayBootstrap
    {
        // Hold the process-wide mutex so the OS-level named mutex isn't released.
        private static System.Threading.Mutex? _processMutex = null;
        [STAThread]
        static void Main()
        {
            // Ensure single-instance at process start (applies to the Tray bootstrap)
            try
            {
                bool createdNew = false;
                var mutexName = "Global\\ToastCloser_mutex";
                try { _processMutex = new System.Threading.Mutex(true, mutexName, out createdNew); } catch { }
                if (!createdNew)
                {
                    try { MessageBox.Show("ToastCloser is already running.", "ToastCloser", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                    return;
                }
            }
            catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var cfg = Config.Load();

            // Ensure logger exists early so Tray UI and Console can use it immediately.
            try
            {
                string exeFolder = string.Empty;
                try { exeFolder = System.IO.Path.GetDirectoryName(System.Environment.GetCommandLineArgs()?.FirstOrDefault() ?? string.Empty) ?? string.Empty; } catch { }
                try { if (string.IsNullOrEmpty(exeFolder)) exeFolder = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? string.Empty) ?? string.Empty; } catch { }
                try { if (string.IsNullOrEmpty(exeFolder)) exeFolder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly()?.Location ?? string.Empty) ?? string.Empty; } catch { }
                if (string.IsNullOrEmpty(exeFolder)) exeFolder = AppContext.BaseDirectory ?? System.IO.Directory.GetCurrentDirectory();

                var logsDir = System.IO.Path.Combine(exeFolder, "logs");
                try { System.IO.Directory.CreateDirectory(logsDir); } catch { }
                try
                {
                    var logPath = System.IO.Path.Combine(logsDir, "auto_closer.log");
                    // Write a TEMP marker to indicate we attempted logger init
                    try { System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"toastcloser_logger_init_attempt_{System.DateTime.UtcNow:yyyyMMddHHmmss}.txt"), $"logPath={logPath}"); } catch { }
                    // Set debug flag from config
                    Program.Logger.IsDebugEnabled = cfg?.VerboseLog ?? false;
                    Program.Logger.Instance = new Program.Logger(logPath);
                }
                catch { }
            }
            catch { }

            var ctx = new TrayApplicationContext(cfg);

            // Start background scanner. For debugging, support an env var
            // `TOASTCLOSER_INLINE` which causes `Program.Main` to be invoked
            // synchronously (inline) so any exceptions or early returns are
            // visible immediately instead of running in a background Task.
            try
            {
                var inline = false;
                try { inline = !string.IsNullOrEmpty(System.Environment.GetEnvironmentVariable("TOASTCLOSER_INLINE")); } catch { inline = false; }
                if (inline)
                {
                    try
                    {
                        Program.Main(new string[] { "--background-service" });
                    }
                    catch (Exception ex)
                    {
                        try { System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"toastcloser_inline_exception_{System.DateTime.UtcNow:yyyyMMddHHmmss}.txt"), ex.ToString()); } catch { }
                    }
                }
                else
                {
                    Task.Run(() =>
                    {
                        try
                        {
                            Program.Main(new string[] { "--background-service" });
                        }
                        catch (Exception ex)
                        {
                            try { System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"toastcloser_background_exception_{System.DateTime.UtcNow:yyyyMMddHHmmss}.txt"), ex.ToString()); } catch { }
                        }
                    });
                }
            }
            catch { }

            Application.Run(ctx);
        }
    }
}
