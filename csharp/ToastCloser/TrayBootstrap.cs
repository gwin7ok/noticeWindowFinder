using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToastCloser
{
    static class TrayBootstrap
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var cfg = Config.Load();

            var ctx = new TrayApplicationContext(cfg);

            // Start background scanner in a thread so UI remains responsive.
            Task.Run(() =>
            {
                try
                {
                    // Call existing Program.Main in background mode by passing marker arg
                    Program.Main(new string[] { "--background-service" });
                }
                catch { }
            });

            Application.Run(ctx);
        }
    }
}
