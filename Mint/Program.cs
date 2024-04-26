using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace Mint
{
    static class Program
    {
        const string _jsonAssembly = @"Mint.Newtonsoft.Json.dll";
        const string _mutexGuid = "{SPAMISH-1A670425-A3BD-4C25-9ED5-29CA1255599A-MINT}";
        internal static Mutex MUTEX;
        static bool _notRunning;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main()
        {
            if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();

            string AppDataExe = Path.Combine(Application.UserAppDataPath, "Mint.exe");

            if (Assembly.GetExecutingAssembly().Location != AppDataExe)
            {
                File.Copy(Assembly.GetExecutingAssembly().Location, AppDataExe, true);
                Process p = new Process();
                p.StartInfo.FileName = AppDataExe;
                p.Start();
                Environment.Exit(0);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (MUTEX = new Mutex(true, _mutexGuid, out _notRunning))
            {
                if (_notRunning)
                {
                    EmbeddedAssembly.Load(_jsonAssembly, _jsonAssembly.Replace("Mint.", string.Empty));
                    AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
                    
                    Options.LoadSettings();
                    Form mintApp = new MainForm();

                    // If AutoStart enabled then start in system tray
                    if (Options.CurrentOptions.AutoStart)
                    {
                        Application.Run();
                    }
                    // Else open form
                    else
                    {
                        Application.Run(mintApp);
                    }
                }
                else
                {
                    MessageBox.Show("Mint is already running in the background!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Environment.Exit(0);
                }
            }
        }

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }
    }
}
