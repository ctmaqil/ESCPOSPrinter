using System;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;

namespace ESCPrintApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (sender, e) =>
            {
                MessageBox.Show($"Application error:\n{e.Exception.Message}",
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                MessageBox.Show($"Critical error:\n{e.ExceptionObject}",
                               "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            };

            SetIE11Mode();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void SetIE11Mode()
        {
            try
            {
                string executableName = Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"))
                {
                    key.SetValue(executableName, 11001, RegistryValueKind.DWord);
                }
            }
            catch
            {
                // Handle registry access errors gracefully
            }
        }
    }
}
