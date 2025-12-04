using System;
using System.Diagnostics;
using System.Windows;

namespace ZibllWindows
{
    public class Util
    {
        public static bool YesNo(string msg, string title)
        {
            var result = MessageBox.Show(msg, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            return result == MessageBoxResult.Yes;
        }

        public static string RunProcess(string exe, string args, string workDir, bool wait)
        {
            string output = "";
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = args,
                    WorkingDirectory = workDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                Process p = Process.Start(psi);
                if (wait)
                {
                    output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return output;
        }
    }
}
