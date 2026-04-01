using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string logPath = Path.Combine(baseDir, "data", "startup-error.log");

        try
        {
            string scriptPath = Path.Combine(baseDir, "ScreenTimeTracker.ps1");
            if (!File.Exists(scriptPath))
            {
                throw new FileNotFoundException("ScreenTimeTracker.ps1 was not found next to the launcher.", scriptPath);
            }

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = ResolvePowerShellPath(),
                Arguments = BuildArguments(scriptPath, args),
                WorkingDirectory = baseDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Process.Start(startInfo);
            return 0;
        }
        catch (Exception ex)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath) ?? baseDir);
                File.WriteAllText(logPath, ex.ToString());
            }
            catch
            {
            }

            try
            {
                MessageBox.Show(
                    "The tracker could not start." + Environment.NewLine + Environment.NewLine +
                    ex.Message + Environment.NewLine + Environment.NewLine +
                    "Details were saved to:" + Environment.NewLine + logPath,
                    "Screen Time Tracker",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch
            {
            }

            return 1;
        }
    }

    private static string ResolvePowerShellPath()
    {
        string windowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        string candidate = Path.Combine(windowsDir, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
        return File.Exists(candidate) ? candidate : "powershell.exe";
    }

    private static string BuildArguments(string scriptPath, IEnumerable<string> args)
    {
        List<string> parts = new List<string>
        {
            "-NoLogo",
            "-NoProfile",
            "-WindowStyle",
            "Hidden",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            Quote(scriptPath)
        };

        if (args != null)
        {
            foreach (string arg in args)
            {
                if (string.IsNullOrWhiteSpace(arg))
                {
                    continue;
                }

                parts.Add(Quote(arg));
            }
        }

        return string.Join(" ", parts);
    }

    private static string Quote(string value)
    {
        if (value == null)
        {
            return "\"\"";
        }

        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }
}
