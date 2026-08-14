using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        try
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string scriptPath = Path.Combine(appDir, "app.ps1");

            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    "未找到主程序：app.ps1\n\n请确认草种测定管理.exe 与 app.ps1 位于同一文件夹。",
                    "草种测定管理",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                return;
            }

            ProcessStartInfo psi = new ProcessStartInfo();

            psi.FileName = "powershell.exe";

            psi.Arguments =
                "-NoProfile " +
                "-ExecutionPolicy Bypass " +
                "-WindowStyle Hidden " +
                "-File \"" + scriptPath + "\"";

            psi.WorkingDirectory = appDir;

            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;

            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "程序启动失败：\n\n" + ex.Message,
                "草种测定管理",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }
}