using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Text;

namespace SuperSelect.App.Services;

internal static class AdminPrivilegeHelper
{
    private const string TaskName = "SuperSelectLauncher";

    public static bool IsRunAsAdministrator()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    public static void EnableRunOnStartupAsAdmin()
    {
        var exePath = GetCurrentExecutablePathOrThrow();

        var xml = $@"<?xml version=""1.0"" encoding=""UTF-16""?>
<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id=""Author"">
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context=""Author"">
    <Exec>
      <Command>""{exePath}""</Command>
    </Exec>
  </Actions>
</Task>";

        var tempXmlPath = Path.Combine(
            Path.GetTempPath(),
            $"SuperSelectTask-{Environment.ProcessId}-{Guid.NewGuid():N}.xml");
        File.WriteAllText(tempXmlPath, xml, Encoding.Unicode);

        try
        {
            RunSchtasksElevated($"/create /tn \"{TaskName}\" /xml \"{tempXmlPath}\" /f");
        }
        finally
        {
            if (File.Exists(tempXmlPath))
            {
                try
                {
                    File.Delete(tempXmlPath);
                }
                catch
                {
                    // Ignore temp cleanup failure.
                }
            }
        }
    }

    public static void DisableRunOnStartupAsAdmin()
    {
        if (!CheckRunOnStartupConfigured())
        {
            return;
        }

        RunSchtasksElevated($"/delete /tn \"{TaskName}\" /f");
    }

    public static bool CheckRunOnStartupConfigured()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "schtasks",
                Arguments = $"/query /tn \"{TaskName}\"",
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };
            using var proc = Process.Start(psi);
            if (proc is null)
            {
                return false;
            }

            if (!proc.WaitForExit(2500))
            {
                try
                {
                    proc.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Ignore timeout cleanup failures.
                }

                return false;
            }

            return proc.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    public static void RestartCurrentProcessAsAdministrator()
    {
        var exePath = GetCurrentExecutablePathOrThrow();
        using var process = Process.Start(
            new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true,
                Verb = "runas",
            });

        if (process is null)
        {
            throw new InvalidOperationException("无法启动提权后的新进程。");
        }
    }

    private static string GetCurrentExecutablePathOrThrow()
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (!string.IsNullOrWhiteSpace(exePath))
        {
            return exePath;
        }

        throw new InvalidOperationException("无法获取当前应用路径。");
    }

    private static void RunSchtasksElevated(string arguments)
    {
        using var process = Process.Start(
            new ProcessStartInfo
            {
                FileName = "schtasks",
                Arguments = arguments,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = true,
                Verb = "runas",
            });

        if (process is null)
        {
            throw new InvalidOperationException("无法启动 schtasks。");
        }

        process.WaitForExit();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"schtasks 执行失败，退出码：{process.ExitCode}");
        }
    }
}
