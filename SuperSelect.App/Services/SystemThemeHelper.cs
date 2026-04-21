using Microsoft.Win32;

namespace SuperSelect.App.Services;

internal static class SystemThemeHelper
{
    private const string PersonalizeRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string AppsUseLightThemeValueName = "AppsUseLightTheme";

    public static bool IsSystemAppDarkMode()
    {
        try
        {
            using var personalizeKey = Registry.CurrentUser.OpenSubKey(PersonalizeRegistryPath, writable: false);
            if (personalizeKey?.GetValue(AppsUseLightThemeValueName) is int value)
            {
                return value == 0;
            }
        }
        catch
        {
            // Ignore registry failures and fall back to light mode.
        }

        return false;
    }
}
