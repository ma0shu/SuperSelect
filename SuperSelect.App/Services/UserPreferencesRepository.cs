using System.IO;
using System.Text.Json;

namespace SuperSelect.App.Services;

internal sealed class UserPreferencesRepository
{
    private sealed class SettingsPayload
    {
        public bool TypeFilterEnabled { get; set; }
    }

    private readonly object _syncRoot = new();
    private readonly string _settingsPath;
    private SettingsPayload _payload;

    public UserPreferencesRepository()
    {
        var storeDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SuperSelect");

        Directory.CreateDirectory(storeDirectory);
        _settingsPath = Path.Combine(storeDirectory, "settings.json");
        _payload = Load();
    }

    public bool TypeFilterEnabled
    {
        get
        {
            lock (_syncRoot)
            {
                return _payload.TypeFilterEnabled;
            }
        }
    }

    public void SetTypeFilterEnabled(bool enabled)
    {
        lock (_syncRoot)
        {
            if (_payload.TypeFilterEnabled == enabled)
            {
                return;
            }

            _payload.TypeFilterEnabled = enabled;
            SaveLocked();
        }
    }

    private SettingsPayload Load()
    {
        if (!File.Exists(_settingsPath))
        {
            return new SettingsPayload();
        }

        try
        {
            var text = File.ReadAllText(_settingsPath);
            return JsonSerializer.Deserialize<SettingsPayload>(text) ?? new SettingsPayload();
        }
        catch
        {
            return new SettingsPayload();
        }
    }

    private void SaveLocked()
    {
        var json = JsonSerializer.Serialize(_payload, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_settingsPath, json);
    }
}
