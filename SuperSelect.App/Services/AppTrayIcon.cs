using System.Drawing;
using System.Windows.Forms;

namespace SuperSelect.App.Services;

internal sealed class AppTrayIcon : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly Icon _icon;
    private readonly ContextMenuStrip _menu;

    public AppTrayIcon(Icon icon)
    {
        _icon = (Icon)icon.Clone();
        _menu = new ContextMenuStrip();

        var openMenuItem = new ToolStripMenuItem("打开");
        openMenuItem.Click += (_, _) => OpenRequested?.Invoke();
        _menu.Items.Add(openMenuItem);

        var exitMenuItem = new ToolStripMenuItem("退出");
        exitMenuItem.Click += (_, _) => ExitRequested?.Invoke();
        _menu.Items.Add(exitMenuItem);

        _notifyIcon = new NotifyIcon
        {
            Text = "SuperSelect",
            Icon = _icon,
            Visible = true,
            ContextMenuStrip = _menu,
        };

        _notifyIcon.DoubleClick += (_, _) => OpenRequested?.Invoke();
    }

    public event Action? OpenRequested;
    public event Action? ExitRequested;

    public void Dispose()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _menu.Dispose();
        _icon.Dispose();
    }
}
