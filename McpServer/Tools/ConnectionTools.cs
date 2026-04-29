using System.ComponentModel;
using ModelContextProtocol.Server;

namespace McpKompas.Tools;

[McpServerToolType]
public sealed class ConnectionTools
{
    [McpServerTool(Name = "kompas_connect")]
    [Description("Connect to a running Kompas 3D instance, or start a new one if none is running. " +
                 "Must be called before any other Kompas tool.")]
    public static string Connect(
        [Description("Set to true to always start a new Kompas instance instead of attaching to an existing one.")]
        bool start_new = false,
        [Description("Set to true to show the Kompas window (visual mode), false to run in background (hidden). Default is true.")]
        bool visible = true)
    {
        var (success, message) = KompasConnector.Connect(start_new, visible);
        if (success && visible) KompasConnector.BringToFront();
        return success ? message : $"ERROR: {message}";
    }

    [McpServerTool(Name = "kompas_set_visible")]
    [Description("Show or hide the Kompas 3D window. Use visible=true for visual mode, visible=false for background mode.")]
    public static string SetVisible(
        [Description("True to show the Kompas window, false to hide it.")]
        bool visible)
    {
        return KompasConnector.SetVisible(visible);
    }

    [McpServerTool(Name = "kompas_maximize")]
    [Description("Maximize the Kompas 3D window to fill the entire screen.")]
    public static string Maximize()
    {
        return KompasConnector.Maximize();
    }

    [McpServerTool(Name = "kompas_disconnect")]
    [Description("Release the COM connection to Kompas 3D.")]
    public static string Disconnect()
    {
        // History refs become meaningless once we lose the connection
        UndoManager.Clear();
        SketchEditState.Current = null;
        return KompasConnector.Disconnect();
    }

    [McpServerTool(Name = "kompas_get_info")]
    [Description("Get version and build information about the connected Kompas 3D instance.")]
    public static string GetInfo()
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        int maj = 0, min = 0, rel = 0, bld = 0;
        k.ksGetSystemVersion(out maj, out min, out rel, out bld);
        return $"Kompas 3D\nVersion: {maj}.{min}\nRelease: {rel}\nBuild: {bld}";
    }

    [McpServerTool(Name = "kompas_show_message")]
    [Description("Display a text message in the Kompas 3D message window.")]
    public static string ShowMessage(
        [Description("The text to display.")]
        string text)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        if (!k.ksMessage(text))
            return "ERROR: Kompas refused to display the message (message window unavailable).";
        return $"Message shown: {text}";
    }
}
