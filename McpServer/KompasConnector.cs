using Kompas6API5;
using KompasAPI7;
using System.Runtime.InteropServices;

namespace McpKompas;

/// <summary>
/// Manages a single COM connection to Kompas 3D.
/// Thread-safe singleton — all MCP tool calls share one instance.
/// </summary>
public static class KompasConnector
{
    private static KompasObject? _kompas;
    private static readonly object _lock = new();

    // P/Invoke: attach to a running COM server from the Running Object Table
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    // P/Invoke: bring Kompas window to foreground
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;
    private const int SW_MAXIMIZE = 3;

    private static KompasObject? TryGetRunningKompas()
    {
        try
        {
            if (CLSIDFromProgID("KOMPAS.Application.5", out Guid clsid) != 0)
                return null;
            if (GetActiveObject(ref clsid, IntPtr.Zero, out object obj) != 0)
                return null;
            return obj as KompasObject;
        }
        catch
        {
            return null;
        }
    }

    public static KompasObject? Instance
    {
        get { lock (_lock) { return _kompas; } }
    }

    /// <summary>
    /// Connect to a running Kompas instance, or start a new one.
    /// When <paramref name="startNew"/> is true and we are already connected to an
    /// instance, the existing connection is released first so a fresh process is started.
    /// </summary>
    public static (bool success, string message) Connect(bool startNew = false, bool visible = true)
    {
        lock (_lock)
        {
            // If already connected and caller wants a fresh instance, release the old one
            if (_kompas != null && startNew)
            {
                try { Marshal.ReleaseComObject(_kompas); } catch { }
                _kompas = null;
            }

            // If still connected, verify the object is alive and reuse it
            if (_kompas != null)
            {
                try
                {
                    int maj = 0, min = 0, rel = 0, bld = 0;
                    _kompas.ksGetSystemVersion(out maj, out min, out rel, out bld);
                    _kompas.Visible = visible;
                    return (true, $"Already connected (visible={visible}). Version: {maj}.{min}.{rel}.{bld}");
                }
                catch
                {
                    _kompas = null;
                }
            }

            // Try to attach to a running instance first (unless startNew was requested)
            if (!startNew)
            {
                _kompas = TryGetRunningKompas();
                if (_kompas != null)
                {
                    _kompas.Visible = visible;
                    int maj = 0, min = 0, rel = 0, bld = 0;
                    _kompas.ksGetSystemVersion(out maj, out min, out rel, out bld);
                    return (true, $"Connected to running Kompas (visible={visible}). Version: {maj}.{min}.{rel}.{bld}");
                }
            }

            // Start a new Kompas instance
            try
            {
                var type = Type.GetTypeFromProgID("KOMPAS.Application.5");
                if (type == null)
                    return (false, "Kompas is not installed. ProgID 'KOMPAS.Application.5' not found in registry.");

                var instance = Activator.CreateInstance(type) as KompasObject;
                if (instance == null)
                    return (false, "Activator.CreateInstance returned null or wrong type for KOMPAS.Application.5.");

                _kompas = instance;
                _kompas.Visible = visible;
                if (!_kompas.ActivateControllerAPI())
                {
                    // Release and bail — without ControllerAPI the rest of the bridge will misbehave
                    try { Marshal.ReleaseComObject(_kompas); } catch { }
                    _kompas = null;
                    return (false, "Kompas started but ActivateControllerAPI() failed. The Kompas instance is unusable.");
                }
                int maj = 0, min = 0, rel = 0, bld = 0;
                _kompas.ksGetSystemVersion(out maj, out min, out rel, out bld);
                return (true, $"Kompas started (visible={visible}). Version: {maj}.{min}.{rel}.{bld}");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to start Kompas: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Toggle Kompas window visibility.
    /// </summary>
    public static string SetVisible(bool visible)
    {
        lock (_lock)
        {
            if (_kompas == null)
                return "Not connected to Kompas.";
            _kompas.Visible = visible;
            return $"Kompas visibility set to {visible}.";
        }
    }

    /// <summary>
    /// Disconnect and release the COM object.
    /// </summary>
    public static string Disconnect()
    {
        lock (_lock)
        {
            if (_kompas == null)
                return "Not connected.";

            try { Marshal.ReleaseComObject(_kompas); } catch { }
            finally { _kompas = null; }
            return "Disconnected from Kompas.";
        }
    }

    /// <summary>
    /// Returns the connected instance or throws with a clear message.
    /// </summary>
    public static KompasObject Require()
    {
        var k = Instance;
        if (k == null)
            throw new InvalidOperationException(
                "Not connected to Kompas. Call kompas_connect first.");
        return k;
    }

    /// <summary>
    /// Convert the int HWND from Kompas into a sign-extended IntPtr.
    /// On 64-bit Windows window handles still fit in 32 bits by ABI, but the high
    /// bit must be sign-extended (negative HWND values are valid). Using `new IntPtr(int)`
    /// performs the correct sign-extension; a plain `(IntPtr)int` cast does the same in
    /// modern .NET, but using the constructor makes the intent explicit.
    /// </summary>
    private static IntPtr HwndToIntPtr(int hwnd) => new IntPtr(hwnd);

    /// <summary>
    /// Bring the Kompas main window to the foreground.
    /// </summary>
    public static void BringToFront()
    {
        lock (_lock)
        {
            if (_kompas == null) return;
            try
            {
                int hwnd = _kompas.ksGetHWindow();
                if (hwnd != 0)
                {
                    IntPtr h = HwndToIntPtr(hwnd);
                    ShowWindow(h, SW_RESTORE);
                    SetForegroundWindow(h);
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Maximize the Kompas main window to fill the entire screen.
    /// </summary>
    public static string Maximize()
    {
        lock (_lock)
        {
            if (_kompas == null)
                return "Not connected to Kompas.";
            try
            {
                int hwnd = _kompas.ksGetHWindow();
                if (hwnd == 0)
                    return "Cannot get Kompas window handle.";
                IntPtr h = HwndToIntPtr(hwnd);
                ShowWindow(h, SW_MAXIMIZE);
                SetForegroundWindow(h);
                return "Kompas window maximized.";
            }
            catch (Exception ex)
            {
                return $"Failed to maximize: {ex.Message}";
            }
        }
    }

    /// <summary>
    /// Refresh the active document window so newly drawn geometry becomes visible.
    /// </summary>
    public static void RefreshView()
    {
        lock (_lock)
        {
            if (_kompas == null) return;
            try
            {
                _kompas.ksRefreshActiveWindow();
            }
            catch { }
        }
    }

    /// <summary>
    /// Get the API7 IApplication interface for advanced operations.
    /// </summary>
    public static IApplication? GetApplication7()
    {
        lock (_lock)
        {
            if (_kompas == null) return null;
            try
            {
                return _kompas.ksGetApplication7() as IApplication;
            }
            catch
            {
                return null;
            }
        }
    }
}
