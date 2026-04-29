using System.ComponentModel;
using System.Text;
using ModelContextProtocol.Server;

namespace McpKompas.Tools;

[McpServerToolType]
public sealed class UndoTools
{
    [McpServerTool(Name = "kompas_undo")]
    [Description("Undo the last operation(s) performed via MCP tools. " +
                 "For 2D: deletes drawn objects. For 3D: removes the last feature from the part tree.")]
    public static string Undo(
        [Description("Number of operations to undo (default 1).")]
        int count = 1)
    {
        return UndoManager.Undo(count);
    }

    [McpServerTool(Name = "kompas_get_history")]
    [Description("Get the list of recent operations that can be undone.")]
    public static string GetHistory(
        [Description("Maximum number of operations to show (default 20).")]
        int limit = 20)
    {
        var history = UndoManager.GetHistory(limit);
        if (history.Count == 0)
            return "No operations recorded.";

        var sb = new StringBuilder();
        sb.AppendLine($"Last {history.Count} operation(s) (newest first):");
        foreach (var op in history)
            sb.AppendLine($"  #{op.Index} [{op.Domain}] {op.Description}");
        return sb.ToString().Trim();
    }

    [McpServerTool(Name = "kompas_clear_history")]
    [Description("Clear the undo history. Does not affect the document — only forgets tracked operations.")]
    public static string ClearHistory()
    {
        UndoManager.Clear();
        return "Undo history cleared.";
    }
}
