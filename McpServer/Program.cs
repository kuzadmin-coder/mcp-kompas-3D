using McpKompas.Tools;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.AspNetCore;
using ModelContextProtocol.Server;

// MCP server for Kompas 3D v24
// Transport: stdio (default) or HTTP/SSE (--http flag)
//
// Usage:
//   McpKompas.exe              — stdio mode (Claude Desktop, VS Code)
//   McpKompas.exe --http       — HTTP/SSE on http://0.0.0.0:3001
//   McpKompas.exe --http 8080  — HTTP/SSE on custom port

bool useHttp = args.Contains("--http");
int port = 3001;

if (useHttp)
{
    int idx = Array.IndexOf(args, "--http");
    if (idx + 1 < args.Length && int.TryParse(args[idx + 1], out int customPort))
        port = customPort;
}

if (useHttp)
{
    // ── HTTP/SSE mode — accessible over the network ──
    var builder = WebApplication.CreateBuilder(args);

    builder.Logging.SetMinimumLevel(LogLevel.Information);

    builder.Services.AddCors(options =>
        options.AddDefaultPolicy(policy =>
            policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod()));

    builder.Services
        .AddMcpServer()
        .WithHttpTransport(options => options.Stateless = true)
        .WithTools<ConnectionTools>()
        .WithTools<DocumentTools>()
        .WithTools<Drawing2DTools>()
        .WithTools<Drawing3DTools>()
        .WithTools<UndoTools>();

    builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

    var app = builder.Build();
    app.UseCors();
    app.MapMcp("/mcp");

    Console.WriteLine($"MCP Kompas 3D server listening on http://0.0.0.0:{port}");
    Console.WriteLine("MCP endpoint: /mcp");

    await app.RunAsync();
}
else
{
    // ── Stdio mode — local, for Claude Desktop / VS Code ──
    var builder = Host.CreateApplicationBuilder(args);

    builder.Logging.SetMinimumLevel(LogLevel.Warning);
    builder.Logging.AddConsole(opts => opts.LogToStandardErrorThreshold = LogLevel.Warning);

    builder.Services
        .AddMcpServer()
        .WithStdioServerTransport()
        .WithTools<ConnectionTools>()
        .WithTools<DocumentTools>()
        .WithTools<Drawing2DTools>()
        .WithTools<Drawing3DTools>()
        .WithTools<UndoTools>();

    await builder.Build().RunAsync();
}
