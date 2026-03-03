using System.Diagnostics;
using System.Text;
using System.Text.Json;
using BEPDesigner.WinUI.Models;

namespace BEPDesigner.WinUI.Services;

public static class PythonBridge
{
    public static async Task<string> GenerateAsync(BepPayload payload)
    {
        var exeDir = AppContext.BaseDirectory;
        var scriptPath = Path.Combine(exeDir, "python", "bep_engine.py");

        if (!File.Exists(scriptPath))
        {
            return "Python engine not found. Expected: " + scriptPath;
        }

        var psi = new ProcessStartInfo
        {
            FileName = "python",
            Arguments = '"' + scriptPath + '"',
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        using var process = new Process { StartInfo = psi };
        process.Start();

        var json = JsonSerializer.Serialize(payload);
        await process.StandardInput.WriteAsync(json);
        process.StandardInput.Close();

        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        var output = await outputTask;
        var error = await errorTask;

        if (process.ExitCode != 0)
        {
            return "Python generation failed:\n" + error;
        }

        return string.IsNullOrWhiteSpace(output) ? "No output from Python engine." : output.Trim();
    }
}
