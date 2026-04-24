using System.Diagnostics;
using System.Text;
using Microsoft.AspNetCore.Components.Forms;

namespace FoodSafetyDashboard.Services;

public class PdfProcessingService
{
    private readonly string _pdfsFolder;
    private readonly string _scriptPath;
    private readonly string _pythonExe;
    private readonly string _pyOdbcConnStr;

    public PdfProcessingService(IConfiguration config, IWebHostEnvironment env)
    {
        var root = env.ContentRootPath;
        var section = config.GetSection("PdfProcessing");

        _pdfsFolder = Path.GetFullPath(Path.Combine(root, section["PdfsFolder"] ?? @"..\pdfs"));
        _scriptPath = Path.GetFullPath(Path.Combine(root, section["ScriptPath"] ?? @"..\parse_food_audits.py"));
        _pythonExe = section["PythonExe"] ?? "python";
        _pyOdbcConnStr = section["PyOdbcConnStr"] ?? string.Empty;
    }

    public string PdfsInPath => Path.Combine(_pdfsFolder, "in");

    public async Task SaveFileAsync(IBrowserFile file)
    {
        Directory.CreateDirectory(PdfsInPath);
        var dest = Path.Combine(PdfsInPath, Path.GetFileName(file.Name));
        await using var fs = new FileStream(dest, FileMode.Create);
        await file.OpenReadStream(maxAllowedSize: 50 * 1024 * 1024).CopyToAsync(fs);
    }

    public IReadOnlyList<string> GetPendingFiles()
    {
        if (!Directory.Exists(PdfsInPath)) return [];
        return [.. Directory.GetFiles(PdfsInPath, "*.pdf")
            .Select(Path.GetFileName)
            .Where(n => n != null)
            .Select(n => n!)
            .OrderBy(n => n)];
    }

    public async Task<(string output, int exitCode)> RunParserAsync(CancellationToken ct = default)
    {
        var workDir = Path.GetDirectoryName(_scriptPath) ?? ".";
        var psi = new ProcessStartInfo
        {
            FileName = _pythonExe,
            Arguments = $"\"{_scriptPath}\" --db --pdfs \"{_pdfsFolder}\"",
            WorkingDirectory = workDir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        psi.Environment["FOODSAFETY_CONN"] = _pyOdbcConnStr;

        using var proc = new Process { StartInfo = psi };
        var sb = new StringBuilder();
        proc.OutputDataReceived += (_, e) => { if (e.Data != null) sb.AppendLine(e.Data); };
        proc.ErrorDataReceived += (_, e) => { if (e.Data != null) sb.AppendLine("[ERR] " + e.Data); };

        proc.Start();
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();
        await proc.WaitForExitAsync(ct);

        return (sb.ToString(), proc.ExitCode);
    }
}
