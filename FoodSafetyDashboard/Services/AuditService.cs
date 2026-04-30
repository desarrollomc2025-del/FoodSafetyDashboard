using FoodSafetyDashboard.Data;
using FoodSafetyDashboard.Data.Models;
using Microsoft.EntityFrameworkCore;

namespace FoodSafetyDashboard.Services;

// ── DTOs ─────────────────────────────────────────────────────────────────────

public record KpiSummary(
    int TotalStores,
    double GlobalScore,
    int Approved,
    int NotApproved,
    int CriticalStores,
    Dictionary<string, double> SectionAverages
);

public record GroupSummary(
    string Group,
    int Count,
    double AvgScore,
    int Approved,
    int NotApproved,
    Dictionary<string, double> SectionAverages
);

public record StoreRow(
    int StoreId,
    long AuditId,
    string Group,
    DateTime? AuditDate,
    double Score,
    string Result,
    int CriticalViolations,
    int TotalViolations,
    string Points,
    Dictionary<string, double?> Sections
);

public record SystemicFinding(
    string FindingType,
    string Category,
    string Description,
    string Severity,
    string Action,
    List<int> AffectedStores,
    int StoreCount
);

public record ActionItem(
    int Index,
    string Action,
    string Stores,
    string Priority,
    string Deadline,
    string Responsible,
    string Status,
    string Notes
);

public record LocationSummary(
    string Name,
    int Count,
    double AvgScore,
    int Approved,
    int NotApproved
);

// ── Servicio ──────────────────────────────────────────────────────────────────

public class AuditService(IDbContextFactory<AppDbContext> dbFactory, IConfiguration config)
{
    private const double ApprovalThreshold = 80.0;

    private static readonly Dictionary<string, string> SectionMap = new()
    {
        ["Food Safety Risk Factors"] = "Fact. Riesgo",
        ["Cleanliness"]              = "Limpieza",
        ["Maintenance and Facility"] = "Mantenimiento",
        ["Storage"]                  = "Almacenaje",
        ["Knowledge and Compliance"] = "Conocimiento",
    };

    private static readonly string[] SectionCols =
        ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];

    private static readonly Dictionary<string, (string Severity, string Action)> SeverityMap = new()
    {
        ["temperature"]         = ("CRÍTICO",  "Revisión y calibración de equipos de refrigeración"),
        ["servsafe"]            = ("CRÍTICO",  "Certificar mínimo 1 empleado por tienda de forma inmediata"),
        ["pest"]                = ("CRÍTICO",  "Fumigación de emergencia en máximo 48 horas"),
        ["data_integrity"]      = ("CRÍTICO",  "Auditoría inmediata de registros y política de integridad"),
        ["cross_contamination"] = ("ALTO",     "Capacitación en prevención de contaminación cruzada"),
        ["sanitizer"]           = ("ALTO",     "Revisión de concentraciones y proceso de sanitización"),
        ["policy_gap"]          = ("ALTO",     "Actualizar manuales y distribuir a toda la red"),
        ["process_gap"]         = ("ALTO",     "Capacitación en procedimientos operativos estándar"),
        ["chemical_control"]    = ("ALTO",     "Capacitación en manejo de químicos y etiquetado"),
        ["equipment_damage"]    = ("MODERADO", "Plan de mantenimiento preventivo mensual"),
        ["cleanliness"]         = ("MODERADO", "Implementar control diario documentado"),
        ["personal_behavior"]   = ("MODERADO", "Capacitación y refuerzo de política de conducta"),
        ["other"]               = ("MODERADO", "Revisar y corregir según hallazgo específico"),
    };

    private static readonly Dictionary<string, string> CategoryMap = new()
    {
        ["temperature"]         = "Fact. Riesgo",
        ["servsafe"]            = "Fact. Riesgo",
        ["pest"]                = "Plagas / Riesgo",
        ["sanitizer"]           = "Limpieza",
        ["chemical_control"]    = "Almacenaje",
        ["equipment_damage"]    = "Mantenimiento",
        ["cleanliness"]         = "Limpieza",
        ["cross_contamination"] = "Fact. Riesgo",
        ["policy_gap"]          = "Conocimiento",
        ["process_gap"]         = "Conocimiento",
        ["data_integrity"]      = "Conocimiento",
        ["personal_behavior"]   = "Conocimiento",
        ["other"]               = "General",
    };

    private static readonly Dictionary<string, string> ResponsableMap = new()
    {
        ["pest"]                = "Gerente de Área",
        ["temperature"]         = "Mantenimiento",
        ["servsafe"]            = "RRHH / Operaciones",
        ["policy_gap"]          = "Calidad / Operaciones",
        ["process_gap"]         = "Calidad",
        ["equipment_damage"]    = "Mantenimiento",
        ["cleanliness"]         = "Gerentes de tienda",
        ["cross_contamination"] = "Calidad",
        ["sanitizer"]           = "Operaciones",
        ["personal_behavior"]   = "RRHH / Calidad",
        ["data_integrity"]      = "Calidad / Auditoría",
        ["chemical_control"]    = "Operaciones",
        ["other"]               = "Operaciones",
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private async Task<Dictionary<int, string>> GetGroupMapAsync()
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var stores = await db.Stores.ToListAsync();
        return stores.ToDictionary(
            s => s.StoreId,
            s => NormalizeGroupName(s.StoreType));
    }

    private static string NormalizeGroupName(string? raw) =>
        (raw ?? "").Trim().ToUpperInvariant() == "DEX" ? "DEX" : "CLÁSICA";

    private static string MapSection(string? raw) =>
        raw is not null && SectionMap.TryGetValue(raw, out var mapped) ? mapped : raw ?? "";

    private static double? GetSectionScore(IEnumerable<AuditSection> sections, string target) =>
        sections.FirstOrDefault(s => MapSection(s.SectionName) == target)?.SectionScore is decimal d ? (double)d : null;

    // ── Consultas públicas ────────────────────────────────────────────────────

    public async Task<List<StoreRow>> GetStoreRowsAsync(int? year = null, int? month = null)
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var groups = await GetGroupMapAsync();
        var query = db.Audits
            .Include(a => a.Sections)
            .AsQueryable();

        if (year.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Year == year.Value);
        if (month.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Month == month.Value);

        var audits = await query.ToListAsync();

        return audits
            .Select(a =>
            {
                var score = (double)(a.PercentageScore ?? 0);
                var grp = a.StoreId.HasValue && groups.TryGetValue(a.StoreId.Value, out var g) ? g : "CLÁSICA";
                var sections = SectionCols.ToDictionary(
                    col => col,
                    col => GetSectionScore(a.Sections, col));
                var points = a.PointsEarned.HasValue && a.PointsPossible.HasValue
                    ? $"{a.PointsEarned}/{a.PointsPossible}"
                    : "";

                var critViolations = a.Sections
                    .FirstOrDefault(s => s.SectionName == "Critical Violations")
                    ?.TotalViolations ?? a.CriticalViolations ?? 0;

                return new StoreRow(
                    a.StoreId ?? 0, a.AuditId, grp,
                    a.AuditStart, score,
                    score >= ApprovalThreshold ? "Aprobado" : "No aprobado",
                    critViolations,
                    a.TotalViolations ?? 0,
                    points, sections);
            })
            .OrderByDescending(r => r.Score)
            .ToList();
    }

    public async Task<KpiSummary> GetKpiSummaryAsync(int? year = null, int? month = null)
    {
        var rows = await GetStoreRowsAsync(year, month);
        if (rows.Count == 0)
            return new KpiSummary(0, 0, 0, 0, 0, []);

        var approved = rows.Count(r => r.Result == "Aprobado");
        var globalScore = rows.Average(r => r.Score);
        var criticalStores = rows.Count(r => r.CriticalViolations > 0);

        var secAvgs = SectionCols.ToDictionary(
            sec => sec,
            sec =>
            {
                var vals = rows.Select(r => r.Sections.GetValueOrDefault(sec)).OfType<double>().ToList();
                return vals.Count > 0 ? vals.Average() : 0.0;
            });

        return new KpiSummary(rows.Count, globalScore, approved, rows.Count - approved, criticalStores, secAvgs);
    }

    public async Task<List<GroupSummary>> GetGroupSummariesAsync(int? year = null, int? month = null)
    {
        var rows = await GetStoreRowsAsync(year, month);

        return rows
            .GroupBy(r => r.Group)
            .OrderBy(g => g.Key)
            .Select(g =>
            {
                var list = g.ToList();
                var approved = list.Count(r => r.Result == "Aprobado");
                var secAvgs = SectionCols.ToDictionary(
                    sec => sec,
                    sec =>
                    {
                        var vals = list.Select(r => r.Sections.GetValueOrDefault(sec)).OfType<double>().ToList();
                        return vals.Count > 0 ? vals.Average() : 0.0;
                    });
                return new GroupSummary(g.Key, list.Count, list.Average(r => r.Score),
                    approved, list.Count - approved, secAvgs);
            })
            .ToList();
    }

    public async Task<List<SystemicFinding>> GetSystemicFindingsAsync()
    {
        await using var db = await dbFactory.CreateDbContextAsync();

        var findings = await db.AuditFindings
            .Include(f => f.Audit)
            .Where(f => f.FindingType != null && f.FindingType != "other")
            .ToListAsync();

        return findings
            .GroupBy(f => f.FindingType!)
            .Select(g =>
            {
                var storeIds = g.Select(f => f.Audit?.StoreId ?? 0).Distinct().Where(s => s > 0).OrderBy(s => s).ToList();
                var topQ = g.GroupBy(f => f.QuestionText ?? "")
                             .OrderByDescending(q => q.Count())
                             .FirstOrDefault()?.Key ?? g.Key;
                var (severity, action) = SeverityMap.GetValueOrDefault(g.Key, ("MODERADO", "Revisar y corregir"));
                var category = CategoryMap.GetValueOrDefault(g.Key, "General");

                return new SystemicFinding(g.Key, category, topQ, severity, action, storeIds, storeIds.Count);
            })
            .OrderBy(f => f.Severity switch { "CRÍTICO" => 0, "ALTO" => 1, _ => 2 })
            .ThenByDescending(f => f.StoreCount)
            .ToList();
    }

    public async Task<List<ActionItem>> GetActionPlanAsync()
    {
        var rows = await GetStoreRowsAsync();
        await using var db = await dbFactory.CreateDbContextAsync();

        var items = new List<ActionItem>();
        int idx = 1;

        // Acciones inmediatas para tiendas con violaciones críticas
        var criticalStores = rows.Where(r => r.CriticalViolations > 0).OrderByDescending(r => r.CriticalViolations);
        foreach (var sr in criticalStores)
        {
            var pestFindings = await db.AuditFindings
                .Where(f => f.AuditId == sr.AuditId && f.FindingType == "pest")
                .FirstOrDefaultAsync();
            var tempFindings = await db.AuditFindings
                .Where(f => f.AuditId == sr.AuditId && f.FindingType == "temperature")
                .FirstOrDefaultAsync();

            if (pestFindings != null)
            {
                var note = (pestFindings.CommentText ?? "Presencia activa confirmada en auditoría.");
                if (note.Length > 120) note = note[..120];
                items.Add(new ActionItem(idx++, $"Fumigación de emergencia — {sr.StoreId}",
                    sr.StoreId.ToString(), "INMEDIATO", "0–48 horas",
                    $"Gerente Área {sr.Group}", "Pendiente", note));
            }
            if (tempFindings != null)
            {
                var note = (tempFindings.CommentText ?? "Temperatura TCS fuera de rango registrada.");
                if (note.Length > 120) note = note[..120];
                items.Add(new ActionItem(idx++, $"Corrección de temperatura — {sr.StoreId}",
                    sr.StoreId.ToString(), "INMEDIATO", "0–48 horas",
                    $"Gerente Área {sr.Group}", "Acción iniciada", note));
            }
        }

        // Acciones sistémicas
        var systemic = await GetSystemicFindingsAsync();
        var plazos = new Dictionary<string, string>
        {
            ["CRÍTICO"]  = "0–48 horas",
            ["ALTO"]     = "1–2 semanas",
            ["MODERADO"] = "2–4 semanas",
        };

        foreach (var sf in systemic)
        {
            var storesStr = sf.StoreCount <= 5
                ? string.Join(", ", sf.AffectedStores)
                : $"{sf.StoreCount} tiendas";

            items.Add(new ActionItem(idx++, sf.Action, storesStr, sf.Severity,
                plazos.GetValueOrDefault(sf.Severity, "A definir"),
                ResponsableMap.GetValueOrDefault(sf.FindingType, "Operaciones"),
                "Pendiente",
                sf.Description.Length > 120 ? sf.Description[..120] : sf.Description));
        }

        return items;
    }

    public async Task<List<LocationSummary>> GetDepartamentoSummariesAsync(int? year = null, int? month = null)
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var query = db.Audits.AsQueryable();

        if (year.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Year == year.Value);
        if (month.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Month == month.Value);

        var audits = await query.ToListAsync();

        return audits
            .Where(a => !string.IsNullOrWhiteSpace(a.Departamento))
            .GroupBy(a => a.Departamento!)
            .Select(g =>
            {
                var scores = g.Select(a => (double)(a.PercentageScore ?? 0)).ToList();
                var avg = scores.Count > 0 ? scores.Average() : 0;
                var approved = scores.Count(s => s >= ApprovalThreshold);
                return new LocationSummary(g.Key, g.Count(), avg, approved, g.Count() - approved);
            })
            .OrderByDescending(l => l.Count)
            .ThenBy(l => l.Name)
            .ToList();
    }

    public async Task<List<LocationSummary>> GetMunicipioSummariesAsync(int? year = null, int? month = null)
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var query = db.Audits.AsQueryable();

        if (year.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Year == year.Value);
        if (month.HasValue)
            query = query.Where(a => a.AuditStart.HasValue && a.AuditStart.Value.Month == month.Value);

        var audits = await query.ToListAsync();

        return audits
            .Where(a => !string.IsNullOrWhiteSpace(a.Municipio))
            .GroupBy(a => a.Municipio!)
            .Select(g =>
            {
                var scores = g.Select(a => (double)(a.PercentageScore ?? 0)).ToList();
                var avg = scores.Count > 0 ? scores.Average() : 0;
                var approved = scores.Count(s => s >= ApprovalThreshold);
                return new LocationSummary(g.Key, g.Count(), avg, approved, g.Count() - approved);
            })
            .OrderByDescending(l => l.Count)
            .ThenBy(l => l.Name)
            .ToList();
    }

    public async Task<List<int>> GetAvailableYearsAsync()
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var dates = await db.Audits
            .Where(a => a.AuditStart.HasValue)
            .Select(a => a.AuditStart!.Value)
            .ToListAsync();
        return dates
            .Select(d => d.Year)
            .Distinct()
            .OrderByDescending(y => y)
            .ToList();
    }

    public async Task<List<int>> GetAvailableMonthsAsync(int? year = null)
    {
        await using var db = await dbFactory.CreateDbContextAsync();
        var query = db.Audits.Where(a => a.AuditStart.HasValue);
        if (year.HasValue)
            query = query.Where(a => a.AuditStart!.Value.Year == year.Value);
        var dates = await query.Select(a => a.AuditStart!.Value).ToListAsync();
        return dates
            .Select(d => d.Month)
            .Distinct()
            .OrderBy(m => m)
            .ToList();
    }
}
