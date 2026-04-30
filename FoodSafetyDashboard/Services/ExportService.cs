using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace FoodSafetyDashboard.Services;

public class ExportService(AuditService auditService)
{
    private static readonly string[] SectionCols =
        ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];

    private static readonly string[] GroupColors =
        ["#2E75B6", "#ED7D31", "#70AD47", "#FFC000", "#7030A0", "#FF4136"];

    private static readonly string[] SectionColors =
        ["#E53935", "#039BE5", "#FB8C00", "#8E24AA", "#43A047"];

    private static readonly Color ColNavy     = Color.FromArgb(27, 42, 74);     // #1B2A4A
    private static readonly Color ColNavy2    = Color.FromArgb(30, 58, 110);    // #1E3A6E subtitle + Fact.Riesgo/Mant.
    private static readonly Color ColTeal     = Color.FromArgb(0, 109, 119);    // #006D77
    private static readonly Color ColHeader   = Color.FromArgb(46, 117, 182);   // #2E75B6 (otras hojas)
    private static readonly Color ColSlate    = Color.FromArgb(45, 55, 72);     // #2D3748 tabla grupos row9
    private static readonly Color ColDanger   = Color.FromArgb(192, 0, 0);
    private static readonly Color ColGreen    = Color.FromArgb(26, 122, 74);    // #1A7A4A Aprobadas
    private static readonly Color ColAmber    = Color.FromArgb(226, 156, 59);   // #E29C3B Críticos
    private static readonly Color ColEarth    = Color.FromArgb(139, 94, 60);    // #8B5E3C Conocimiento
    private static readonly Color ColValBg    = Color.FromArgb(244, 246, 250);  // #F4F6FA fondo valores KPI
    private static readonly Color ColApproved = Color.FromArgb(198, 239, 206);
    private static readonly Color ColFailed   = Color.FromArgb(255, 199, 206);
    private static readonly Color ColStripe   = Color.FromArgb(235, 241, 250);

    private static readonly Dictionary<string, Color> SectionKpiColors = new()
    {
        ["Fact. Riesgo"]  = Color.FromArgb(30, 58, 110),   // ColNavy2
        ["Limpieza"]      = Color.FromArgb(0, 109, 119),   // ColTeal
        ["Mantenimiento"] = Color.FromArgb(30, 58, 110),   // ColNavy2
        ["Almacenaje"]    = Color.FromArgb(27, 42, 74),    // ColNavy
        ["Conocimiento"]  = Color.FromArgb(139, 94, 60),   // ColEarth
    };

    static ExportService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<byte[]> ExportStoreDetailAsync(int? year = null, int? month = null)
    {
        var rows       = await auditService.GetStoreRowsAsync(year, month);
        var groups     = await auditService.GetGroupSummariesAsync(year, month);
        var kpis       = await auditService.GetKpiSummaryAsync(year, month);
        var actionPlan = await auditService.GetActionPlanAsync();
        var systemic   = await auditService.GetSystemicFindingsAsync();

        using var pkg = new ExcelPackage();

        AddResumenEjecutivoSheet(pkg, kpis, groups, rows, systemic);
        AddDetalleSheet(pkg, rows);
        AddPlanAccionSheet(pkg, actionPlan);
        AddAnalisisGraficoSheet(pkg, groups, kpis, rows);

        return await pkg.GetAsByteArrayAsync();
    }

    // ── Hoja 1: Resumen Ejecutivo ──────────────────────────────────────────────

    private static void AddResumenEjecutivoSheet(
        ExcelPackage pkg, KpiSummary kpis, List<GroupSummary> groups, List<StoreRow> rows,
        List<SystemicFinding> systemic)
    {
        var ws = pkg.Workbook.Worksheets.Add("Resumen Ejecutivo");
        ws.TabColor = ColNavy;
        ws.View.ShowGridLines = false;

        // kpiCols: tiendas(1) + global(1) + grupos + aprobadas(1) + noAprobadas(1) + críticos(1) + secciones(5)
        int kpiCols   = 5 + groups.Count + SectionCols.Length;
        int totalCols = Math.Max(kpiCols, 11);

        double[] colWidths = [14, 50, 22, 13, 16, 13, 20, 13, 12, 13, 13, 13];
        for (int i = 0; i < Math.Min(colWidths.Length, totalCols); i++)
            ws.Column(i + 1).Width = colWidths[i];

        // ── Título ────────────────────────────────────────────────────────────
        ws.Row(1).Height = 36;
        ws.Cells[1, 1, 1, totalCols].Merge = true;
        ws.Cells[1, 1].Value = "REPORTE EJECUTIVO — GESTIÓN DE CALIDAD FSE INTERNACIONAL";
        StyleTitle(ws.Cells[1, 1]);

        // ── Subtítulo ─────────────────────────────────────────────────────────
        ws.Row(2).Height = 22;
        var dates = rows.Where(r => r.AuditDate.HasValue).Select(r => r.AuditDate!.Value).ToList();
        var groupNames = string.Join(" y ", groups.Select(g => $"{g.Group} ({g.Count})"));
        string subtitle;
        if (dates.Count > 0)
        {
            var cult = new CultureInfo("es");
            subtitle = $"Período: {dates.Min().ToString("dd MMMM yyyy", cult)} – {dates.Max().ToString("dd MMMM yyyy", cult)}" +
                       $"  |  {kpis.TotalStores} Tiendas  |  Grupos: {groupNames}";
        }
        else
            subtitle = $"{kpis.TotalStores} Tiendas  |  Grupos: {groupNames}";

        ws.Cells[2, 1, 2, totalCols].Merge = true;
        ws.Cells[2, 1].Value = subtitle;
        FillCell(ws.Cells[2, 1], ColNavy2);
        ws.Cells[2, 1].Style.Font.Italic = true;
        ws.Cells[2, 1].Style.Font.Color.SetColor(Color.White);
        ws.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        ws.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        // ── KPI Section header ────────────────────────────────────────────────
        ws.Row(4).Height = 24;
        ws.Cells[4, 1, 4, totalCols].Merge = true;
        ws.Cells[4, 1].Value = "INDICADORES CLAVE DE DESEMPEÑO";
        StyleSectionHeader(ws.Cells[4, 1]);

        // ── KPI headers (row 5) ───────────────────────────────────────────────
        ws.Row(5).Height = 30;
        int col = 1;
        SetKpiHeader(ws.Cells[5, col++], "Tiendas Evaluadas", ColNavy);
        SetKpiHeader(ws.Cells[5, col++], "Nota Global",       ColTeal);
        foreach (var g in groups)
            SetKpiHeader(ws.Cells[5, col++], $"Grupo {g.Group}", g.AvgScore >= 80 ? ColTeal : ColDanger);
        SetKpiHeader(ws.Cells[5, col++], "Aprobadas",            ColGreen);
        SetKpiHeader(ws.Cells[5, col++], "No Aprobadas",         ColDanger);
        SetKpiHeader(ws.Cells[5, col++], "Violaciones Críticas", ColAmber);
        foreach (var sec in SectionCols)
            SetKpiHeader(ws.Cells[5, col++], sec, SectionKpiColors.GetValueOrDefault(sec, ColHeader));

        // ── KPI values (row 6) ────────────────────────────────────────────────
        ws.Row(6).Height = 36;
        col = 1;
        SetKpiValue(ws.Cells[6, col++], kpis.TotalStores,                           null,    Color.FromArgb(0, 70, 127));
        SetKpiValue(ws.Cells[6, col++], kpis.GlobalScore / 100.0,                   "0.0%",  ColTeal);
        foreach (var g in groups)
            SetKpiValue(ws.Cells[6, col++], g.AvgScore / 100.0,                     "0.0%",  g.AvgScore >= 80 ? ColTeal : ColDanger);
        SetKpiValue(ws.Cells[6, col++], $"{kpis.Approved} / {kpis.TotalStores}",    null,    ColTeal);
        SetKpiValue(ws.Cells[6, col++], $"{kpis.NotApproved} / {kpis.TotalStores}", null,    ColDanger);
        SetKpiValue(ws.Cells[6, col++], rows.Sum(r => r.CriticalViolations),        null,    Color.FromArgb(120, 40, 40));
        foreach (var sec in SectionCols)
            SetKpiValue(ws.Cells[6, col++], kpis.SectionAverages.GetValueOrDefault(sec) / 100.0, "0.0%", ColHeader);

        // ── Resultados por Grupo ──────────────────────────────────────────────
        ws.Row(8).Height = 24;
        ws.Cells[8, 1, 8, totalCols].Merge = true;
        ws.Cells[8, 1].Value = "RESULTADOS POR GRUPO";
        StyleSectionHeader(ws.Cells[8, 1]);

        ws.Row(9).Height = 26;
        string[] grpHdrs = ["Grupo", "Tiendas", "Nota %", "Fact. Riesgo", "Limpieza",
            "Mantenimiento", "Almacenaje", "Conocimiento", "Aprobadas", "No Aprobadas", "Estado"];
        for (int i = 0; i < grpHdrs.Length; i++)
            ws.Cells[9, i + 1].Value = grpHdrs[i];
        StyleHeader(ws.Cells[9, 1, 9, grpHdrs.Length], ColSlate);

        int gr = 10;
        foreach (var g in groups)
        {
            ws.Row(gr).Height = 22;
            ws.Cells[gr, 1].Value = g.Group;
            ws.Cells[gr, 2].Value = g.Count;
            ws.Cells[gr, 3].Value = g.AvgScore / 100.0;
            ws.Cells[gr, 3].Style.Numberformat.Format = "0.0%";
            for (int i = 0; i < SectionCols.Length; i++)
            {
                ws.Cells[gr, 4 + i].Value = g.SectionAverages.GetValueOrDefault(SectionCols[i]) / 100.0;
                ws.Cells[gr, 4 + i].Style.Numberformat.Format = "0.0%";
            }
            ws.Cells[gr, 9].Value  = g.Approved;
            ws.Cells[gr, 10].Value = g.NotApproved;

            var estado = g.NotApproved == 0 ? "✓ Todo aprobado" : $"⚠ {g.NotApproved} no aprobada(s)";
            ws.Cells[gr, 11].Value = estado;
            FillCell(ws.Cells[gr, 11], g.NotApproved == 0 ? ColApproved : ColFailed);

            if (gr % 2 == 0) StripRow(ws, gr, 1, 10);
            gr++;
        }

        // ── Fila de totales ───────────────────────────────────────────────────
        ws.Row(gr).Height = 24;
        ws.Cells[gr, 1].Value = "TOTAL";
        ws.Cells[gr, 2].Value = kpis.TotalStores;
        ws.Cells[gr, 3].Value = kpis.GlobalScore / 100.0;
        ws.Cells[gr, 3].Style.Numberformat.Format = "0.0%";
        for (int i = 0; i < SectionCols.Length; i++)
        {
            ws.Cells[gr, 4 + i].Value = kpis.SectionAverages.GetValueOrDefault(SectionCols[i]) / 100.0;
            ws.Cells[gr, 4 + i].Style.Numberformat.Format = "0.0%";
        }
        ws.Cells[gr, 9].Value  = kpis.Approved;
        ws.Cells[gr, 10].Value = kpis.NotApproved;
        ws.Cells[gr, 11].Value = kpis.NotApproved == 0 ? "✓ Todo aprobado" : $"⚠ {kpis.NotApproved} no aprobada(s)";
        FillCell(ws.Cells[gr, 11], kpis.NotApproved == 0 ? ColApproved : ColFailed);

        var totalRange = ws.Cells[gr, 1, gr, 11];
        totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
        totalRange.Style.Fill.BackgroundColor.SetColor(ColSlate);
        totalRange.Style.Font.Color.SetColor(Color.White);
        totalRange.Style.Font.Bold = true;
        ws.Cells[gr, 11].Style.Font.Color.SetColor(Color.Black);

        gr++;

        // ── Top / Bottom Performers ────────────────────────────────────────────
        var top5    = rows.OrderByDescending(r => r.Score).Take(5).ToList();
        var bottom5 = rows.OrderBy(r => r.Score).Take(5).ToList();
        int perfRow = gr + 2;
        int bRow    = gr + 2;

        // Top 5 (cols 1–5)
        ws.Row(perfRow).Height = 22;
        ws.Cells[perfRow, 1, perfRow, 5].Merge = true;
        ws.Cells[perfRow, 1].Value = "TOP 5 — MEJORES TIENDAS";
        StyleSectionHeader(ws.Cells[perfRow, 1]);
        perfRow++;
        string[] pHdrs = ["Tienda", "Grupo", "Nota %", "Resultado"];
        for (int i = 0; i < pHdrs.Length; i++) ws.Cells[perfRow, i + 1].Value = pHdrs[i];
        StyleHeader(ws.Cells[perfRow, 1, perfRow, 4]);
        perfRow++;
        foreach (var s in top5)
        {
            ws.Cells[perfRow, 1].Value = s.StoreId;
            ws.Cells[perfRow, 2].Value = s.Group;
            ws.Cells[perfRow, 3].Value = s.Score / 100.0;
            ws.Cells[perfRow, 3].Style.Numberformat.Format = "0.0%";
            ws.Cells[perfRow, 4].Value = s.Result;
            FillCell(ws.Cells[perfRow, 4], ColApproved);
            perfRow++;
        }

        // Bottom 5 (cols 7–11)
        ws.Row(bRow).Height = 22;
        ws.Cells[bRow, 7, bRow, 11].Merge = true;
        ws.Cells[bRow, 7].Value = "BOTTOM 5 — TIENDAS CON MENOR NOTA";
        StyleSectionHeader(ws.Cells[bRow, 7]);
        bRow++;
        for (int i = 0; i < pHdrs.Length; i++) ws.Cells[bRow, 7 + i].Value = pHdrs[i];
        StyleHeader(ws.Cells[bRow, 7, bRow, 10]);
        bRow++;
        foreach (var s in bottom5)
        {
            ws.Cells[bRow, 7].Value  = s.StoreId;
            ws.Cells[bRow, 8].Value  = s.Group;
            ws.Cells[bRow, 9].Value  = s.Score / 100.0;
            ws.Cells[bRow, 9].Style.Numberformat.Format = "0.0%";
            ws.Cells[bRow, 10].Value = s.Result;
            FillCell(ws.Cells[bRow, 10], s.Result == "Aprobado" ? ColApproved : ColFailed);
            bRow++;
        }

        // ── Análisis Sistémico ─────────────────────────────────────────────────
        if (systemic.Count == 0) return;

        int sysRow = Math.Max(perfRow, bRow) + 2;

        ws.Row(sysRow).Height = 24;
        ws.Cells[sysRow, 1, sysRow, totalCols].Merge = true;
        ws.Cells[sysRow, 1].Value = "ANÁLISIS SISTÉMICO DE HALLAZGOS";
        StyleSectionHeader(ws.Cells[sysRow, 1]);
        sysRow++;

        string[] sysHdrs = ["#", "Categoría", "Hallazgo Principal", "Severidad",
            "Acción Recomendada", "Tiendas Afectadas", "# Tiendas"];
        for (int i = 0; i < sysHdrs.Length; i++)
            ws.Cells[sysRow, i + 1].Value = sysHdrs[i];
        StyleHeader(ws.Cells[sysRow, 1, sysRow, sysHdrs.Length]);
        ws.Column(3).Width = 45;
        ws.Column(5).Width = 40;
        ws.Column(6).Width = 30;
        sysRow++;

        int sysIdx = 1;
        foreach (var sf in systemic)
        {
            ws.Row(sysRow).Height = 30;
            ws.Cells[sysRow, 1].Value = sysIdx++;
            ws.Cells[sysRow, 2].Value = sf.Category;
            ws.Cells[sysRow, 3].Value = sf.Description.Length > 100
                ? sf.Description[..100] + "…"
                : sf.Description;
            ws.Cells[sysRow, 3].Style.WrapText = true;
            ws.Cells[sysRow, 4].Value = sf.Severity;
            ws.Cells[sysRow, 5].Value = sf.Action;
            ws.Cells[sysRow, 5].Style.WrapText = true;
            ws.Cells[sysRow, 6].Value = sf.StoreCount <= 6
                ? string.Join(", ", sf.AffectedStores)
                : $"{sf.StoreCount} tiendas";
            ws.Cells[sysRow, 7].Value = sf.StoreCount;
            ws.Cells[sysRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var (sevColor, sevFg) = sf.Severity switch
            {
                "CRÍTICO"  => (ColDanger,                    Color.White),
                "ALTO"     => (Color.FromArgb(230, 100, 0),  Color.White),
                "MODERADO" => (Color.FromArgb(255, 199, 0),  Color.Black),
                _          => (Color.FromArgb(180, 180, 180), Color.Black),
            };
            FillCell(ws.Cells[sysRow, 4], sevColor);
            ws.Cells[sysRow, 4].Style.Font.Color.SetColor(sevFg);
            ws.Cells[sysRow, 4].Style.Font.Bold = true;
            ws.Cells[sysRow, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            if (sysRow % 2 == 0) StripRow(ws, sysRow, 1, sysHdrs.Length);
            sysRow++;
        }
    }

    // ── Hoja 2: Detalle por Tienda ─────────────────────────────────────────────

    private static void AddDetalleSheet(ExcelPackage pkg, List<StoreRow> rows)
    {
        var ws = pkg.Workbook.Worksheets.Add("Detalle por Tienda");
        ws.TabColor = ColTeal;
        ws.View.ShowGridLines = false;

        ws.Row(1).Height = 30;
        ws.Cells[1, 1, 1, 13].Merge = true;
        ws.Cells[1, 1].Value = "DETALLE DE EVALUACIÓN POR TIENDA";
        StyleTitle(ws.Cells[1, 1]);

        ws.Row(2).Height = 30;
        string[] headers = ["Tienda", "Grupo", "Fecha Eval.", "Nota %", "Resultado", "Críticos",
            "Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento", "# Fallos", "Puntos"];
        for (int c = 0; c < headers.Length; c++)
            ws.Cells[2, c + 1].Value = headers[c];
        StyleHeader(ws.Cells[2, 1, 2, headers.Length]);

        int r = 3;
        foreach (var row in rows)
        {
            ws.Row(r).Height = 20;
            ws.Cells[r, 1].Value = row.StoreId;
            ws.Cells[r, 2].Value = row.Group;
            ws.Cells[r, 3].Value = row.AuditDate?.ToString("yyyy-MM-dd") ?? "";
            ws.Cells[r, 4].Value = row.Score / 100.0;
            ws.Cells[r, 4].Style.Numberformat.Format = "0.0%";
            ws.Cells[r, 5].Value = row.Result;
            ws.Cells[r, 6].Value = row.CriticalViolations;

            for (int i = 0; i < SectionCols.Length; i++)
            {
                var val = row.Sections.GetValueOrDefault(SectionCols[i]);
                if (val.HasValue)
                {
                    ws.Cells[r, 7 + i].Value = val.Value / 100.0;
                    ws.Cells[r, 7 + i].Style.Numberformat.Format = "0.0%";
                }
            }

            ws.Cells[r, 12].Value = row.TotalViolations;
            ws.Cells[r, 13].Value = row.Points;

            FillCell(ws.Cells[r, 5], row.Result == "Aprobado" ? ColApproved : ColFailed);
            if (r % 2 == 0) StripRow(ws, r, 1, 13);
            r++;
        }

        if (ws.Dimension != null)
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
    }

    // ── Hoja 3: Plan de Acción ─────────────────────────────────────────────────

    private static void AddPlanAccionSheet(ExcelPackage pkg, List<ActionItem> items)
    {
        var ws = pkg.Workbook.Worksheets.Add("Plan de Acción");
        ws.TabColor = Color.FromArgb(192, 57, 43);
        ws.View.ShowGridLines = false;

        ws.Column(1).Width = 4;
        ws.Column(2).Width = 50;
        ws.Column(3).Width = 22;
        ws.Column(4).Width = 13;
        ws.Column(5).Width = 16;
        ws.Column(6).Width = 20;
        ws.Column(7).Width = 13;
        ws.Column(8).Width = 38;

        ws.Row(1).Height = 30;
        ws.Cells[1, 1, 1, 8].Merge = true;
        ws.Cells[1, 1].Value = "PLAN DE ACCIÓN CORRECTIVO — FSE INTERNACIONAL";
        StyleTitle(ws.Cells[1, 1]);

        ws.Row(2).Height = 28;
        string[] headers = ["#", "Acción", "Tiendas", "Prioridad", "Plazo", "Responsable", "Estado", "Observaciones"];
        for (int i = 0; i < headers.Length; i++)
            ws.Cells[2, i + 1].Value = headers[i];
        StyleHeader(ws.Cells[2, 1, 2, headers.Length]);

        int r = 3;
        bool odd = true;
        foreach (var item in items)
        {
            ws.Row(r).Height = 30;
            ws.Cells[r, 1].Value = item.Index;
            ws.Cells[r, 2].Value = item.Action;
            ws.Cells[r, 3].Value = item.Stores;
            ws.Cells[r, 4].Value = item.Priority;
            ws.Cells[r, 5].Value = item.Deadline;
            ws.Cells[r, 6].Value = item.Responsible;
            ws.Cells[r, 7].Value = item.Status;
            ws.Cells[r, 8].Value = item.Notes;

            ws.Cells[r, 2].Style.WrapText = true;
            ws.Cells[r, 8].Style.WrapText = true;

            // Priority chip
            var (priColor, priFg) = item.Priority switch
            {
                "INMEDIATO" or "CRÍTICO" => (ColDanger,                            Color.White),
                "ALTO"                   => (Color.FromArgb(230, 100, 0),          Color.White),
                "MODERADO"               => (Color.FromArgb(255, 199, 0),          Color.Black),
                _                        => (Color.FromArgb(180, 180, 180),        Color.Black),
            };
            FillCell(ws.Cells[r, 4], priColor);
            ws.Cells[r, 4].Style.Font.Color.SetColor(priFg);
            ws.Cells[r, 4].Style.Font.Bold = true;
            ws.Cells[r, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Status chip
            var statColor = item.Status switch
            {
                "Completado"                     => ColApproved,
                "En curso" or "Acción iniciada"  => Color.FromArgb(255, 235, 156),
                _                                => Color.FromArgb(242, 242, 242),
            };
            FillCell(ws.Cells[r, 7], statColor);
            ws.Cells[r, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Alternating background for non-colored columns
            var rowBg = odd ? Color.White : Color.FromArgb(252, 245, 245);
            foreach (int c in new[] { 1, 2, 3, 5, 6, 8 })
            {
                ws.Cells[r, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[r, c].Style.Fill.BackgroundColor.SetColor(rowBg);
            }

            odd = !odd;
            r++;
        }
    }

    // ── Hoja 4: Análisis Gráfico ───────────────────────────────────────────────

    private void AddAnalisisGraficoSheet(
        ExcelPackage pkg, List<GroupSummary> groups, KpiSummary kpis, List<StoreRow> rows)
    {
        var ws = pkg.Workbook.Worksheets.Add("Análisis Gráfico");
        ws.TabColor = ColHeader;
        ws.View.ShowGridLines = false;

        // Gráfica 1 — Nota por Grupo
        int g1Row = 1;
        WriteTableHeader(ws, g1Row, 1, ["Grupo", "Nota %"]);
        for (int i = 0; i < groups.Count; i++)
        {
            ws.Cells[g1Row + 1 + i, 1].Value = groups[i].Group;
            ws.Cells[g1Row + 1 + i, 2].Value = Math.Round(groups[i].AvgScore, 1);
        }
        int g1End = g1Row + groups.Count;
        AddColumnChart(ws, "chart_grupos", "Nota Global por Grupo",
            ws.Cells[g1Row + 1, 2, g1End, 2], ws.Cells[g1Row + 1, 1, g1End, 1],
            "Nota %", row: g1Row - 1, col: 3, width: 500, height: 280,
            distributedColors: GroupColors, itemCount: groups.Count);

        // Gráfica 2 — Promedio por Sección
        var sections = kpis.SectionAverages.ToList();
        int g2Row = g1End + 3;
        WriteTableHeader(ws, g2Row, 1, ["Sección", "Promedio %"]);
        for (int i = 0; i < sections.Count; i++)
        {
            ws.Cells[g2Row + 1 + i, 1].Value = sections[i].Key;
            ws.Cells[g2Row + 1 + i, 2].Value = Math.Round(sections[i].Value, 1);
        }
        int g2End = g2Row + sections.Count;
        AddColumnChart(ws, "chart_secciones", "Promedio por Sección",
            ws.Cells[g2Row + 1, 2, g2End, 2], ws.Cells[g2Row + 1, 1, g2End, 1],
            "Score %", row: g2Row - 1, col: 3, width: 540, height: 280,
            distributedColors: SectionColors, itemCount: sections.Count);

        // Gráfica 3 — CLÁSICA vs DEX por Sección
        var clasica = groups.FirstOrDefault(g => !g.Group.Equals("DEX", StringComparison.OrdinalIgnoreCase));
        var dex     = groups.FirstOrDefault(g =>  g.Group.Equals("DEX", StringComparison.OrdinalIgnoreCase));
        int g3Row = g2End + 3;
        WriteTableHeader(ws, g3Row, 1, ["Sección", "CLÁSICA", "DEX"]);
        for (int i = 0; i < SectionCols.Length; i++)
        {
            ws.Cells[g3Row + 1 + i, 1].Value = SectionCols[i];
            ws.Cells[g3Row + 1 + i, 2].Value = Math.Round(clasica?.SectionAverages.GetValueOrDefault(SectionCols[i]) ?? 0, 1);
            ws.Cells[g3Row + 1 + i, 3].Value = Math.Round(dex?.SectionAverages.GetValueOrDefault(SectionCols[i]) ?? 0, 1);
        }
        int g3End = g3Row + SectionCols.Length;

        var chart3 = ws.Drawings.AddChart("chart_vs", eChartType.ColumnClustered) as ExcelBarChart;
        if (chart3 != null)
        {
            chart3.Title.Text = "Nota por Sección — CLÁSICA vs DEX";
            chart3.SetPosition(g3Row - 1, 0, 3, 0);
            chart3.SetSize(620, 300);
            var s3a = chart3.Series.Add(ws.Cells[g3Row + 1, 2, g3End, 2], ws.Cells[g3Row + 1, 1, g3End, 1]);
            s3a.Header = "CLÁSICA";
            SetSeriesColor(s3a, Color.FromArgb(46, 117, 182));
            var s3b = chart3.Series.Add(ws.Cells[g3Row + 1, 3, g3End, 3], ws.Cells[g3Row + 1, 1, g3End, 1]);
            s3b.Header = "DEX";
            SetSeriesColor(s3b, Color.FromArgb(237, 125, 49));
            chart3.YAxis.MinValue = 50;
            chart3.YAxis.MaxValue = 100;
            chart3.DataLabel.ShowValue = true;
        }

        // Gráfica 4 — Ranking de Tiendas (barra horizontal, todas las tiendas)
        var sorted = rows.OrderByDescending(r => r.Score).ToList();
        int g4Row = g3End + 3;
        WriteTableHeader(ws, g4Row, 1, ["Tienda", "Nota %", "Grupo"]);
        for (int i = 0; i < sorted.Count; i++)
        {
            ws.Cells[g4Row + 1 + i, 1].Value = sorted[i].StoreId.ToString();
            ws.Cells[g4Row + 1 + i, 2].Value = Math.Round(sorted[i].Score, 1);
            ws.Cells[g4Row + 1 + i, 3].Value = sorted[i].Group;
        }
        int g4End  = g4Row + sorted.Count;
        int rankH  = Math.Max(320, sorted.Count * 24 + 80);

        var chart4 = ws.Drawings.AddChart("chart_ranking", eChartType.BarClustered) as ExcelBarChart;
        if (chart4 != null)
        {
            chart4.Title.Text = "Nota por Tienda (mayor a menor)";
            chart4.SetPosition(g4Row - 1, 0, 3, 0);
            chart4.SetSize(700, rankH);

            var s4 = chart4.Series.Add(
                ws.Cells[g4Row + 1, 2, g4End, 2],
                ws.Cells[g4Row + 1, 1, g4End, 1]);
            s4.Header = "Nota %";

            // Color por grupo: azul=CLÁSICA, naranja=DEX
            if (s4 is ExcelBarChartSerie bs4)
            {
                for (int i = 0; i < sorted.Count; i++)
                {
                    var dp = bs4.DataPoints.Add(i);
                    dp.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                    var barColor = sorted[i].Group == "DEX"
                        ? Color.FromArgb(237, 125, 49)
                        : Color.FromArgb(46, 117, 182);
                    dp.Fill.SolidFill.Color.SetRgbColor(barColor);
                }
            }

            chart4.DataLabel.ShowValue = true;
            chart4.Legend.Remove();
        }

        ws.Column(1).AutoFit();
        ws.Column(2).AutoFit();
        ws.Column(3).AutoFit();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static void SetKpiHeader(ExcelRange cell, string label, Color bg)
    {
        cell.Value = label;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(bg);
        cell.Style.Font.Color.SetColor(Color.White);
        cell.Style.Font.Bold = true;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        cell.Style.WrapText = true;
    }

    private static void SetKpiValue(ExcelRange cell, object value, string? format, Color _)
    {
        cell.Value = value;
        if (format != null) cell.Style.Numberformat.Format = format;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(ColValBg);
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 14;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    private static void FillCell(ExcelRange cell, Color color)
    {
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(color);
    }

    private static void StripRow(ExcelWorksheet ws, int row, int fromCol, int toCol)
    {
        for (int c = fromCol; c <= toCol; c++)
        {
            if (ws.Cells[row, c].Style.Fill.PatternType == ExcelFillStyle.None)
            {
                ws.Cells[row, c].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, c].Style.Fill.BackgroundColor.SetColor(ColStripe);
            }
        }
    }

    private static void AddColumnChart(
        ExcelWorksheet ws, string name, string title,
        ExcelRange valueRange, ExcelRange categoryRange, string seriesHeader,
        int row, int col, int width, int height,
        Color seriesColor = default, string[]? distributedColors = null, int itemCount = 0)
    {
        var chart = ws.Drawings.AddChart(name, eChartType.ColumnClustered) as ExcelBarChart;
        if (chart is null) return;

        chart.Title.Text = title;
        chart.SetPosition(row, 0, col, 0);
        chart.SetSize(width, height);

        var series = chart.Series.Add(valueRange, categoryRange);
        series.Header = seriesHeader;

        if (distributedColors != null && itemCount > 0)
            ApplyDistributedColors(series, distributedColors, itemCount);
        else if (seriesColor != default)
            SetSeriesColor(series, seriesColor);

        chart.YAxis.MinValue = 50;
        chart.YAxis.MaxValue = 100;
        chart.DataLabel.ShowValue = true;
        chart.Legend.Remove();
    }

    private static void ApplyDistributedColors(ExcelChartSerie serie, string[] colors, int count)
    {
        if (serie is not ExcelBarChartSerie barSerie) return;
        for (int i = 0; i < count; i++)
        {
            var dp = barSerie.DataPoints.Add(i);
            dp.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            dp.Fill.SolidFill.Color.SetRgbColor(ColorTranslator.FromHtml(colors[i % colors.Length]));
        }
    }

    private static void SetSeriesColor(ExcelChartSerie serie, Color color)
    {
        if (serie is ExcelBarChartSerie s)
        {
            s.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            s.Fill.SolidFill.Color.SetRgbColor(color);
        }
    }

    private static void WriteTableHeader(ExcelWorksheet ws, int row, int startCol, string[] headers)
    {
        for (int i = 0; i < headers.Length; i++)
            ws.Cells[row, startCol + i].Value = headers[i];
        StyleHeader(ws.Cells[row, startCol, row, startCol + headers.Length - 1]);
    }

    private static void StyleHeader(ExcelRange range, Color? bg = null)
    {
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(bg ?? ColHeader);
        range.Style.Font.Color.SetColor(Color.White);
        range.Style.Font.Bold = true;
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }

    private static void StyleTitle(ExcelRange cell)
    {
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(ColNavy);
        cell.Style.Font.Color.SetColor(Color.White);
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 16;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    private static void StyleSectionHeader(ExcelRange cell)
    {
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(ColTeal);
        cell.Style.Font.Color.SetColor(Color.White);
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 12;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }
}
