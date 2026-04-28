using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;

namespace FoodSafetyDashboard.Services;

public class ExportService(AuditService auditService)
{
    private static readonly string[] SectionCols =
        ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];

    // Mismas paletas que la web (Home.razor)
    private static readonly string[] GroupColors =
        ["#2E75B6", "#ED7D31", "#70AD47", "#FFC000", "#7030A0", "#FF4136"];

    private static readonly string[] SectionColors =
        ["#E53935", "#039BE5", "#FB8C00", "#8E24AA", "#43A047"];

    static ExportService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<byte[]> ExportStoreDetailAsync(int? year = null, int? month = null)
    {
        var rows   = await auditService.GetStoreRowsAsync(year, month);
        var groups = await auditService.GetGroupSummariesAsync(year, month);
        var kpis   = await auditService.GetKpiSummaryAsync(year, month);

        using var pkg = new ExcelPackage();

        AddDetalleSheet(pkg, rows);
        AddResumenSheet(pkg, groups);
        AddDashboardSheet(pkg, groups, kpis, rows);

        return await pkg.GetAsByteArrayAsync();
    }

    // ── Hoja 1: Detalle por Tienda ────────────────────────────────────────────

    private static void AddDetalleSheet(ExcelPackage pkg, List<StoreRow> rows)
    {
        var ws = pkg.Workbook.Worksheets.Add("Detalle por Tienda");

        string[] headers = ["Tienda", "Grupo", "Fecha", "Nota %", "Resultado", "Críticos",
            "Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento",
            "# Fallos", "Puntos"];

        for (int c = 0; c < headers.Length; c++)
            ws.Cells[1, c + 1].Value = headers[c];
        StyleHeader(ws.Cells[1, 1, 1, headers.Length]);

        int r = 2;
        foreach (var row in rows)
        {
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

            var bgColor = row.Result == "Aprobado"
                ? Color.FromArgb(198, 239, 206)
                : Color.FromArgb(255, 199, 206);
            ws.Cells[r, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[r, 5].Style.Fill.BackgroundColor.SetColor(bgColor);

            r++;
        }

        if (ws.Dimension != null)
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
    }

    // ── Hoja 2: Resumen por Grupo ─────────────────────────────────────────────

    private static void AddResumenSheet(ExcelPackage pkg, List<GroupSummary> groups)
    {
        var ws = pkg.Workbook.Worksheets.Add("Resumen por Grupo");

        string[] headers = ["Grupo", "Tiendas", "Nota %", "Fact. Riesgo", "Limpieza",
            "Mantenimiento", "Almacenaje", "Conocimiento", "Aprobadas", "No Aprobadas"];

        for (int c = 0; c < headers.Length; c++)
            ws.Cells[1, c + 1].Value = headers[c];
        StyleHeader(ws.Cells[1, 1, 1, headers.Length]);

        int r = 2;
        foreach (var g in groups)
        {
            ws.Cells[r, 1].Value = g.Group;
            ws.Cells[r, 2].Value = g.Count;
            ws.Cells[r, 3].Value = g.AvgScore / 100.0;
            ws.Cells[r, 3].Style.Numberformat.Format = "0.0%";

            for (int i = 0; i < SectionCols.Length; i++)
            {
                ws.Cells[r, 4 + i].Value = g.SectionAverages.GetValueOrDefault(SectionCols[i]) / 100.0;
                ws.Cells[r, 4 + i].Style.Numberformat.Format = "0.0%";
            }

            ws.Cells[r, 9].Value = g.Approved;
            ws.Cells[r, 10].Value = g.NotApproved;
            r++;
        }

        if (ws.Dimension != null)
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
    }

    // ── Hoja 3: Dashboard con Gráficas ────────────────────────────────────────

    private void AddDashboardSheet(
        ExcelPackage pkg,
        List<GroupSummary> groups,
        KpiSummary kpis,
        List<StoreRow> rows)
    {
        var ws = pkg.Workbook.Worksheets.Add("Dashboard");

        // Datos y gráfica 1 — Nota por Grupo (col A-B, chart col D)
        int g1Row = 1;
        WriteTableHeader(ws, g1Row, 1, ["Grupo", "Nota %"]);
        for (int i = 0; i < groups.Count; i++)
        {
            ws.Cells[g1Row + 1 + i, 1].Value = groups[i].Group;
            ws.Cells[g1Row + 1 + i, 2].Value = Math.Round(groups[i].AvgScore, 1);
        }
        int g1End = g1Row + groups.Count;
        AddColumnChart(ws, "chart_grupos", "Nota Global por Grupo",
            ws.Cells[g1Row + 1, 2, g1End, 2],
            ws.Cells[g1Row + 1, 1, g1End, 1],
            "Nota %",
            row: g1Row - 1, col: 3,
            width: 500, height: 280,
            distributedColors: GroupColors, itemCount: groups.Count);

        // Datos y gráfica 2 — Promedio por Sección (col A-B, chart col D)
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
            ws.Cells[g2Row + 1, 2, g2End, 2],
            ws.Cells[g2Row + 1, 1, g2End, 1],
            "Score %",
            row: g2Row - 1, col: 3,
            width: 540, height: 280,
            distributedColors: SectionColors, itemCount: sections.Count);

        // Datos y gráfica 3 — CLÁSICA vs DEX por Sección
        var clasica = groups.FirstOrDefault(g => g.Group == "CLÁSICA");
        var dex     = groups.FirstOrDefault(g => g.Group == "DEX");
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

            var s3a = chart3.Series.Add(
                ws.Cells[g3Row + 1, 2, g3End, 2],
                ws.Cells[g3Row + 1, 1, g3End, 1]);
            s3a.Header = "CLÁSICA";
            SetSeriesColor(s3a, Color.FromArgb(46, 117, 182));

            var s3b = chart3.Series.Add(
                ws.Cells[g3Row + 1, 3, g3End, 3],
                ws.Cells[g3Row + 1, 1, g3End, 1]);
            s3b.Header = "DEX";
            SetSeriesColor(s3b, Color.FromArgb(237, 125, 49));

            chart3.YAxis.MinValue = 50;
            chart3.YAxis.MaxValue = 100;
            chart3.DataLabel.ShowValue = true;
        }

        // Datos y gráfica 4 — Ranking de Tiendas (barra horizontal)
        var sorted = rows.OrderByDescending(r => r.Score).ToList();
        int g4Row = g3End + 3;
        WriteTableHeader(ws, g4Row, 1, ["Tienda", "Nota %"]);
        for (int i = 0; i < sorted.Count; i++)
        {
            ws.Cells[g4Row + 1 + i, 1].Value = sorted[i].StoreId.ToString();
            ws.Cells[g4Row + 1 + i, 2].Value = Math.Round(sorted[i].Score, 1);
        }
        int g4End = g4Row + sorted.Count;
        int rankH  = Math.Max(300, sorted.Count * 22 + 80);

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
            SetSeriesColor(s4, Color.FromArgb(46, 117, 182));

            chart4.DataLabel.ShowValue = true;
            chart4.Legend.Remove();
        }

        ws.Column(1).AutoFit();
        ws.Column(2).AutoFit();
        ws.Column(3).AutoFit();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static void AddColumnChart(
        ExcelWorksheet ws,
        string name,
        string title,
        ExcelRange valueRange,
        ExcelRange categoryRange,
        string seriesHeader,
        int row, int col,
        int width, int height,
        Color seriesColor = default,
        string[]? distributedColors = null,
        int itemCount = 0)
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
            dp.Fill.SolidFill.Color.SetRgbColor(
                ColorTranslator.FromHtml(colors[i % colors.Length]));
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

    private static void StyleHeader(ExcelRange range)
    {
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(46, 117, 182));
        range.Style.Font.Color.SetColor(Color.White);
        range.Style.Font.Bold = true;
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }
}
