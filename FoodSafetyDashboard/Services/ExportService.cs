using ClosedXML.Excel;
using FoodSafetyDashboard.Services;

namespace FoodSafetyDashboard.Services;

public class ExportService(AuditService auditService)
{
    private static readonly string[] SectionCols =
        ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];

    public async Task<byte[]> ExportStoreDetailAsync(int? year = null, int? month = null)
    {
        var rows = await auditService.GetStoreRowsAsync(year, month);
        var groups = await auditService.GetGroupSummariesAsync(year, month);
        var kpis = await auditService.GetKpiSummaryAsync(year, month);

        using var wb = new XLWorkbook();

        // ── Hoja 1: Detalle por Tienda ────────────────────────────────────────
        var ws = wb.AddWorksheet("Detalle por Tienda");

        // Cabecera
        ws.Cell(1, 1).Value = "Tienda";
        ws.Cell(1, 2).Value = "Grupo";
        ws.Cell(1, 3).Value = "Fecha";
        ws.Cell(1, 4).Value = "Nota %";
        ws.Cell(1, 5).Value = "Resultado";
        ws.Cell(1, 6).Value = "Críticos";
        ws.Cell(1, 7).Value = "Fact. Riesgo";
        ws.Cell(1, 8).Value = "Limpieza";
        ws.Cell(1, 9).Value = "Mantenimiento";
        ws.Cell(1, 10).Value = "Almacenaje";
        ws.Cell(1, 11).Value = "Conocimiento";
        ws.Cell(1, 12).Value = "# Fallos";
        ws.Cell(1, 13).Value = "Puntos";

        var headerRange = ws.Range(1, 1, 1, 13);
        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#2E75B6");
        headerRange.Style.Font.FontColor = XLColor.White;
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        int dataRow = 2;
        foreach (var r in rows)
        {
            ws.Cell(dataRow, 1).Value = r.StoreId;
            ws.Cell(dataRow, 2).Value = r.Group;
            ws.Cell(dataRow, 3).Value = r.AuditDate?.ToString("yyyy-MM-dd") ?? "";
            ws.Cell(dataRow, 4).Value = r.Score / 100.0;
            ws.Cell(dataRow, 4).Style.NumberFormat.Format = "0.0%";
            ws.Cell(dataRow, 5).Value = r.Result;
            ws.Cell(dataRow, 6).Value = r.CriticalViolations;

            string[] secCols = ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];
            for (int i = 0; i < secCols.Length; i++)
            {
                var val = r.Sections.GetValueOrDefault(secCols[i]);
                if (val.HasValue)
                {
                    ws.Cell(dataRow, 7 + i).Value = val.Value / 100.0;
                    ws.Cell(dataRow, 7 + i).Style.NumberFormat.Format = "0.0%";
                }
            }

            ws.Cell(dataRow, 12).Value = r.TotalViolations;
            ws.Cell(dataRow, 13).Value = r.Points;

            // Color de resultado
            var resultColor = r.Result == "Aprobado" ? XLColor.FromHtml("#C6EFCE") : XLColor.FromHtml("#FFC7CE");
            ws.Cell(dataRow, 5).Style.Fill.BackgroundColor = resultColor;

            dataRow++;
        }

        ws.Columns().AdjustToContents();

        // ── Hoja 2: Resumen por Grupo ─────────────────────────────────────────
        var ws2 = wb.AddWorksheet("Resumen por Grupo");

        ws2.Cell(1, 1).Value = "Grupo";
        ws2.Cell(1, 2).Value = "Tiendas";
        ws2.Cell(1, 3).Value = "Nota %";
        ws2.Cell(1, 4).Value = "Fact. Riesgo";
        ws2.Cell(1, 5).Value = "Limpieza";
        ws2.Cell(1, 6).Value = "Mantenimiento";
        ws2.Cell(1, 7).Value = "Almacenaje";
        ws2.Cell(1, 8).Value = "Conocimiento";
        ws2.Cell(1, 9).Value = "Aprobadas";
        ws2.Cell(1, 10).Value = "No Aprobadas";

        ws2.Range(1, 1, 1, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#2E75B6");
        ws2.Range(1, 1, 1, 10).Style.Font.FontColor = XLColor.White;
        ws2.Range(1, 1, 1, 10).Style.Font.Bold = true;

        int gr = 2;
        foreach (var g in groups)
        {
            ws2.Cell(gr, 1).Value = g.Group;
            ws2.Cell(gr, 2).Value = g.Count;
            ws2.Cell(gr, 3).Value = g.AvgScore / 100.0;
            ws2.Cell(gr, 3).Style.NumberFormat.Format = "0.0%";

            string[] secKeys = ["Fact. Riesgo", "Limpieza", "Mantenimiento", "Almacenaje", "Conocimiento"];
            for (int i = 0; i < secKeys.Length; i++)
            {
                ws2.Cell(gr, 4 + i).Value = g.SectionAverages.GetValueOrDefault(secKeys[i]) / 100.0;
                ws2.Cell(gr, 4 + i).Style.NumberFormat.Format = "0.0%";
            }

            ws2.Cell(gr, 9).Value = g.Approved;
            ws2.Cell(gr, 10).Value = g.NotApproved;
            gr++;
        }

        ws2.Columns().AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }
}
