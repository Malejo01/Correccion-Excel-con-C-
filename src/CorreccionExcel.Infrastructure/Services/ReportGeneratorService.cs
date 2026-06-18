using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using CorreccionExcel.Core.Models;
using CorreccionExcel.Core.Services;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace CorreccionExcel.Infrastructure.Services;

/// <summary>
/// Genera el archivo Excel consolidado de notas y los logs de trazabilidad (JSON + TXT).
/// </summary>
public sealed class ReportGeneratorService : IReportGeneratorService
{
    private readonly ILogger<ReportGeneratorService> _logger;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public ReportGeneratorService(ILogger<ReportGeneratorService> logger)
    {
        _logger = logger;
    }

    // ─── Excel consolidado ────────────────────────────────────────────────────

    public void GenerarConsolidado(IReadOnlyList<StudentEvaluation> evaluaciones, string directorioSalida)
    {
        Directory.CreateDirectory(directorioSalida);
        var rutaSalida = Path.Combine(directorioSalida, "Consolidado_Notas.xlsx");

        using var paquete = new ExcelPackage();
        var hoja = paquete.Workbook.Worksheets.Add("Consolidado");

        EscribirEncabezados(hoja);
        EscribirFilasAlumnos(hoja, evaluaciones);
        AplicarFormato(hoja, evaluaciones.Count);

        try
        {
            paquete.SaveAs(new FileInfo(rutaSalida));
            _logger.LogInformation("Consolidado generado en: {Ruta}", rutaSalida);
        }
        catch (Exception)
        {
            _logger.LogWarning("¡ATENCIÓN! No se pudo sobrescribir 'Consolidado_Notas.xlsx' porque está bloqueado o abierto en otro programa.");
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var rutaAlternativa = Path.Combine(directorioSalida, $"Consolidado_Notas_{timestamp}.xlsx");
            try
            {
                paquete.SaveAs(new FileInfo(rutaAlternativa));
                _logger.LogInformation("Consolidado guardado con éxito en ruta alternativa: {Ruta}", rutaAlternativa);
            }
            catch (Exception exAlt)
            {
                _logger.LogError(exAlt, "Error crítico: tampoco se pudo escribir en la ruta alternativa.");
                throw;
            }
        }
    }

    private static void EscribirEncabezados(ExcelWorksheet hoja)
    {
        string[] columnas =
        [
            "N°", "Alumno", "Tema", "Nota Final", "Estado",
            "Col. Intermedias (/20)", "PromedioX (/5)", "PromedioY (/5)",
            "Covarianza (/15)", "Desvío X (/7.5)", "Desvío Y (/7.5)",
            "Pearson (/15)", "Interpretación (/10)", "Pendiente (/7.5)", "Ordenada (/7.5)",
            "Hardcodeados", "Arrastres", "Errores"
        ];

        for (int i = 0; i < columnas.Length; i++)
        {
            hoja.Cells[1, i + 1].Value = columnas[i];
        }

        // Estilo de encabezado
        using var rango = hoja.Cells[1, 1, 1, columnas.Length];
        rango.Style.Font.Bold = true;
        rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
        rango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 73, 125));
        rango.Style.Font.Color.SetColor(Color.White);
        rango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }

    private static void EscribirFilasAlumnos(ExcelWorksheet hoja, IReadOnlyList<StudentEvaluation> evaluaciones)
    {
        // Orden: por tema, luego por nota descendente
        var ordenadas = evaluaciones
            .OrderBy(e => e.TemaAsignado)
            .ThenByDescending(e => e.NotaFinal)
            .ToList();

        for (int i = 0; i < ordenadas.Count; i++)
        {
            var ev = ordenadas[i];
            int fila = i + 2;

            double ObtenerPuntaje(string metrica) =>
                ev.ResultadosMetricas.FirstOrDefault(r => r.MetricName == metrica)?.PuntajeObtenido ?? 0;

            hoja.Cells[fila, 1].Value = i + 1;
            hoja.Cells[fila, 2].Value = ev.NombreAlumno;
            hoja.Cells[fila, 3].Value = $"Tema {ev.TemaAsignado}";
            hoja.Cells[fila, 4].Value = Math.Round(ev.NotaFinal, 2);
            hoja.Cells[fila, 5].Value = ev.Aprobado ? "APROBADO" : "REPROBADO";
            hoja.Cells[fila, 6].Value = Math.Round(ev.PuntajeColumnasIntermedias, 2);
            hoja.Cells[fila, 7].Value = ObtenerPuntaje("PromedioX");
            hoja.Cells[fila, 8].Value = ObtenerPuntaje("PromedioY");
            hoja.Cells[fila, 9].Value = ObtenerPuntaje("Covarianza");
            hoja.Cells[fila, 10].Value = ObtenerPuntaje("DesvioX");
            hoja.Cells[fila, 11].Value = ObtenerPuntaje("DesvioY");
            hoja.Cells[fila, 12].Value = ObtenerPuntaje("CorrelacionPearson");
            hoja.Cells[fila, 13].Value = ObtenerPuntaje("InterpretacionPearson");
            hoja.Cells[fila, 14].Value = ObtenerPuntaje("PendienteRecta");
            hoja.Cells[fila, 15].Value = ObtenerPuntaje("OrdenadaOrigen");
            hoja.Cells[fila, 16].Value = ev.TotalMetricasHardcodeadas;
            hoja.Cells[fila, 17].Value = ev.TotalMetricasArrastre;
            hoja.Cells[fila, 18].Value = string.Join(" | ", ev.ErroresDetectados.Take(3));

            // Resaltar reprobados en rojo
            if (!ev.Aprobado)
            {
                using var rangFila = hoja.Cells[fila, 1, fila, 18];
                rangFila.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rangFila.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                rangFila.Style.Font.Color.SetColor(Color.FromArgb(156, 0, 6));
            }
            else
            {
                using var rangFila = hoja.Cells[fila, 1, fila, 18];
                rangFila.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rangFila.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                rangFila.Style.Font.Color.SetColor(Color.FromArgb(0, 97, 0));
            }

            // Columna nota en negrita
            hoja.Cells[fila, 4].Style.Font.Bold = true;
        }
    }

    private static void AplicarFormato(ExcelWorksheet hoja, int cantidadAlumnos)
    {
        // Bordes en tabla completa
        if (cantidadAlumnos > 0)
        {
            using var tabla = hoja.Cells[1, 1, cantidadAlumnos + 1, 18];
            tabla.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            tabla.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            tabla.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            tabla.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        // Autofit de columnas
        hoja.Cells.AutoFitColumns(8, 40);
    }

    // ─── Logs JSON + TXT ──────────────────────────────────────────────────────

    public void GenerarLogs(IReadOnlyList<StudentEvaluation> evaluaciones, string directorioSalida)
    {
        Directory.CreateDirectory(directorioSalida);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        GenerarLogJson(evaluaciones, directorioSalida, timestamp);
        GenerarLogTxt(evaluaciones, directorioSalida, timestamp);
    }

    private void GenerarLogJson(IReadOnlyList<StudentEvaluation> evaluaciones, string directorio, string timestamp)
    {
        var rutaJson = Path.Combine(directorio, $"log_detallado_{timestamp}.json");
        var json = JsonSerializer.Serialize(evaluaciones, JsonOpts);
        File.WriteAllText(rutaJson, json, Encoding.UTF8);
        _logger.LogInformation("Log JSON generado: {Ruta}", rutaJson);
    }

    private void GenerarLogTxt(IReadOnlyList<StudentEvaluation> evaluaciones, string directorio, string timestamp)
    {
        var rutaTxt = Path.Combine(directorio, $"resumen_{timestamp}.txt");
        var sb = new StringBuilder();

        sb.AppendLine("═══════════════════════════════════════════════════════════════");
        sb.AppendLine($"  CORRECCIÓN AUTOMÁTICA - Probabilidad y Estadística");
        sb.AppendLine($"  Fecha: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
        sb.AppendLine("═══════════════════════════════════════════════════════════════");
        sb.AppendLine();

        // Resumen global
        int aprobados = evaluaciones.Count(e => e.Aprobado);
        int reprobados = evaluaciones.Count - aprobados;
        double promedio = evaluaciones.Count > 0 ? evaluaciones.Average(e => e.NotaFinal) : 0;

        sb.AppendLine($"Total alumnos procesados : {evaluaciones.Count}");
        sb.AppendLine($"Aprobados                : {aprobados}");
        sb.AppendLine($"Reprobados               : {reprobados}");
        sb.AppendLine($"Nota promedio            : {promedio:F2}");
        sb.AppendLine();

        // Detalle por tema
        foreach (var tema in evaluaciones.GroupBy(e => e.TemaAsignado).OrderBy(g => g.Key))
        {
            sb.AppendLine($"───── Tema {tema.Key} ({tema.Count()} alumnos) ─────────────────────────────");
            foreach (var ev in tema.OrderByDescending(e => e.NotaFinal))
            {
                string estado = ev.Aprobado ? "✓" : "✗";
                sb.AppendLine($"  {estado} {ev.NombreAlumno,-35} Nota: {ev.NotaFinal,6:F2}  " +
                              $"[HC:{ev.TotalMetricasHardcodeadas} AR:{ev.TotalMetricasArrastre}]");

                foreach (var metrica in ev.ResultadosMetricas.Where(r => r.TipoMatch != MetricMatchType.MasterMatch))
                {
                    sb.AppendLine($"      → {metrica.MetricName}: {metrica.Observacion}");
                }
            }
            sb.AppendLine();
        }

        File.WriteAllText(rutaTxt, sb.ToString(), Encoding.UTF8);
        _logger.LogInformation("Resumen TXT generado: {Ruta}", rutaTxt);
    }
}
