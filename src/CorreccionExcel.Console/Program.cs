using System.Collections.Generic;
using CorreccionExcel.Application.Services;
using CorreccionExcel.Core.Configuration;
using CorreccionExcel.Core.Data;
using CorreccionExcel.Core.Models;
using CorreccionExcel.Core.Services;
using CorreccionExcel.Infrastructure.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

// ─── Licencia EPPlus (obligatorio para uso no comercial) ─────────────────────
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// ─── Bootstrap Generic Host + configuración tipada ───────────────────────────
var builder = Host.CreateApplicationBuilder(args);

var evaluationOptions = builder.Configuration
    .GetSection("Evaluation")
    .Get<EvaluationOptions>() ?? new EvaluationOptions();

var scannerOptions = builder.Configuration
    .GetSection("Scanner")
    .Get<ScannerOptions>() ?? new ScannerOptions();

builder.Services.AddSingleton(evaluationOptions);
builder.Services.AddSingleton(scannerOptions);

builder.Services.AddScoped<IFileScannerService, FileScannerService>();
builder.Services.AddScoped<IExcelReaderService, ExcelReaderService>();
builder.Services.AddScoped<IEvaluationEngine, EvaluationEngine>();
builder.Services.AddScoped<IReportGeneratorService, ReportGeneratorService>();

using var host = builder.Build();
using var scope = host.Services.CreateScope();
var proveedor = scope.ServiceProvider;
var logger = proveedor.GetRequiredService<ILogger<Program>>();

try
{
    string directorioRaiz = ObtenerDirectorioRaiz(args);
    string directorioSalida = Path.Combine(directorioRaiz, "Resultados");

    logger.LogInformation("═══════════════════════════════════════════════════");
    logger.LogInformation("  Corrección automática de Probabilidad y Estadística");
    logger.LogInformation("  Directorio raíz: {Dir}", directorioRaiz);
    logger.LogInformation("═══════════════════════════════════════════════════");

    // 1. Escanear directorio raíz buscando subcarpetas de alumnos
    var scanner = proveedor.GetRequiredService<IFileScannerService>();
    var scanResult = scanner.EscanearDirectorio(directorioRaiz);

    if (scanResult.ArchivosAlumnos.Count == 0)
    {
        logger.LogError("No se encontraron archivos de alumnos en las carpetas Tema A/B/C/D.");
        return 1;
    }

    foreach (var advertencia in scanResult.Advertencias)
        logger.LogWarning("{Advertencia}", advertencia);

    logger.LogInformation("Alumnos a corregir:  {N}", scanResult.ArchivosAlumnos.Count);

    // 2. Cargar maestros desde el Repositorio de la Verdad Absoluta
    logger.LogInformation("Cargando métricas de master desde repositorio en memoria...");
    var masters = new Dictionary<string, MasterMetrics>(StringComparer.OrdinalIgnoreCase);
    foreach (var tema in MasterMetricsRepository.GetTemasDisponibles())
    {
        var metrics = MasterMetricsRepository.GetMetricsForTema(tema);
        if (metrics is not null)
        {
            masters[tema] = metrics;
            logger.LogDebug("Master Tema {Tema} cargado: Pearson={Pearson:F6}", tema, metrics.CorrelacionPearson);
        }
    }

    var lector = proveedor.GetRequiredService<IExcelReaderService>();
    var motor = proveedor.GetRequiredService<IEvaluationEngine>();

    var evaluaciones = new List<StudentEvaluation>();
    int procesados = 0;

    foreach (var alumno in scanResult.ArchivosAlumnos)
    {
        logger.LogInformation("[{N}/{Total}] Procesando: {Nombre} (Tema {Tema})",
            ++procesados, scanResult.ArchivosAlumnos.Count, alumno.NombreAlumno, alumno.Tema);

        if (!masters.TryGetValue(alumno.Tema, out var masterBase))
        {
            logger.LogWarning("Sin master en el repositorio para el Tema {Tema}. Archivo omitido: {Archivo}", alumno.Tema, alumno.RutaCompleta);
            continue;
        }

        try
        {
            // Leer la hoja del alumno
            var celdas = lector.LeerHojaAlumno(alumno.RutaCompleta);

            // Clonar o construir la métrica master con el dataset X e Y específico que usó el alumno
            // Esto permite que el motor evalúe correctamente la lógica de arrastre en caliente.
            var masterTema = new MasterMetrics
            {
                Tema = masterBase.Tema,
                PromedioX = masterBase.PromedioX,
                PromedioY = masterBase.PromedioY,
                Covarianza = masterBase.Covarianza,
                DesvioEstandarX = masterBase.DesvioEstandarX,
                DesvioEstandarY = masterBase.DesvioEstandarY,
                CorrelacionPearson = masterBase.CorrelacionPearson,
                PendienteRecta = masterBase.PendienteRecta,
                OrdenadaOrigen = masterBase.OrdenadaOrigen,
                DatosX = lector.UltimosDatosX,
                DatosY = lector.UltimosDatosY,
                CantidadDatos = Math.Min(lector.UltimosDatosX.Length, lector.UltimosDatosY.Length)
            };

            var resultado = motor.Evaluar(
                alumno.NombreAlumno,
                alumno.RutaCompleta,
                lector.UltimaHojaSeleccionada,
                celdas,
                masterTema);

            evaluaciones.Add(resultado);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error procesando: {Archivo}", alumno.RutaCompleta);
        }
    }

    if (evaluaciones.Count == 0)
    {
        logger.LogWarning("No se generó ninguna evaluación.");
        return 0;
    }

    // 3. Generar reportes
    var reporteador = proveedor.GetRequiredService<IReportGeneratorService>();
    reporteador.GenerarConsolidado(evaluaciones, directorioSalida);
    reporteador.GenerarLogs(evaluaciones, directorioSalida);

    // 4. Resumen final en consola
    int aprobados = evaluaciones.Count(e => e.Aprobado);
    logger.LogInformation("───────────────────────────────────────────────────");
    logger.LogInformation("Corrección finalizada.");
    logger.LogInformation("  Procesados: {T} | Aprobados: {A} | Reprobados: {R}",
        evaluaciones.Count, aprobados, evaluaciones.Count - aprobados);
    logger.LogInformation("  Nota promedio: {P:F2}", evaluaciones.Average(e => e.NotaFinal));
    logger.LogInformation("  Resultados en: {Dir}", directorioSalida);
    logger.LogInformation("───────────────────────────────────────────────────");

    return 0;
}
catch (Exception ex)
{
    logger.LogCritical(ex, "Error fatal en la aplicación.");
    return 2;
}

// ─── Helper: resolver directorio raíz ────────────────────────────────────────
static string ObtenerDirectorioRaiz(string[] args)
{
    if (args.Length > 0 && Directory.Exists(args[0]))
        return Path.GetFullPath(args[0]);

    Console.Write("Ingresá la ruta al directorio raíz de evaluaciones: ");
    string? input = Console.ReadLine()?.Trim().Trim('"');

    if (!string.IsNullOrWhiteSpace(input) && Directory.Exists(input))
        return Path.GetFullPath(input);

    throw new DirectoryNotFoundException(
        $"Directorio no encontrado: '{input}'. " +
        "Usá: CorreccionExcel.Console.exe \"C:\\ruta\\al\\directorio\"");
}
