namespace CorreccionExcel.Core.Models;

/// <summary>
/// Almacena el estado completo de corrección de un alumno para un tema.
/// </summary>
public sealed class StudentEvaluation
{
    public string NombreAlumno { get; init; } = string.Empty;
    public string TemaAsignado { get; init; } = string.Empty;
    public string RutaArchivo { get; init; } = string.Empty;
    public string HojaDetectada { get; init; } = string.Empty;

    public double NotaFinal { get; set; }
    public bool Aprobado => NotaFinal >= 60.0;

    public List<MetricCheckResult> ResultadosMetricas { get; init; } = [];
    public List<string> ErroresDetectados { get; init; } = [];

    // Columnas intermedias (bonus)
    public bool TieneColumnasIntermedias { get; set; }
    public int ColumnasIntermediasEncontradas { get; set; }
    public double PuntajeColumnasIntermedias { get; set; }

    // Resumen de la evaluación
    public int TotalMetricasCorrectas => ResultadosMetricas.Count(r => r.TipoMatch == MetricMatchType.MasterMatch);
    public int TotalMetricasArrastre => ResultadosMetricas.Count(r => r.TipoMatch == MetricMatchType.CarryOverMatch);
    public int TotalMetricasHardcodeadas => ResultadosMetricas.Count(r => r.TipoMatch == MetricMatchType.Hardcoded);
    public int TotalMetricasSinMatch => ResultadosMetricas.Count(r => r.TipoMatch is MetricMatchType.NoMatch or MetricMatchType.NotFound);
    public bool HuboErrorProcesamiento { get; set; }
}
