namespace CorreccionExcel.Core.Models;

/// <summary>
/// Posibles resultados de comparación de una métrica individual.
/// </summary>
public enum MetricMatchType
{
    /// <summary>Coincide con el master y tiene fórmula legítima. 100% del peso.</summary>
    MasterMatch,
    /// <summary>Coincide con la recalculación por arrastre y tiene fórmula legítima. 66% del peso.</summary>
    CarryOverMatch,
    /// <summary>El valor coincide pero la celda está hardcodeada (sin fórmula). 0 puntos.</summary>
    Hardcoded,
    /// <summary>No coincide con ningún cálculo esperado. 0 puntos.</summary>
    NoMatch,
    /// <summary>No se encontró celda candidata en el archivo del alumno.</summary>
    NotFound
}

/// <summary>
/// Resultado detallado de la comparación de una métrica específica.
/// </summary>
public sealed class MetricCheckResult
{
    public string MetricName { get; init; } = string.Empty;
    public double ValorEsperado { get; init; }
    public double? ValorRecalculado { get; init; }
    public double? ValorAlumno { get; init; }
    public MetricMatchType TipoMatch { get; init; }
    public bool TieneFormula { get; init; }
    public string Formula { get; init; } = string.Empty;
    public string CeldaDireccion { get; init; } = string.Empty;
    public double PesoBase { get; init; }
    public double PuntajeObtenido { get; init; }
    public string Observacion { get; init; } = string.Empty;
}
