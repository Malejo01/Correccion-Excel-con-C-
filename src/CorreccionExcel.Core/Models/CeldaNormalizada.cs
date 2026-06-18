namespace CorreccionExcel.Core.Models;

/// <summary>
/// Representa una celda normalizada extraída del archivo Excel del alumno.
/// </summary>
public sealed class CeldaNormalizada
{
    public int Fila { get; init; }
    public int Columna { get; init; }
    public string Direccion { get; init; } = string.Empty;
    public double? ValorNumerico { get; init; }
    public string? Formula { get; init; }
    public string? TextoValor { get; init; }
    public bool EsNumerica => ValorNumerico.HasValue;
    public bool TieneFormula => !string.IsNullOrWhiteSpace(Formula);
}
