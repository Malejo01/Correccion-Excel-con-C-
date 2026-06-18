namespace CorreccionExcel.Core.Models;

/// <summary>
/// Almacena los valores exactos calculados por el profesor para un tema específico.
/// </summary>
public sealed class MasterMetrics
{
    public string Tema { get; init; } = string.Empty;
    public double PromedioX { get; init; }
    public double PromedioY { get; init; }
    public double Covarianza { get; init; }
    public double DesvioEstandarX { get; init; }
    public double DesvioEstandarY { get; init; }
    public double CorrelacionPearson { get; init; }
    public double PendienteRecta { get; init; }
    public double OrdenadaOrigen { get; init; }
    public int CantidadDatos { get; init; }
    public double[] DatosX { get; init; } = [];
    public double[] DatosY { get; init; } = [];
}
