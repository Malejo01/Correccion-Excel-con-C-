namespace CorreccionExcel.Core.Models;

/// <summary>
/// Contexto de evaluación que mantiene el estado del alumno durante el proceso de corrección.
/// Almacena los valores del alumno para aplicar la lógica de "efecto arrastre".
/// </summary>
public sealed class EvaluationContext
{
    public string NombreAlumno { get; init; } = string.Empty;
    public string Tema { get; init; } = string.Empty;
    public MasterMetrics Master { get; init; } = new();

    // Valores del alumno para métricas base (usados en recálculo de arrastre)
    public double? AlumnoPromedioX { get; set; }
    public double? AlumnoPromedioY { get; set; }
    public double? AlumnoCovarianza { get; set; }
    public double? AlumnoDesvioX { get; set; }
    public double? AlumnoDesvioY { get; set; }
    public double? AlumnoPearson { get; set; }

    // Indica si las métricas base matchearon con el master (o arrastraron error)
    public bool PromedioXEsCorrecto { get; set; }
    public bool PromedioYEsCorrecto { get; set; }
}
