using CorreccionExcel.Core.Models;

namespace CorreccionExcel.Core.Services;

/// <summary>
/// Núcleo lógico del sistema. Compara los datos del alumno contra el MasterMetrics,
/// aplica tolerancia, detecta hardcodeo y calcula el efecto arrastre.
/// </summary>
public interface IEvaluationEngine
{
    /// <summary>
    /// Evalúa la corrección de un alumno dado el conjunto de celdas de su hoja y las métricas del master.
    /// </summary>
    /// <param name="nombreAlumno">Nombre del alumno (derivado del nombre de archivo).</param>
    /// <param name="rutaArchivo">Ruta del archivo del alumno.</param>
    /// <param name="hojaDetectada">Nombre de la hoja seleccionada por el lector.</param>
    /// <param name="celdasAlumno">Celdas normalizadas de la hoja del alumno.</param>
    /// <param name="master">Métricas oficiales del profesor para el tema.</param>
    StudentEvaluation Evaluar(
        string nombreAlumno,
        string rutaArchivo,
        string hojaDetectada,
        IReadOnlyList<CeldaNormalizada> celdasAlumno,
        MasterMetrics master);
}
