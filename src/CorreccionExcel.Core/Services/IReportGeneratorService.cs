using CorreccionExcel.Core.Models;

namespace CorreccionExcel.Core.Services;

/// <summary>
/// Genera el archivo Excel consolidado de notas y los logs de trazabilidad.
/// </summary>
public interface IReportGeneratorService
{
    /// <summary>
    /// Genera el archivo Consolidado_Notas.xlsx en el directorio de salida indicado.
    /// </summary>
    /// <param name="evaluaciones">Lista de resultados de todos los alumnos procesados.</param>
    /// <param name="directorioSalida">Directorio donde se escribirán los archivos de salida.</param>
    void GenerarConsolidado(IReadOnlyList<StudentEvaluation> evaluaciones, string directorioSalida);

    /// <summary>
    /// Genera el log detallado en formato JSON y el resumen en TXT.
    /// </summary>
    void GenerarLogs(IReadOnlyList<StudentEvaluation> evaluaciones, string directorioSalida);
}
