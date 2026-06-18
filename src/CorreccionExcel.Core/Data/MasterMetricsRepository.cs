using CorreccionExcel.Core.Models;

namespace CorreccionExcel.Core.Data;

/// <summary>
/// Repositorio en memoria que contiene los valores exactos de métricas de master
/// validados por el profesor. Fuente única de verdad para la evaluación.
/// Elimina cualquier variabilidad en la lectura de archivos maestros.
/// </summary>
public static class MasterMetricsRepository
{
    private static readonly Dictionary<string, MasterMetrics> Metricas;

    static MasterMetricsRepository()
    {
        Metricas = new Dictionary<string, MasterMetrics>(StringComparer.OrdinalIgnoreCase)
        {
            {
                "A",
                new MasterMetrics
                {
                    Tema = "A",
                    PromedioX = 7.625,
                    PromedioY = 82.125,
                    Covarianza = 38.546875,
                    DesvioEstandarX = 3.966657913,
                    DesvioEstandarY = 9.942805188,
                    CorrelacionPearson = 0.977362107,
                    PendienteRecta = 2.449851043,
                    OrdenadaOrigen = 63.4448858
                }
            },
            {
                "B",
                new MasterMetrics
                {
                    Tema = "B",
                    PromedioX = 8.625,
                    PromedioY = 83.5,
                    Covarianza = 38.5625,
                    DesvioEstandarX = 3.966657913,
                    DesvioEstandarY = 10.03742995,
                    CorrelacionPearson = 0.968540768,
                    PendienteRecta = 2.450844091,
                    OrdenadaOrigen = 62.36146971
                }
            },
            {
                "C",
                new MasterMetrics
                {
                    Tema = "C",
                    PromedioX = 9.625,
                    PromedioY = 85.125,
                    Covarianza = 37.796875,
                    DesvioEstandarX = 3.966657913,
                    DesvioEstandarY = 9.829006817,
                    CorrelacionPearson = 0.969441283,
                    PendienteRecta = 2.402184707,
                    OrdenadaOrigen = 62.00397219
                }
            },
            {
                "D",
                new MasterMetrics
                {
                    Tema = "D",
                    PromedioX = 10.625,
                    PromedioY = 86.625,
                    Covarianza = 36.484375,
                    DesvioEstandarX = 3.966657913,
                    DesvioEstandarY = 9.512327528,
                    CorrelacionPearson = 0.966930740,
                    PendienteRecta = 2.318768620,
                    OrdenadaOrigen = 61.98808342
                }
            }
        };
    }

    /// <summary>
    /// Obtiene las métricas validadas por el profesor para un tema específico.
    /// </summary>
    /// <param name="tema">Identificador del tema: A, B, C o D</param>
    /// <returns>Métricas oficiales del tema, o null si no existe.</returns>
    public static MasterMetrics? GetMetricsForTema(string tema)
    {
        if (string.IsNullOrWhiteSpace(tema))
            return null;

        return Metricas.TryGetValue(tema.Trim(), out var metrics) ? metrics : null;
    }

    /// <summary>
    /// Verifica si un tema está registrado en el repositorio.
    /// </summary>
    public static bool TemaExists(string tema)
        => !string.IsNullOrWhiteSpace(tema) && Metricas.ContainsKey(tema.Trim());

    /// <summary>
    /// Retorna todos los temas disponibles.
    /// </summary>
    public static IReadOnlyCollection<string> GetTemasDisponibles()
        => Metricas.Keys.ToList().AsReadOnly();
}
