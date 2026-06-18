namespace CorreccionExcel.Core.Validators;

/// <summary>
/// Implementa la validación de coincidencia numérica con tolerancia relativa + epsilon.
/// La tolerancia relativa maneja escalas variables (valores pequeños como 0.05 hasta grandes como 100+).
/// El epsilon base (0.005) asegura que diferencias muy pequeñas siempre se acepten.
/// </summary>
public static class ToleranceValidator
{
    /// <summary>
    /// Epsilon base: margen mínimo aceptable incluso para valores pequeños.
    /// </summary>
    private const double EpsilonBase = 0.005;

    /// <summary>
    /// Porcentaje de tolerancia relativa: 5% del valor esperado.
    /// Ej: si esperado=100, tolera ±5.
    /// Si esperado=0.5, tolera ±0.025 (pero el epsilon de 0.005 asegura un mínimo).
    /// </summary>
    private const double PorcentajeTolerancia = 0.05;

    /// <summary>
    /// Valida si el valor del alumno coincide con el valor esperado,
    /// usando tolerancia relativa (5%) + epsilon base.
    /// </summary>
    /// <param name="valorAlumno">Valor calculado por el alumno.</param>
    /// <param name="valorEsperado">Valor oficial del master.</param>
    /// <returns>true si la diferencia está dentro de la tolerancia permitida.</returns>
    public static bool IsWithinTolerance(double valorAlumno, double valorEsperado, double? epsilonBase = null, double? porcentajeTolerancia = null)
    {
        double diferencia = Math.Abs(valorAlumno - valorEsperado);

        double epsilon = epsilonBase ?? EpsilonBase;
        double porcentaje = porcentajeTolerancia ?? PorcentajeTolerancia;
        
        // Tolerancia relativa: 5% del valor esperado
        double toleranciaRelativa = Math.Abs(valorEsperado) * porcentaje;
        
        // Tolerancia final: máximo entre relativa y epsilon base
        double toleranciaFinal = Math.Max(toleranciaRelativa, epsilon);

        return diferencia <= toleranciaFinal;
    }

    /// <summary>
    /// Calcula el margen de error como porcentaje para reportes.
    /// </summary>
    public static double CalcularPorcentajeError(double valorAlumno, double valorEsperado)
    {
        if (valorEsperado == 0)
            return valorAlumno == 0 ? 0 : double.MaxValue;

        return Math.Abs((valorAlumno - valorEsperado) / valorEsperado) * 100.0;
    }

    /// <summary>
    /// Retorna el margen de tolerancia permitido para un valor esperado.
    /// Útil para debugging y trazabilidad.
    /// </summary>
    public static (double minimo, double maximo) GetToleranceRange(double valorEsperado, double? epsilonBase = null, double? porcentajeTolerancia = null)
    {
        double epsilon = epsilonBase ?? EpsilonBase;
        double porcentaje = porcentajeTolerancia ?? PorcentajeTolerancia;
        double tolerancia = Math.Max(Math.Abs(valorEsperado) * porcentaje, epsilon);
        return (valorEsperado - tolerancia, valorEsperado + tolerancia);
    }
}
