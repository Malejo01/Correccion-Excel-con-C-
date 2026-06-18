using System.Globalization;
using System.Text;

namespace CorreccionExcel.Core.Validators;

/// <summary>
/// Valida y sanitiza la interpretación textual del coeficiente de Pearson.
/// Realiza limpieza de espacios, conversión a minúsculas, remoción de tildes,
/// y verifica coherencia lógica del texto contra el valor numérico de Pearson.
/// </summary>
public static class PearsonInterpretationValidator
{
    /// <summary>
    /// Diccionario de palabras clave que indican fuerza de correlación.
    /// </summary>
    private static readonly Dictionary<string, double> PalabrasFuerza = new(StringComparer.OrdinalIgnoreCase)
    {
        { "muy_fuerte", 0.9 },
        { "muyfuerte", 0.9 },
        { "muy fuerte", 0.9 },
        { "fuerte", 0.7 },
        { "moderada", 0.5 },
        { "moderado", 0.5 },
        { "media", 0.5 },
        { "medio", 0.5 },
        { "debil", 0.3 },
        { "débil", 0.3 },
        { "muy_debil", 0.05 },
        { "muydebil", 0.05 },
        { "muy débil", 0.05 },
        { "nula", 0.0 },
        { "ninguna", 0.0 },
        { "alta", 0.7 },
        { "baja", 0.3 },
        { "bajo", 0.3 },
    };

    /// <summary>
    /// Palabras clave que indican dirección (signo) de la correlación.
    /// </summary>
    private static readonly string[] PalabrasPositivas = ["positiva", "positivo", "directa", "creciente"];
    private static readonly string[] PalabrasNegativas = ["negativa", "negativo", "inversa", "decreciente"];

    /// <summary>
    /// Valida si el texto del alumno es una interpretación coherente del valor de Pearson.
    /// </summary>
    /// <param name="textoAlumno">Texto escrito por el alumno (ej. "Correlación fuerte positiva").</param>
    /// <param name="valorPearson">Coeficiente de Pearson calculado por el alumno.</param>
    /// <returns>true si el texto es coherente con el valor; false en caso contrario.</returns>
    public static bool IsValidInterpretation(string? textoAlumno, double valorPearson)
    {
        if (string.IsNullOrWhiteSpace(textoAlumno))
            return false;

        string texto = SanitizarTexto(textoAlumno);

        // Validar coherencia de signo (positivo/negativo)
        if (!ValidarSigno(texto, valorPearson))
            return false;

        // Validar coherencia de fuerza (débil, moderada, fuerte, etc.)
        if (!ValidarFuerza(texto, valorPearson))
            return false;

        return true;
    }

    /// <summary>
    /// Limpia el texto: espacios extra, minúsculas, remoción de tildes.
    /// </summary>
    private static string SanitizarTexto(string texto)
    {
        // Pasar a minúsculas
        texto = texto.ToLowerInvariant();

        // Remover tildes/acentos
        texto = RemoverAcentos(texto);

        // Normalizar espacios (múltiples espacios a uno)
        texto = System.Text.RegularExpressions.Regex.Replace(texto, @"\s+", " ").Trim();

        return texto;
    }

    /// <summary>
    /// Remueve acentos y tildes del texto.
    /// </summary>
    private static string RemoverAcentos(string texto)
    {
        var textosNormalizado = texto.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();

        foreach (char c in textosNormalizado)
        {
            var cat = CharUnicodeInfo.GetUnicodeCategory(c);
            if (cat != UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Valida que el signo (positivo/negativo) sea coherente con el valor de Pearson.
    /// </summary>
    private static bool ValidarSigno(string texto, double valorPearson)
    {
        bool esPositivoAlumno = PalabrasPositivas.Any(p => texto.Contains(p, StringComparison.OrdinalIgnoreCase));
        bool esNegativoAlumno = PalabrasNegativas.Any(p => texto.Contains(p, StringComparison.OrdinalIgnoreCase));

        bool esPositivoEsperado = valorPearson >= 0;

        // Si el alumno menciona signo explícitamente, debe coincidir
        if (esPositivoAlumno && !esPositivoEsperado)
            return false;

        if (esNegativoAlumno && esPositivoEsperado)
            return false;

        return true;
    }

    /// <summary>
    /// Valida que la fuerza descrita sea coherente con la magnitud de Pearson.
    /// </summary>
    private static bool ValidarFuerza(string texto, double valorPearson)
    {
        double absValor = Math.Abs(valorPearson);

        // Detectar todas las palabras de fuerza mencionadas en el texto
        var fuerza = DetectarFuerza(texto);

        if (fuerza.Count == 0)
        {
            // Si no menciona fuerza explícitamente, puede ser válido
            // (ej. solo menciona "positiva" sin adjetivo de fuerza)
            return true;
        }

        // Verificar que al menos una palabra de fuerza sea coherente
        foreach (var (palabra, umbral) in fuerza)
        {
            if (EsFuerzaCoherente(absValor, umbral))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Detecta todas las palabras clave de fuerza presentes en el texto.
    /// </summary>
    private static List<(string palabra, double umbral)> DetectarFuerza(string texto)
    {
        var resultado = new List<(string, double)>();

        foreach (var (palabra, umbral) in PalabrasFuerza)
        {
            if (texto.Contains(palabra, StringComparison.OrdinalIgnoreCase))
                resultado.Add((palabra, umbral));
        }

        return resultado;
    }

    /// <summary>
    /// Valida si la fuerza descrita es coherente con el valor absoluto de Pearson.
    /// Permite un margen de 0.1 entre el umbral esperado y el valor actual.
    /// </summary>
    private static bool EsFuerzaCoherente(double absValor, double umbral)
    {
        const double margenFuerza = 0.1;
        double diferencia = Math.Abs(absValor - umbral);
        return diferencia <= margenFuerza;
    }

    /// <summary>
    /// Genera un texto de interpretación esperado para un valor de Pearson.
    /// Útil para validación y reportes.
    /// </summary>
    public static string GenerarInterpretacionEsperada(double valorPearson)
    {
        double abs = Math.Abs(valorPearson);
        string signo = valorPearson >= 0 ? "positiva" : "negativa";
        string fuerza = abs switch
        {
            >= 0.9 => "muy fuerte",
            >= 0.7 => "fuerte",
            >= 0.5 => "moderada",
            >= 0.3 => "debil",
            _ => "muy debil o nula"
        };

        return $"{fuerza} {signo}";
    }
}
