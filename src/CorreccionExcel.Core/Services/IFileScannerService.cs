using CorreccionExcel.Core.Models;

namespace CorreccionExcel.Core.Services;

/// <summary>
/// Escanea el directorio raíz para descubrir archivos master y archivos de alumnos por tema.
/// </summary>
public interface IFileScannerService
{
    /// <summary>
    /// Escanea el directorio raíz y devuelve la información de masters y alumnos organizados por tema.
    /// </summary>
    /// <param name="directorioRaiz">Ruta absoluta al directorio que contiene las carpetas de tema y los masters.</param>
    ScanResult EscanearDirectorio(string directorioRaiz);
}
