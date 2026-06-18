using System.Collections.Generic;
using CorreccionExcel.Core.Models;

namespace CorreccionExcel.Core.Services;

/// <summary>
/// Encapsula toda la interacción con EPPlus para extraer celdas de archivos Excel.
/// </summary>
public interface IExcelReaderService
{
    /// <summary>
    /// Lee el archivo Excel del alumno y devuelve la colección de celdas de la hoja con mayor densidad numérica.
    /// </summary>
    /// <param name="rutaArchivo">Ruta absoluta al archivo .xlsx del alumno.</param>
    /// <returns>Colección de celdas normalizadas de la hoja objetivo.</returns>
    IReadOnlyList<CeldaNormalizada> LeerHojaAlumno(string rutaArchivo);

    /// <summary>
    /// Devuelve el nombre de la hoja que fue seleccionada como objetivo en la última llamada a LeerHojaAlumno.
    /// </summary>
    string UltimaHojaSeleccionada { get; }

    /// <summary>
    /// Devuelve la serie de datos X extraída del alumno en la última lectura.
    /// </summary>
    double[] UltimosDatosX { get; }

    /// <summary>
    /// Devuelve la serie de datos Y extraída del alumno en la última lectura.
    /// </summary>
    double[] UltimosDatosY { get; }
}
