using System;
using System.Collections.Generic;
using System.IO;
using CorreccionExcel.Core.Configuration;
using CorreccionExcel.Core.Models;
using CorreccionExcel.Core.Services;
using Microsoft.Extensions.Logging;

namespace CorreccionExcel.Infrastructure.Services;

/// <summary>
/// Escanea el directorio raíz únicamente buscando las subcarpetas de alumnos por tema (Tema A, Tema B, Tema C, Tema D).
/// No contiene ninguna lógica de descubrimiento, lectura o procesamiento de archivos master.
/// </summary>
public sealed class FileScannerService : IFileScannerService
{
    private readonly ILogger<FileScannerService> _logger;
    private readonly ScannerOptions _options;

    public FileScannerService(ILogger<FileScannerService> logger, ScannerOptions options)
    {
        _logger = logger;
        _options = options ?? new ScannerOptions();
    }

    public ScanResult EscanearDirectorio(string directorioRaiz)
    {
        if (!Directory.Exists(directorioRaiz))
            throw new DirectoryNotFoundException($"El directorio raíz no existe: {directorioRaiz}");

        var resultado = new ScanResult();

        // Escanear únicamente las carpetas de los alumnos, sin intentar descubrir masters
        DescubrirArchivosAlumnos(directorioRaiz, resultado);

        _logger.LogInformation(
            "Escaneo completado. Se omitió la búsqueda de archivos master. Alumnos encontrados: {Alumnos}.",
            resultado.ArchivosAlumnos.Count);

        return resultado;
    }

    private void DescubrirArchivosAlumnos(string directorioRaiz, ScanResult resultado)
    {
        foreach (var tema in _options.Temas)
        {
            var nombreCarpeta = $"Tema {tema}";
            var rutaCarpeta = Path.Combine(directorioRaiz, nombreCarpeta);

            if (!Directory.Exists(rutaCarpeta))
            {
                resultado.Advertencias.Add($"Carpeta de tema no encontrada: {nombreCarpeta}");
                _logger.LogWarning("Carpeta no encontrada: {Carpeta}", rutaCarpeta);
                continue;
            }

            var searchOption = _options.RecursiveSearch ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var archivosExcel = Directory.GetFiles(rutaCarpeta, "*.xlsx", searchOption).ToList();

            if (_options.IncludeXlsFiles)
                archivosExcel.AddRange(Directory.GetFiles(rutaCarpeta, "*.xls", searchOption));

            foreach (var archivo in archivosExcel)
            {
                var nombreArchivo = Path.GetFileNameWithoutExtension(archivo);

                // Excluir cualquier archivo temporal de Excel o master erróneo
                if (nombreArchivo.StartsWith("~$", StringComparison.OrdinalIgnoreCase) ||
                    nombreArchivo.Contains("Master", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                resultado.ArchivosAlumnos.Add(new StudentFileInfo
                {
                    RutaCompleta = archivo,
                    NombreAlumno = NormalizarNombreAlumno(nombreArchivo),
                    Tema = tema
                });

                _logger.LogDebug("Alumno encontrado: {Nombre} (Tema {Tema})", nombreArchivo, tema);
            }
        }
    }

    private static string NormalizarNombreAlumno(string nombreArchivo)
    {
        return nombreArchivo
            .Replace('_', ' ')
            .Replace('-', ' ')
            .Trim();
    }
}
