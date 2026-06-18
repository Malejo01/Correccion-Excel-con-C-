using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CorreccionExcel.Core.Models;
using CorreccionExcel.Core.Services;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace CorreccionExcel.Infrastructure.Services;

/// <summary>
/// Implementación EPPlus del servicio de lectura de Excel.
/// Responsabilidad única: interactuar con la librería y devolver datos normalizados de alumnos,
/// abstrayéndose completamente de cualquier lectura o parseo de masters.
/// </summary>
public sealed class ExcelReaderService : IExcelReaderService
{
    private readonly ILogger<ExcelReaderService> _logger;

    public string UltimaHojaSeleccionada { get; private set; } = string.Empty;
    public double[] UltimosDatosX { get; private set; } = [];
    public double[] UltimosDatosY { get; private set; } = [];

    public ExcelReaderService(ILogger<ExcelReaderService> logger)
    {
        _logger = logger;
    }

    // ─── Lectura del archivo del alumno ───────────────────────────────────────

    public IReadOnlyList<CeldaNormalizada> LeerHojaAlumno(string rutaArchivo)
    {
        UltimosDatosX = [];
        UltimosDatosY = [];
        UltimaHojaSeleccionada = string.Empty;

        using var paquete = AbrirPaquete(rutaArchivo);
        if (paquete is null) return [];

        var hoja = SeleccionarHojaPorContenido(paquete.Workbook.Worksheets, rutaArchivo);
        if (hoja is null) return [];

        UltimaHojaSeleccionada = hoja.Name;

        // Extraer dinámicamente las series de datos X e Y reales del desarrollo del alumno
        var (datosX, datosY) = ExtraerDatosBivariados(hoja);
        UltimosDatosX = datosX;
        UltimosDatosY = datosY;

        _logger.LogDebug(
            "Archivo {Archivo} — Datos extraídos: X_Count={LenX}, Y_Count={LenY}",
            Path.GetFileName(rutaArchivo), datosX.Length, datosY.Length);

        return ExtraerCeldas(hoja);
    }

    // ─── Selección de hoja por contenido (mayor densidad numérica) ───────────

    private ExcelWorksheet? SeleccionarHojaPorContenido(ExcelWorksheets hojas, string rutaArchivo)
    {
        ExcelWorksheet? mejorHoja = null;
        double mejorScore = double.MinValue;

        foreach (var hoja in hojas)
        {
            if (hoja.Dimension is null) continue;

            int densidad = ContarCeldasNumericas(hoja);
            int formulas = ContarCeldasConFormula(hoja);
            int columnasConSerie = ContarColumnasConSerieNumerica(hoja, minLargo: 3);
            double score = CalcularScoreHoja(densidad, formulas, columnasConSerie);

            _logger.LogDebug(
                "Archivo {Archivo} — Hoja '{Nombre}': densidad={Densidad}, formulas={Formulas}, columnasSerie={ColumnasSerie}, score={Score}.",
                Path.GetFileName(rutaArchivo), hoja.Name, densidad, formulas, columnasConSerie, score);

            if (score > mejorScore)
            {
                mejorScore = score;
                mejorHoja = hoja;
            }
        }

        if (mejorHoja is null)
            _logger.LogWarning("No se encontró ninguna hoja con datos numéricos en: {Archivo}", rutaArchivo);

        return mejorHoja;
    }

    private static double CalcularScoreHoja(int densidad, int formulas, int columnasConSerie)
    {
        // Heurística compuesta: prioriza trabajo real (series y fórmulas) sin ignorar densidad numérica.
        return densidad + (formulas * 2.0) + (columnasConSerie * 5.0);
    }

    private static int ContarCeldasNumericas(ExcelWorksheet hoja)
    {
        if (hoja.Dimension is null) return 0;

        int count = 0;
        for (int fila = hoja.Dimension.Start.Row; fila <= hoja.Dimension.End.Row; fila++)
        {
            for (int col = hoja.Dimension.Start.Column; col <= hoja.Dimension.End.Column; col++)
            {
                var celda = hoja.Cells[fila, col];
                if (celda.Value is not null && TryParseNumerico(celda.Value, out _))
                    count++;
            }
        }
        return count;
    }

    private static int ContarCeldasConFormula(ExcelWorksheet hoja)
    {
        if (hoja.Dimension is null) return 0;

        int count = 0;
        for (int fila = hoja.Dimension.Start.Row; fila <= hoja.Dimension.End.Row; fila++)
        {
            for (int col = hoja.Dimension.Start.Column; col <= hoja.Dimension.End.Column; col++)
            {
                var celda = hoja.Cells[fila, col];
                if (!string.IsNullOrWhiteSpace(celda.Formula))
                    count++;
            }
        }
        return count;
    }

    private static int ContarColumnasConSerieNumerica(ExcelWorksheet hoja, int minLargo)
    {
        if (hoja.Dimension is null) return 0;

        int columnasConSerie = 0;
        for (int col = hoja.Dimension.Start.Column; col <= hoja.Dimension.End.Column; col++)
        {
            if (ExtraerSerieContigua(hoja, col).Values.Length >= minLargo)
                columnasConSerie++;
        }

        return columnasConSerie;
    }

    // ─── Extracción de celdas ─────────────────────────────────────────────────

    private static IReadOnlyList<CeldaNormalizada> ExtraerCeldas(ExcelWorksheet hoja)
    {
        if (hoja.Dimension is null) return [];

        var resultado = new List<CeldaNormalizada>();

        for (int fila = hoja.Dimension.Start.Row; fila <= hoja.Dimension.End.Row; fila++)
        {
            for (int col = hoja.Dimension.Start.Column; col <= hoja.Dimension.End.Column; col++)
            {
                var celda = hoja.Cells[fila, col];
                if (celda.Value is null) continue;

                double? valorNumerico = null;
                string? textoValor = null;

                if (TryParseNumerico(celda.Value, out double numVal))
                    valorNumerico = numVal;
                else
                    textoValor = celda.Value?.ToString();

                // Solo incluir celdas que tengan algún dato relevante
                if (valorNumerico is null && string.IsNullOrWhiteSpace(textoValor)) continue;

                resultado.Add(new CeldaNormalizada
                {
                    Fila = fila,
                    Columna = col,
                    Direccion = celda.Address,
                    ValorNumerico = valorNumerico,
                    Formula = string.IsNullOrWhiteSpace(celda.Formula) ? null : NormalizarFormula(celda.Formula),
                    TextoValor = textoValor
                });
            }
        }

        return resultado.AsReadOnly();
    }

    // ─── Extracción bivariada para el arrastre ───────────────────────────────

    private static (double[] datosX, double[] datosY) ExtraerDatosBivariados(ExcelWorksheet hoja)
    {
        if (hoja.Dimension is null) return ([], []);

        var series = new List<SerieColumna>();

        for (int col = hoja.Dimension.Start.Column; col <= hoja.Dimension.End.Column; col++)
        {
            var extraida = ExtraerSerieContigua(hoja, col);
            if (extraida.Values.Length >= 2)
                series.Add(new SerieColumna(col, extraida.StartRow, extraida.Values));
        }

        if (series.Count < 2) return ([], []);

        var mejorPar = BuscarMejorPar(series);
        if (mejorPar is not null)
        {
            var x = mejorPar.Value.X.Columna <= mejorPar.Value.Y.Columna ? mejorPar.Value.X : mejorPar.Value.Y;
            var y = mejorPar.Value.X.Columna <= mejorPar.Value.Y.Columna ? mejorPar.Value.Y : mejorPar.Value.X;
            return ([.. x.Values], [.. y.Values]);
        }

        // Fallback: dos columnas más densas (comportamiento anterior) si no hay par coherente.
        var top2 = series
            .OrderByDescending(s => s.Values.Length)
            .Take(2)
            .OrderBy(s => s.Columna)
            .ToList();

        if (top2.Count < 2) return ([], []);
        return ([.. top2[0].Values], [.. top2[1].Values]);
    }

    private static (int StartRow, double[] Values) ExtraerSerieContigua(ExcelWorksheet hoja, int columna)
    {
        if (hoja.Dimension is null) return (0, []);

        int startRow = -1;
        var secuencia = new List<double>();

        for (int fila = hoja.Dimension.Start.Row; fila <= hoja.Dimension.End.Row; fila++)
        {
            var celda = hoja.Cells[fila, columna];
            if (TryParseNumerico(celda.Value, out double val))
            {
                if (startRow < 0)
                    startRow = fila;

                secuencia.Add(val);
                continue;
            }

            if (secuencia.Count > 0)
                break;
        }

        return startRow < 0 ? (0, []) : (startRow, [.. secuencia]);
    }

    private static (SerieColumna X, SerieColumna Y)? BuscarMejorPar(List<SerieColumna> series)
    {
        (SerieColumna X, SerieColumna Y)? mejor = null;
        double mejorScore = double.MinValue;

        for (int i = 0; i < series.Count; i++)
        {
            for (int j = i + 1; j < series.Count; j++)
            {
                var a = series[i];
                var b = series[j];
                double score = CalcularScorePar(a, b);

                if (score > mejorScore)
                {
                    mejorScore = score;
                    mejor = (a, b);
                }
            }
        }

        return mejor;
    }

    private static double CalcularScorePar(SerieColumna a, SerieColumna b)
    {
        int lenA = a.Values.Length;
        int lenB = b.Values.Length;
        int endA = a.StartRow + lenA - 1;
        int endB = b.StartRow + lenB - 1;

        int overlap = Math.Max(0, Math.Min(endA, endB) - Math.Max(a.StartRow, b.StartRow) + 1);
        int gapColumnas = Math.Abs(a.Columna - b.Columna);
        int diferenciaLargo = Math.Abs(lenA - lenB);

        double score = 0;
        score += overlap * 4.0;
        score += Math.Min(lenA, lenB) * 1.5;
        score -= diferenciaLargo * 1.25;
        score -= Math.Max(0, gapColumnas - 1) * 2.0;

        return score;
    }

    private sealed record SerieColumna(int Columna, int StartRow, double[] Values);

    // ─── Helpers ──────────────────────────────────────────────────────────────

    private ExcelPackage? AbrirPaquete(string rutaArchivo)
    {
        try
        {
            return new ExcelPackage(new FileInfo(rutaArchivo));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "No se pudo abrir el archivo Excel: {Archivo}", rutaArchivo);
            return null;
        }
    }

    private static bool TryParseNumerico(object? valor, out double resultado)
    {
        resultado = 0;
        if (valor is null) return false;
        if (valor is double d) { resultado = d; return true; }
        if (valor is int i) { resultado = i; return true; }
        if (valor is long l) { resultado = l; return true; }
        if (valor is decimal dec) { resultado = (double)dec; return true; }
        if (valor is float f) { resultado = f; return true; }
        return double.TryParse(valor.ToString(), System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out resultado);
    }

    private static string NormalizarFormula(string formula)
    {
        string normalizada = formula.Trim();
        return normalizada.StartsWith('=') ? normalizada : $"={normalizada}";
    }
}
