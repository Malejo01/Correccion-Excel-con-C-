using System.Text.RegularExpressions;
using CorreccionExcel.Core.Configuration;
using CorreccionExcel.Core.Models;
using CorreccionExcel.Core.Services;
using Microsoft.Extensions.Logging;
using CorreccionExcel.Core.Validators;

namespace CorreccionExcel.Application.Services;

public sealed class EvaluationEngine : IEvaluationEngine
{
    private static readonly Regex RegexFuncionesEstadisticas = new(
        @"\b(AVERAGE|STDEV\.S|STDEV\.P|STDEV|STDEVP|COVARIANCE\.S|COVARIANCE\.P|COVAR|PEARSON|CORREL|SLOPE|INTERCEPT)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RegexFormulasAlgebraicas = new(
        @"(?:[A-Z]+\d+|\$[A-Z]+\$?\d+)[\s\S]*?[+\-*/^]|[+\-*/^][\s\S]*?(?:[A-Z]+\d+|\$[A-Z]+\$?\d+)|\b(SUM|SQRT|SUMPRODUCT|POWER|ABS|COUNT)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly ILogger<EvaluationEngine> _logger;
    private readonly EvaluationOptions _options;

    private static readonly Regex RegexTokenFuncion = new(@"\b([A-Z][A-Z0-9\.]*)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RegexReferenciaCelda = new(@"\$?[A-Z]{1,3}\$?\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RegexNumero = new(@"\b\d+(?:[\.,]\d+)?\b", RegexOptions.Compiled);

    private static readonly HashSet<string> FuncionesPromedio = new(StringComparer.OrdinalIgnoreCase)
    {
        "AVERAGE", "PROMEDIO"
    };

    private static readonly HashSet<string> FuncionesCovarianza = new(StringComparer.OrdinalIgnoreCase)
    {
        "COVARIANCE.S", "COVARIANCE.P", "COVAR", "COVARIANZA.M", "COVARIANZA.P"
    };

    private static readonly HashSet<string> FuncionesDesvio = new(StringComparer.OrdinalIgnoreCase)
    {
        "STDEV.S", "STDEV.P", "STDEV", "STDEVP", "DESVEST", "DESVEST.P", "DESVEST.M"
    };

    private static readonly HashSet<string> FuncionesPearson = new(StringComparer.OrdinalIgnoreCase)
    {
        "PEARSON", "CORREL", "COEF.DE.CORREL"
    };

    private static readonly HashSet<string> FuncionesPendiente = new(StringComparer.OrdinalIgnoreCase)
    {
        "SLOPE", "PENDIENTE"
    };

    private static readonly HashSet<string> FuncionesOrdenada = new(StringComparer.OrdinalIgnoreCase)
    {
        "INTERCEPT", "INTERSECCION.EJE"
    };

    public EvaluationEngine(ILogger<EvaluationEngine> logger, EvaluationOptions options)
    {
        _logger = logger;
        _options = options ?? new EvaluationOptions();
    }

    public StudentEvaluation Evaluar(string nombreAlumno, string rutaArchivo, string hojaDetectada, IReadOnlyList<CeldaNormalizada> celdasAlumno, MasterMetrics master)
    {
        var evaluacion = new StudentEvaluation
        {
            NombreAlumno = nombreAlumno,
            TemaAsignado = master.Tema,
            RutaArchivo = rutaArchivo,
            HojaDetectada = hojaDetectada
        };

        try
        {
            var contexto = new EvaluationContext
            {
                NombreAlumno = nombreAlumno,
                Tema = master.Tema,
                Master = master
            };

            var celdasNumericas = celdasAlumno.Where(c => c.EsNumerica).ToList();

            EvaluarColumnasIntermedias(celdasAlumno, evaluacion);
            EvaluarPromedios(celdasNumericas, contexto, evaluacion);
            EvaluarCovarianza(celdasNumericas, contexto, evaluacion);
            EvaluarDesvios(celdasNumericas, contexto, evaluacion);
            EvaluarPearson(celdasNumericas, contexto, evaluacion);
            EvaluarInterpretacionPearson(celdasAlumno, contexto, evaluacion);
            EvaluarRecta(celdasNumericas, contexto, evaluacion);

            double totalMetricas = evaluacion.ResultadosMetricas.Sum(r => r.PuntajeObtenido);
            evaluacion.NotaFinal = Math.Round(totalMetricas + evaluacion.PuntajeColumnasIntermedias, 2);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al evaluar alumno {Nombre}", nombreAlumno);
            evaluacion.HuboErrorProcesamiento = true;
            evaluacion.ErroresDetectados.Add($"Error interno durante la evaluación: {ex.Message}");
        }

        return evaluacion;
    }

    private void EvaluarColumnasIntermedias(IReadOnlyList<CeldaNormalizada> celdas, StudentEvaluation evaluacion)
    {
        var formulasConResta = celdas.Where(c => c.TieneFormula && c.Formula!.Contains('-')).ToList();
        int columnasConPatron = formulasConResta.GroupBy(c => c.Columna).Count(g => g.Count() >= 2 && TienePatronDiferenciaConFija(g.Select(c => c.Formula!)));
        int columnasConCuadrado = celdas.Where(c => c.TieneFormula && (c.Formula!.Contains('^') || c.Formula.Contains("POWER", StringComparison.OrdinalIgnoreCase))).GroupBy(c => c.Columna).Count();
        int columnasProducto = celdas.Where(c => c.TieneFormula && c.Formula!.Contains('*')).GroupBy(c => c.Columna).Count(g => g.Count() >= 2);

        bool tieneCol1 = columnasConPatron >= 1;
        bool tieneCol2 = columnasConPatron >= 2;
        bool tieneCol3 = columnasProducto >= 1;
        bool tieneCol4 = columnasConCuadrado >= 1;
        bool tieneCol5 = columnasConCuadrado >= 2;

        int columnasDetectadas = 0;
        if (tieneCol1) columnasDetectadas++;
        if (tieneCol2) columnasDetectadas++;
        if (tieneCol3) columnasDetectadas++;
        if (tieneCol4) columnasDetectadas++;
        if (tieneCol5) columnasDetectadas++;

        bool tienePrimerasTres = tieneCol1 && tieneCol2 && tieneCol3;
        bool tieneCuadrados = tieneCol4 && tieneCol5;

        if (tienePrimerasTres || tieneCuadrados)
        {
            evaluacion.ColumnasIntermediasEncontradas = columnasDetectadas;
            evaluacion.TieneColumnasIntermedias = true;
            evaluacion.PuntajeColumnasIntermedias = _options.PesoColumnasIntermedias;
        }
        else
        {
            evaluacion.ColumnasIntermediasEncontradas = columnasDetectadas;
            evaluacion.TieneColumnasIntermedias = columnasDetectadas == 5;
            evaluacion.PuntajeColumnasIntermedias = Math.Round(columnasDetectadas * (_options.PesoColumnasIntermedias / 5.0), 2);
        }
    }

    private static bool TienePatronDiferenciaConFija(IEnumerable<string> formulas)
    {
        var patronRef = new Regex(@"[A-Z]+\d+\s*-\s*\$[A-Z]+\$\d+|\$[A-Z]+\$\d+\s*-\s*[A-Z]+\d+", RegexOptions.IgnoreCase);
        return formulas.Any(f => patronRef.IsMatch(f));
    }

    private void EvaluarPromedios(List<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        var resultadoX = EvaluarMetricaSimple(celdas, ctx.Master.PromedioX, "PromedioX", _options.PesoPromedioX, null);
        var resultadoY = EvaluarMetricaSimple(celdas, ctx.Master.PromedioY, "PromedioY", _options.PesoPromedioY, null);

        evaluacion.ResultadosMetricas.Add(resultadoX);
        evaluacion.ResultadosMetricas.Add(resultadoY);

        ctx.AlumnoPromedioX = resultadoX.ValorAlumno ?? ctx.Master.PromedioX;
        ctx.AlumnoPromedioY = resultadoY.ValorAlumno ?? ctx.Master.PromedioY;
    }

    private void EvaluarCovarianza(List<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        double covArrastre = RecalcularCovarianza(ctx.Master.DatosX, ctx.Master.DatosY, ctx.AlumnoPromedioX!.Value, ctx.AlumnoPromedioY!.Value);
        var resultado = EvaluarMetricaSimple(celdas, ctx.Master.Covarianza, "Covarianza", _options.PesoCovarianza, AplicarSiDifiere(covArrastre, ctx.Master.Covarianza));
        evaluacion.ResultadosMetricas.Add(resultado);
        ctx.AlumnoCovarianza = resultado.ValorAlumno ?? covArrastre;
    }

    private void EvaluarDesvios(List<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        double desvXArrastre = RecalcularDesvio(ctx.Master.DatosX, ctx.AlumnoPromedioX!.Value);
        double desvYArrastre = RecalcularDesvio(ctx.Master.DatosY, ctx.AlumnoPromedioY!.Value);

        var resultadoX = EvaluarMetricaSimple(celdas, ctx.Master.DesvioEstandarX, "DesvioX", _options.PesoDesvioX, AplicarSiDifiere(desvXArrastre, ctx.Master.DesvioEstandarX));
        var resultadoY = EvaluarMetricaSimple(celdas, ctx.Master.DesvioEstandarY, "DesvioY", _options.PesoDesvioY, AplicarSiDifiere(desvYArrastre, ctx.Master.DesvioEstandarY));

        evaluacion.ResultadosMetricas.Add(resultadoX);
        evaluacion.ResultadosMetricas.Add(resultadoY);

        ctx.AlumnoDesvioX = resultadoX.ValorAlumno ?? desvXArrastre;
        ctx.AlumnoDesvioY = resultadoY.ValorAlumno ?? desvYArrastre;
    }

    private void EvaluarPearson(List<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        double pearsonArrastre = RecalcularPearson(ctx.AlumnoCovarianza!.Value, ctx.AlumnoDesvioX!.Value, ctx.AlumnoDesvioY!.Value);
        var resultado = EvaluarMetricaSimple(celdas, ctx.Master.CorrelacionPearson, "CorrelacionPearson", _options.PesoPearson, AplicarSiDifiere(pearsonArrastre, ctx.Master.CorrelacionPearson));
        evaluacion.ResultadosMetricas.Add(resultado);
        ctx.AlumnoPearson = resultado.ValorAlumno ?? pearsonArrastre;
    }

    private void EvaluarInterpretacionPearson(IReadOnlyList<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        var pearsonNumerico = ctx.AlumnoPearson ?? ctx.Master.CorrelacionPearson;
        var celdasTexto = celdas.Where(c => !c.EsNumerica && !string.IsNullOrWhiteSpace(c.TextoValor)).ToList();
        
        string? interpretacionAlumno = BuscarInterpretacionTexto(celdasTexto, pearsonNumerico);
        bool esCorrecta = interpretacionAlumno is not null;

        evaluacion.ResultadosMetricas.Add(ConstruirResultado("InterpretacionPearson", pearsonNumerico, null, null, null, 
            esCorrecta ? MetricMatchType.MasterMatch : MetricMatchType.NoMatch, _options.PesoInterpretacion, esCorrecta ? _options.PesoInterpretacion : 0,
            esCorrecta ? $"Interpretación correcta: '{interpretacionAlumno}'" : "No se encontró interpretación o no concuerda con el valor R."));
    }

    private void EvaluarRecta(List<CeldaNormalizada> celdas, EvaluationContext ctx, StudentEvaluation evaluacion)
    {
        double pendienteArrastre = RecalcularPendiente(ctx.AlumnoCovarianza!.Value, ctx.AlumnoDesvioX!.Value);
        double ordenadaArrastre = RecalcularOrdenada(ctx.AlumnoPromedioY!.Value, pendienteArrastre, ctx.AlumnoPromedioX!.Value);

        var resultadoPendiente = EvaluarMetricaSimple(celdas, ctx.Master.PendienteRecta, "PendienteRecta", _options.PesoPendiente, AplicarSiDifiere(pendienteArrastre, ctx.Master.PendienteRecta));
        var resultadoOrdenada = EvaluarMetricaSimple(celdas, ctx.Master.OrdenadaOrigen, "OrdenadaOrigen", _options.PesoOrdenada, AplicarSiDifiere(ordenadaArrastre, ctx.Master.OrdenadaOrigen));

        evaluacion.ResultadosMetricas.Add(resultadoPendiente);
        evaluacion.ResultadosMetricas.Add(resultadoOrdenada);
    }

    private MetricCheckResult EvaluarMetricaSimple(IEnumerable<CeldaNormalizada> celdas, double valorEsperado, string nombreMetrica, double pesoBase, double? valorRecalculado)
    {
        var candidata = BuscarCeldaCandidataMasProxima(celdas, valorEsperado, valorRecalculado, nombreMetrica);

        if (candidata is null)
            return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, null, null, MetricMatchType.NotFound, pesoBase, 0, "No se encontró celda candidata.");

        double valorAlumno = candidata.ValorNumerico!.Value;
        bool formulaLegitima = IsLegitimateCalculation(candidata, nombreMetrica);

        if (CoincideConTolerancia(valorAlumno, valorEsperado))
        {
            if (!formulaLegitima)
                return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, valorAlumno, candidata, MetricMatchType.Hardcoded, pesoBase, 0, "Valor correcto pero hardcodeado.");

            return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, valorAlumno, candidata, MetricMatchType.MasterMatch, pesoBase, pesoBase, "Match perfecto con el master.");
        }

        if (valorRecalculado.HasValue && CoincideConTolerancia(valorAlumno, valorRecalculado.Value))
        {
            if (!formulaLegitima)
                return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, valorAlumno, candidata, MetricMatchType.Hardcoded, pesoBase, 0, "Coincide con arrastre pero está hardcodeado.");

            return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, valorAlumno, candidata, MetricMatchType.CarryOverMatch, pesoBase, Math.Round(pesoBase * _options.MultiplicadorArrastre, 2), $"Arrastre válido. Procedimiento correcto.");
        }

        return ConstruirResultado(nombreMetrica, valorEsperado, valorRecalculado, valorAlumno, candidata, MetricMatchType.NoMatch, pesoBase, 0, $"Sin coincidencia. Alumno={valorAlumno:F4}, Esperado={valorEsperado:F4}");
    }

    private static bool IsLegitimateCalculation(CeldaNormalizada celda, string nombreMetrica)
    {
        if (!celda.TieneFormula || string.IsNullOrWhiteSpace(celda.Formula)) return false;
        string formula = celda.Formula;

        // Filtro específico por métrica para evitar falsos positivos por cercanía numérica.
        if (FormulaEsCompatibleConMetrica(formula, nombreMetrica))
            return true;

        return RegexFuncionesEstadisticas.IsMatch(formula) || RegexFormulasAlgebraicas.IsMatch(formula);
    }

    private CeldaNormalizada? BuscarCeldaCandidataMasProxima(IEnumerable<CeldaNormalizada> celdas, double valorEsperado, double? valorRecalculado, string nombreMetrica)
    {
        double ventanaMaster = Math.Max(Math.Abs(valorEsperado) * _options.CandidateSearch.RelativeWindowPercentage, _options.CandidateSearch.MinimumAbsoluteWindow);
        double ventanaArrastre = valorRecalculado.HasValue ? Math.Max(Math.Abs(valorRecalculado.Value) * _options.CandidateSearch.RelativeWindowPercentage, _options.CandidateSearch.MinimumAbsoluteWindow) : 0;

        var celdasNumericas = celdas.Where(c => c.EsNumerica).ToList();

        bool EnVentana(double val)
        {
            double distMaster = Math.Abs(val - valorEsperado);
            if (distMaster <= ventanaMaster)
                return true;

            if (!valorRecalculado.HasValue)
                return false;

            double distArrastre = Math.Abs(val - valorRecalculado.Value);
            return distArrastre <= ventanaArrastre;
        }

        var formulasCompatibles = celdasNumericas
            .Where(c => c.TieneFormula && c.Formula is not null && FormulaEsCompatibleConMetrica(c.Formula, nombreMetrica) && EnVentana(c.ValorNumerico!.Value));

        var mejorPorFormula = EncontrarCeldaMasProxima(formulasCompatibles, valorEsperado, valorRecalculado);
        if (mejorPorFormula is not null)
            return mejorPorFormula;

        var candidatasPorVentana = celdasNumericas.Where(c => EnVentana(c.ValorNumerico!.Value));
        return EncontrarCeldaMasProxima(candidatasPorVentana, valorEsperado, valorRecalculado);
    }

    private static CeldaNormalizada? EncontrarCeldaMasProxima(IEnumerable<CeldaNormalizada> celdas, double valorEsperado, double? valorRecalculado)
    {
        CeldaNormalizada? mejor = null;
        double menorDist = double.MaxValue;

        foreach (var celda in celdas)
        {
            double val = celda.ValorNumerico!.Value;
            double distMaster = Math.Abs(val - valorEsperado);
            double distArrastre = valorRecalculado.HasValue ? Math.Abs(val - valorRecalculado.Value) : double.MaxValue;
            double dist = Math.Min(distMaster, distArrastre);

            if (dist < menorDist)
            {
                menorDist = dist;
                mejor = celda;
            }
        }

        return mejor;
    }

    private static bool FormulaEsCompatibleConMetrica(string formula, string nombreMetrica)
    {
        var analisis = AnalizarFormula(formula);

        return nombreMetrica switch
        {
            "PromedioX" or "PromedioY" => ContieneFuncion(analisis, FuncionesPromedio),
            "Covarianza" =>
                ContieneFuncion(analisis, FuncionesCovarianza) ||
                (analisis.TieneReferencias && analisis.TieneResta && analisis.TieneMultiplicacion),
            "DesvioX" or "DesvioY" =>
                ContieneFuncion(analisis, FuncionesDesvio) ||
                (analisis.TieneRaiz || (analisis.TienePotencia && analisis.TieneDivision)),
            "CorrelacionPearson" =>
                ContieneFuncion(analisis, FuncionesPearson) ||
                (analisis.TieneDivision && analisis.TieneReferencias && analisis.TieneParentesis),
            "PendienteRecta" =>
                ContieneFuncion(analisis, FuncionesPendiente) ||
                (analisis.TieneDivision && analisis.TieneReferencias),
            "OrdenadaOrigen" =>
                ContieneFuncion(analisis, FuncionesOrdenada) ||
                (analisis.TieneResta && analisis.TieneMultiplicacion && analisis.TieneReferencias),
            _ => RegexFuncionesEstadisticas.IsMatch(formula) || RegexFormulasAlgebraicas.IsMatch(formula)
        };
    }

    private static bool ContieneFuncion(FormulaAnalysis analisis, HashSet<string> funcionesEsperadas)
        => analisis.Funciones.Any(funcionesEsperadas.Contains);

    private static FormulaAnalysis AnalizarFormula(string formula)
    {
        string texto = formula.ToUpperInvariant();
        var funciones = RegexTokenFuncion.Matches(texto)
            .Select(m => m.Groups[1].Value)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return new FormulaAnalysis(
            Funciones: funciones,
            TieneReferencias: RegexReferenciaCelda.IsMatch(texto),
            TieneNumeros: RegexNumero.IsMatch(texto),
            TieneResta: texto.Contains('-'),
            TieneMultiplicacion: texto.Contains('*'),
            TieneDivision: texto.Contains('/'),
            TienePotencia: texto.Contains('^') || texto.Contains("POWER", StringComparison.OrdinalIgnoreCase) || texto.Contains("POTENCIA", StringComparison.OrdinalIgnoreCase),
            TieneRaiz: texto.Contains("SQRT", StringComparison.OrdinalIgnoreCase) || texto.Contains("RAIZ", StringComparison.OrdinalIgnoreCase),
            TieneParentesis: texto.Contains('(') && texto.Contains(')'));
    }

    private sealed record FormulaAnalysis(
        HashSet<string> Funciones,
        bool TieneReferencias,
        bool TieneNumeros,
        bool TieneResta,
        bool TieneMultiplicacion,
        bool TieneDivision,
        bool TienePotencia,
        bool TieneRaiz,
        bool TieneParentesis);

    private static double RecalcularCovarianza(double[] x, double[] y, double promedioX, double promedioY)
    {
        int n = Math.Min(x.Length, y.Length);
        if (n == 0) return 0;
        return Enumerable.Range(0, n).Sum(i => (x[i] - promedioX) * (y[i] - promedioY)) / n;
    }

    private static double RecalcularDesvio(double[] datos, double promedio)
    {
        if (datos.Length == 0) return 0;
        return Math.Sqrt(datos.Sum(d => Math.Pow(d - promedio, 2)) / datos.Length);
    }

    private static double RecalcularPearson(double covarianza, double desvX, double desvY) => (desvX * desvY) == 0 ? 0 : covarianza / (desvX * desvY);
    private static double RecalcularPendiente(double covarianza, double desvX) => (desvX * desvX) == 0 ? 0 : covarianza / (desvX * desvX);
    private static double RecalcularOrdenada(double promedioY, double pendiente, double promedioX) => promedioY - pendiente * promedioX;

    private static string? BuscarInterpretacionTexto(IEnumerable<CeldaNormalizada> celdasTexto, double pearsonAlumno)
    {
        foreach (var celda in celdasTexto)
        {
            if (PearsonInterpretationValidator.IsValidInterpretation(celda.TextoValor, pearsonAlumno))
                return celda.TextoValor;
        }
        return null;
    }

    private bool CoincideConTolerancia(double alumno, double esperado)
        => ToleranceValidator.IsWithinTolerance(alumno, esperado, _options.Tolerance.EpsilonBase, _options.Tolerance.RelativePercentage);

    private double? AplicarSiDifiere(double valorArrastre, double valorMaster) => CoincideConTolerancia(valorArrastre, valorMaster) ? null : valorArrastre;

    private static MetricCheckResult ConstruirResultado(string nombre, double esperado, double? recalculado, double? valorAlumno, CeldaNormalizada? celda, MetricMatchType tipo, double pesoBase, double puntaje, string observacion)
    {
        return new MetricCheckResult
        {
            MetricName = nombre, ValorEsperado = esperado, ValorRecalculado = recalculado, ValorAlumno = valorAlumno,
            TipoMatch = tipo, TieneFormula = celda?.TieneFormula ?? false, Formula = celda?.Formula ?? string.Empty,
            CeldaDireccion = celda?.Direccion ?? string.Empty, PesoBase = pesoBase, PuntajeObtenido = puntaje, Observacion = observacion
        };
    }
}