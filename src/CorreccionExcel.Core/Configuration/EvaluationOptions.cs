namespace CorreccionExcel.Core.Configuration;

public sealed class EvaluationOptions
{
    public double MultiplicadorArrastre { get; set; } = 0.66;
    public double PesoColumnasIntermedias { get; set; } = 20.0;
    public double PesoPromedioX { get; set; } = 5.0;
    public double PesoPromedioY { get; set; } = 5.0;
    public double PesoCovarianza { get; set; } = 15.0;
    public double PesoDesvioX { get; set; } = 7.5;
    public double PesoDesvioY { get; set; } = 7.5;
    public double PesoPearson { get; set; } = 15.0;
    public double PesoInterpretacion { get; set; } = 10.0;
    public double PesoPendiente { get; set; } = 7.5;
    public double PesoOrdenada { get; set; } = 7.5;

    public ToleranceOptions Tolerance { get; set; } = new();
    public CandidateSearchOptions CandidateSearch { get; set; } = new();
}

public sealed class ToleranceOptions
{
    public double EpsilonBase { get; set; } = 0.005;
    public double RelativePercentage { get; set; } = 0.05;
}

public sealed class CandidateSearchOptions
{
    public double RelativeWindowPercentage { get; set; } = 0.15;
    public double MinimumAbsoluteWindow { get; set; } = 1.0;
}
