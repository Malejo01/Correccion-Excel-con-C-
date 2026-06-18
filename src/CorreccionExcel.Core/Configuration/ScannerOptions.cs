namespace CorreccionExcel.Core.Configuration;

public sealed class ScannerOptions
{
    public List<string> Temas { get; set; } = ["A", "B", "C", "D"];
    public bool RecursiveSearch { get; set; }
    public bool IncludeXlsFiles { get; set; }
}
