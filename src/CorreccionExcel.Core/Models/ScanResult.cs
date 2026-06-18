namespace CorreccionExcel.Core.Models;

/// <summary>
/// Información de un archivo de alumno descubierto por el escáner.
/// </summary>
public sealed class StudentFileInfo
{
    public string RutaCompleta { get; init; } = string.Empty;
    public string NombreAlumno { get; init; } = string.Empty;
    public string Tema { get; init; } = string.Empty;
}

/// <summary>
/// Resultado completo del escaneo del directorio raíz.
/// </summary>
public sealed class ScanResult
{
    public Dictionary<string, string> MasterPorTema { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public List<StudentFileInfo> ArchivosAlumnos { get; init; } = [];
    public List<string> Advertencias { get; init; } = [];
}
