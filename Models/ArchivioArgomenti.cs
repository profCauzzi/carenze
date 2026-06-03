using System.Text.Json.Serialization;

namespace CarenzeGenerator.Models;

public class ArchivioArgomenti
{
    [JsonPropertyName("metadati")]
    public Dictionary<string, object>? Metadati { get; set; }

    [JsonPropertyName("argomentiPerAnno")]
    public Dictionary<string, List<Argomento>> ArgomentiPerAnno { get; set; } = new();
}
