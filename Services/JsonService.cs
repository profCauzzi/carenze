using System.Text.Json;
using CarenzeGenerator.Models;

namespace CarenzeGenerator.Services;

public static class JsonService
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public static Dictionary<string, List<Argomento>> CaricaArgomentiPerAnno(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File argomenti non trovato: {path}");

        string json = File.ReadAllText(path);
        var archivio = JsonSerializer.Deserialize<ArchivioArgomenti>(json, Options);

        if (archivio?.ArgomentiPerAnno == null || archivio.ArgomentiPerAnno.Count == 0)
            throw new InvalidDataException("Il file degli argomenti non contiene argomentiPerAnno oppure è vuoto.");

        return archivio.ArgomentiPerAnno;
    }

    public static List<Studente> CaricaStudentiSePresenti(string path)
    {
        if (!File.Exists(path))
            return new List<Studente>();

        string json = File.ReadAllText(path);

        if (string.IsNullOrWhiteSpace(json))
            return new List<Studente>();

        return JsonSerializer.Deserialize<List<Studente>>(json, Options) ?? new List<Studente>();
    }

    public static void SalvaStudenti(string path, List<Studente> studenti)
    {
        string? dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir))
            Directory.CreateDirectory(dir);

        string json = JsonSerializer.Serialize(studenti, Options);
        File.WriteAllText(path, json);
    }
}
