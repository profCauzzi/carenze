namespace CarenzeGenerator.Models;

public class Studente
{
    public string Nome { get; set; } = "";
    public string Cognome { get; set; } = "";
    public string Classe { get; set; } = "";
    public string Disciplina { get; set; } = "Informatica";
    public string CarenzeRilevate { get; set; } = "";
    public List<Argomento> Argomenti { get; set; } = new();
    public bool ProvaAccertamento { get; set; } = true;
    public List<string> TipologieAccertamento { get; set; } = new();
    public string Data { get; set; } = "";
    public string Docente { get; set; } = "";
}
