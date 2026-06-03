namespace CarenzeGenerator.Models;

public class Argomento
{
    public string Titolo { get; set; } = "";
    public List<string> ClassiOrigine { get; set; } = new();
    public List<string> Conoscenze { get; set; } = new();
    public List<string> Abilita { get; set; } = new();

    public Argomento Clone()
    {
        return new Argomento
        {
            Titolo = Titolo,
            ClassiOrigine = ClassiOrigine.ToList(),
            Conoscenze = Conoscenze.ToList(),
            Abilita = Abilita.ToList()
        };
    }
}
