using CarenzeGenerator.Models;
using CarenzeGenerator.Services;

namespace CarenzeGenerator;

internal class Program
{
    private static readonly string DataDir = "Data";
    private static readonly string OutputDir = "output";
    private static readonly string ArgomentiPath = Path.Combine(DataDir, "argomenti_per_anno_ottimizzato.json");
    private static readonly string StudentiPath = Path.Combine(DataDir, "studenti.json");

    private static List<Studente> _studenti = new();
    private static Dictionary<string, List<Argomento>> _argomentiPerAnno = new();

    private static readonly List<string> CarenzePredefinite = new()
    {
        "Presenta difficoltà nell'analisi e nella risoluzione di problemi di carattere informatico.",
        "Evidenzia una limitata autonomia nell'applicazione delle procedure operative.",
        "Mostra incertezze nell'utilizzo dei principali strumenti e linguaggi affrontati durante l'anno.",
        "Necessita di consolidare le competenze logico-deduttive e di problem solving.",
        "Manifesta difficoltà nel trasferire le conoscenze teoriche alle attività pratiche e laboratoriali.",
        "Presenta una conoscenza frammentaria dei contenuti disciplinari fondamentali.",
        "Evidenzia carenze metodologiche che incidono sull'efficacia dello studio individuale.",
        "Necessita di migliorare la precisione e la correttezza nell'esecuzione delle attività proposte.",
        "Mostra difficoltà nell'individuazione autonoma delle strategie risolutive più adeguate.",
        "Necessita di consolidare le competenze progettuali e la capacità di verificare criticamente i risultati ottenuti.",
        "Presenta difficoltà nel riconoscere e correggere errori procedurali e logici.",
        "Evidenzia una partecipazione non sempre efficace alle attività laboratoriali, con conseguente acquisizione parziale delle competenze previste."
    };

    static void Main()
    {
        Directory.CreateDirectory(DataDir);
        Directory.CreateDirectory(OutputDir);

        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.InputEncoding = System.Text.Encoding.UTF8;

        try
        {
            _argomentiPerAnno = JsonService.CaricaArgomentiPerAnno(ArgomentiPath);
            _studenti = JsonService.CaricaStudentiSePresenti(StudentiPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Errore iniziale: " + ex.Message);
            Console.WriteLine("Verifica che il file Data/argomenti_per_anno_ottimizzato.json sia presente e corretto.");
            return;
        }

        bool esci = false;

        while (!esci)
        {
            StampaMenu();
            string scelta = LeggiStringa("Scelta: ");

            switch (scelta)
            {
                case "1":
                    InserisciStudente();
                    break;
                case "2":
                    ElencaStudenti();
                    break;
                case "3":
                    ModificaStudente();
                    break;
                case "4":
                    EliminaStudente();
                    break;
                case "5":
                    GeneraSchede();
                    break;
                case "6":
                    SalvaStudenti();
                    break;
                case "7":
                    RicaricaStudenti();
                    break;
                case "8":
                    ElencaArgomentiDisponibili();
                    break;
                case "0":
                    SalvaStudenti();
                    esci = true;
                    break;
                default:
                    Console.WriteLine("Scelta non valida.");
                    break;
            }
        }
    }

    private static void StampaMenu()
    {
        Console.WriteLine();
        Console.WriteLine("==============================================");
        Console.WriteLine(" GENERATORE SCHEDE CARENZE - INFORMATICA");
        Console.WriteLine("==============================================");
        Console.WriteLine("1. Inserisci nuovo studente");
        Console.WriteLine("2. Elenca studenti inseriti");
        Console.WriteLine("3. Modifica studente");
        Console.WriteLine("4. Elimina studente");
        Console.WriteLine("5. Genera schede DOCX");
        Console.WriteLine("6. Salva archivio studenti");
        Console.WriteLine("7. Ricarica archivio studenti");
        Console.WriteLine("8. Visualizza argomenti disponibili");
        Console.WriteLine("0. Salva ed esci");
        Console.WriteLine("==============================================");
    }

    private static void InserisciStudente()
    {
        Console.WriteLine();
        Console.WriteLine("--- Nuovo studente ---");

        var studente = new Studente
        {
            Nome = LeggiStringaObbligatoria("Nome: "),
            Cognome = LeggiStringaObbligatoria("Cognome: "),
            Classe = LeggiStringaObbligatoria("Classe, es. 1F: ").ToUpperInvariant(),
            Disciplina = LeggiStringaConDefault("Disciplina", "Informatica"),
            CarenzeRilevate = SelezionaCarenzeRilevate()
        };

        string anno = AnnoDaClasse(studente.Classe);

        if (!_argomentiPerAnno.ContainsKey(anno))
        {
            Console.WriteLine($"Non sono presenti argomenti per l'anno '{anno}'.");
            Console.WriteLine("Lo studente sarà inserito senza argomenti. Potrai modificarlo dopo.");
        }
        else
        {
            studente.Argomenti = SelezionaArgomenti(_argomentiPerAnno[anno]);
        }

        studente.ProvaAccertamento = LeggiSiNo("Prova di accertamento? [S/N]: ", true);
        studente.TipologieAccertamento = SelezionaTipologie();

        studente.Data = LeggiStringaConDefault("Data scheda", DateTime.Now.ToString("dd/MM/yyyy"));
        studente.Docente = LeggiStringaConDefault("Docente", "");

        _studenti.Add(studente);
        Console.WriteLine("Studente inserito correttamente.");
    }

    private static void ElencaStudenti()
    {
        Console.WriteLine();
        Console.WriteLine("--- Studenti inseriti ---");

        if (_studenti.Count == 0)
        {
            Console.WriteLine("Nessuno studente inserito.");
            return;
        }

        for (int i = 0; i < _studenti.Count; i++)
        {
            var s = _studenti[i];
            Console.WriteLine($"{i + 1}. {s.Classe} - {s.Cognome} {s.Nome} - Argomenti: {s.Argomenti.Count}");
        }
    }

    private static void ModificaStudente()
    {
        if (_studenti.Count == 0)
        {
            Console.WriteLine("Nessuno studente da modificare.");
            return;
        }

        ElencaStudenti();
        int index = LeggiIntero("Numero studente da modificare: ", 1, _studenti.Count) - 1;
        var s = _studenti[index];

        Console.WriteLine();
        Console.WriteLine($"--- Modifica {s.Cognome} {s.Nome} ---");
        Console.WriteLine("Premi INVIO per mantenere il valore corrente.");

        s.Nome = LeggiStringaConDefault("Nome", s.Nome);
        s.Cognome = LeggiStringaConDefault("Cognome", s.Cognome);
        s.Classe = LeggiStringaConDefault("Classe", s.Classe).ToUpperInvariant();
        s.Disciplina = LeggiStringaConDefault("Disciplina", s.Disciplina);

        if (LeggiSiNo("Vuoi riselezionare le carenze rilevate? [S/N]: ", false))
            s.CarenzeRilevate = SelezionaCarenzeRilevate();

        s.Data = LeggiStringaConDefault("Data scheda", s.Data);
        s.Docente = LeggiStringaConDefault("Docente", s.Docente);

        if (LeggiSiNo("Vuoi riselezionare gli argomenti? [S/N]: ", false))
        {
            string anno = AnnoDaClasse(s.Classe);
            if (_argomentiPerAnno.ContainsKey(anno))
                s.Argomenti = SelezionaArgomenti(_argomentiPerAnno[anno]);
            else
                Console.WriteLine($"Nessun argomento disponibile per '{anno}'.");
        }

        if (LeggiSiNo("Vuoi modificare prova/tipologia accertamento? [S/N]: ", false))
        {
            s.ProvaAccertamento = LeggiSiNo("Prova di accertamento? [S/N]: ", s.ProvaAccertamento);
            s.TipologieAccertamento = SelezionaTipologie();
        }

        Console.WriteLine("Studente modificato.");
    }

    private static void EliminaStudente()
    {
        if (_studenti.Count == 0)
        {
            Console.WriteLine("Nessuno studente da eliminare.");
            return;
        }

        ElencaStudenti();
        int index = LeggiIntero("Numero studente da eliminare: ", 1, _studenti.Count) - 1;
        var s = _studenti[index];

        if (LeggiSiNo($"Confermi eliminazione di {s.Cognome} {s.Nome}? [S/N]: ", false))
        {
            _studenti.RemoveAt(index);
            Console.WriteLine("Studente eliminato.");
        }
    }

    private static void GeneraSchede()
    {
        if (_studenti.Count == 0)
        {
            Console.WriteLine("Nessuno studente inserito.");
            return;
        }

        Directory.CreateDirectory(OutputDir);

        foreach (var studente in _studenti)
        {
            string path = DocxService.GeneraScheda(studente, OutputDir);
            Console.WriteLine("Creato: " + path);
        }

        Console.WriteLine("Generazione completata.");
    }

    private static void SalvaStudenti()
    {
        JsonService.SalvaStudenti(StudentiPath, _studenti);
        Console.WriteLine($"Archivio salvato in {StudentiPath}");
    }

    private static void RicaricaStudenti()
    {
        _studenti = JsonService.CaricaStudentiSePresenti(StudentiPath);
        Console.WriteLine($"Archivio ricaricato. Studenti presenti: {_studenti.Count}");
    }

    private static void ElencaArgomentiDisponibili()
    {
        Console.WriteLine();
        Console.WriteLine("--- Argomenti disponibili ---");

        foreach (var anno in _argomentiPerAnno.Keys.OrderBy(x => x))
        {
            Console.WriteLine();
            Console.WriteLine(EtichettaAnno(anno).ToUpperInvariant());

            var lista = _argomentiPerAnno[anno];
            for (int i = 0; i < lista.Count; i++)
                Console.WriteLine($"{i + 1}. {lista[i].Titolo}");
        }
    }

    private static string SelezionaCarenzeRilevate()
    {
        Console.WriteLine();
        Console.WriteLine("--- Seleziona carenze rilevate ---");

        for (int i = 0; i < CarenzePredefinite.Count; i++)
            Console.WriteLine($"{i + 1}. {CarenzePredefinite[i]}");

        Console.WriteLine();
        Console.WriteLine("Inserisci i numeri separati da virgola, es. 1,3,5");
        Console.WriteLine("Scrivi T per selezionare tutte le carenze.");
        Console.WriteLine("Premi INVIO per inserire una carenza personalizzata.");
        Console.Write("Scelta: ");

        string input = (Console.ReadLine() ?? "").Trim();

        if (string.IsNullOrWhiteSpace(input))
            return LeggiStringaConDefault("Carenza personalizzata", "Metodo di studio inefficace.");

        if (input.Equals("T", StringComparison.OrdinalIgnoreCase))
            return string.Join(Environment.NewLine, CarenzePredefinite);

        List<string> selezionate = new();

        string[] parti = input.Split(
            ',',
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries
        );

        foreach (string parte in parti)
        {
            if (int.TryParse(parte, out int numero))
            {
                if (numero >= 1 && numero <= CarenzePredefinite.Count)
                {
                    string carenza = CarenzePredefinite[numero - 1];

                    if (!selezionate.Contains(carenza))
                        selezionate.Add(carenza);
                }
            }
        }

        if (selezionate.Count == 0)
        {
            Console.WriteLine("Nessuna carenza valida selezionata. Verrà usata una carenza generica.");
            return "Metodo di studio inefficace.";
        }

        return string.Join(Environment.NewLine, selezionate);
    }

    private static List<Argomento> SelezionaArgomenti(List<Argomento> disponibili)
    {
        Console.WriteLine();
        Console.WriteLine("--- Seleziona argomenti da inserire nella scheda ---");

        for (int i = 0; i < disponibili.Count; i++)
            Console.WriteLine($"{i + 1}. {disponibili[i].Titolo}");

        Console.WriteLine();
        Console.WriteLine("Inserisci i numeri separati da virgola, es. 1,3,5");
        Console.WriteLine("Scrivi T per selezionare tutti.");
        Console.Write("Scelta: ");

        string input = (Console.ReadLine() ?? "").Trim();

        if (input.Equals("T", StringComparison.OrdinalIgnoreCase))
            return disponibili.Select(a => a.Clone()).ToList();

        var selezionati = new List<Argomento>();
        var parti = input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var parte in parti)
        {
            if (int.TryParse(parte, out int numero) && numero >= 1 && numero <= disponibili.Count)
            {
                var arg = disponibili[numero - 1];
                if (!selezionati.Any(a => a.Titolo.Equals(arg.Titolo, StringComparison.OrdinalIgnoreCase)))
                    selezionati.Add(arg.Clone());
            }
        }

        return selezionati;
    }

    private static List<string> SelezionaTipologie()
    {
        var voci = new List<string> { "ORALE", "SCRITTO", "PRATICO" };
        var selezionate = new List<string>();

        Console.WriteLine();
        Console.WriteLine("--- Tipologia accertamento ---");

        for (int i = 0; i < voci.Count; i++)
            Console.WriteLine($"{i + 1}. {voci[i]}");

        Console.WriteLine("Inserisci i numeri separati da virgola, es. 2,3");
        Console.Write("Scelta: ");

        string input = Console.ReadLine() ?? "";
        var parti = input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var parte in parti)
        {
            if (int.TryParse(parte, out int numero) && numero >= 1 && numero <= voci.Count)
                selezionate.Add(voci[numero - 1]);
        }

        if (selezionate.Count == 0)
            selezionate.Add("SCRITTO");

        return selezionate;
    }

    private static string AnnoDaClasse(string classe)
    {
        classe = (classe ?? "").Trim();

        if (classe.StartsWith("1")) return "prime";
        if (classe.StartsWith("2")) return "seconde";
        if (classe.StartsWith("3")) return "terze";
        if (classe.StartsWith("4")) return "quarte";
        if (classe.StartsWith("5")) return "quinte";

        return "altre";
    }

    private static string EtichettaAnno(string anno)
    {
        return anno switch
        {
            "prime" => "Classi prime",
            "seconde" => "Classi seconde",
            "terze" => "Classi terze",
            "quarte" => "Classi quarte",
            "quinte" => "Classi quinte",
            _ => anno
        };
    }

    private static string LeggiStringa(string messaggio)
    {
        Console.Write(messaggio);
        return Console.ReadLine() ?? "";
    }

    private static string LeggiStringaObbligatoria(string messaggio)
    {
        while (true)
        {
            string valore = LeggiStringa(messaggio).Trim();
            if (!string.IsNullOrWhiteSpace(valore))
                return valore;

            Console.WriteLine("Valore obbligatorio.");
        }
    }

    private static string LeggiStringaConDefault(string etichetta, string valoreDefault)
    {
        Console.Write($"{etichetta} [{valoreDefault}]: ");
        string valore = Console.ReadLine() ?? "";
        return string.IsNullOrWhiteSpace(valore) ? valoreDefault : valore.Trim();
    }

    private static bool LeggiSiNo(string messaggio, bool defaultValue)
    {
        Console.Write(messaggio);
        string valore = (Console.ReadLine() ?? "").Trim().ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(valore))
            return defaultValue;

        return valore is "S" or "SI" or "Y" or "YES";
    }

    private static int LeggiIntero(string messaggio, int min, int max)
    {
        while (true)
        {
            Console.Write(messaggio);
            string input = Console.ReadLine() ?? "";

            if (int.TryParse(input, out int valore) && valore >= min && valore <= max)
                return valore;

            Console.WriteLine($"Inserisci un numero tra {min} e {max}.");
        }
    }
}
