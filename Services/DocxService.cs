using CarenzeGenerator.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CarenzeGenerator.Services;

public static class DocxService
{
    public static string GeneraScheda(Studente studente, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        string fileName = $"{Pulisci(studente.Classe)}_{Pulisci(studente.Cognome)}_{Pulisci(studente.Nome)}_carenze.docx";
        string path = Path.Combine(outputDir, fileName);

        using WordprocessingDocument document =
            WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);

        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();

        Body body = new Body();

        //AggiungiParagrafo(body, "Scheda segnalazione carenze da colmare", false, false, null, 20);
        AggiungiParagrafo(body, "ANNO SCOLASTICO 2025 / 2026 - SCRUTINIO FINALE", true, false, JustificationValues.Center, 24);
        AggiungiParagrafo(body, "SCHEDA di SEGNALAZIONE delle CARENZE da COLMARE", true, true, JustificationValues.Center, 26);
        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, $"ALUNNO/A {studente.Nome} {studente.Cognome}    CLASSE {studente.Classe}");
        AggiungiParagrafo(body, $"DISCIPLINA {studente.Disciplina}");
        AggiungiParagrafo(body, "");

        TitoloSezione(body, "CARENZE RILEVATE");

        string carenze = string.IsNullOrWhiteSpace(studente.CarenzeRilevate)
            ? "Metodo di studio inefficace."
            : studente.CarenzeRilevate;

        foreach (string riga in carenze.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries))
            AggiungiParagrafo(body, riga.Trim());

        AggiungiParagrafo(body, "");

        TitoloSezione(body, "INDICAZIONI DI LAVORO PER IL RECUPERO DELLE CARENZE");

        AggiungiParagrafo(body, "ARGOMENTI", true);

        if (studente.Argomenti.Count == 0)
        {
            AggiungiParagrafo(body, "• Nessun argomento selezionato.");
        }
        else
        {
            foreach (var argomento in studente.Argomenti)
                AggiungiParagrafo(body, "• " + argomento.Titolo);
        }

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "CONOSCENZE DA RECUPERARE", true);

        foreach (var argomento in studente.Argomenti)
        {
            AggiungiParagrafo(body, argomento.Titolo, true);

            if (argomento.Conoscenze.Count == 0)
            {
                AggiungiParagrafo(body, "• Conoscenze non specificate.");
            }
            else
            {
                foreach (var conoscenza in argomento.Conoscenze)
                    AggiungiParagrafo(body, "• " + conoscenza);
            }
        }

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "ABILITA’ DA RAGGIUNGERE", true);

        foreach (var argomento in studente.Argomenti)
        {
            AggiungiParagrafo(body, argomento.Titolo, true);

            if (argomento.Abilita.Count == 0)
            {
                AggiungiParagrafo(body, "• Abilità non specificate.");
            }
            else
            {
                foreach (var abilita in argomento.Abilita)
                    AggiungiParagrafo(body, "• " + abilita);
            }
        }

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "INDICAZIONI OPERATIVE DI LAVORO", true);
        AggiungiParagrafo(
            body,
            "Al fine di colmare le lacune evidenziate, si raccomanda lo studio dei contenuti teorici e la ripetizione guidata degli esercizi disponibili sul sito cauzzi.prof (https://cauzzi.prof), con particolare attenzione agli argomenti oggetto di recupero."
        );

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "PROVA DI ACCERTAMENTO", true);
        AggiungiParagrafo(body, studente.ProvaAccertamento ? "☑ SI    ☐ NO" : "☐ SI    ☑ NO");

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "TIPOLOGIA DELL’ACCERTAMENTO", true);
        AggiungiParagrafo(body, $"{Check(studente, "ORALE")} ORALE");
        AggiungiParagrafo(body, $"{Check(studente, "SCRITTO")} SCRITTO");
        AggiungiParagrafo(body, $"{Check(studente, "PRATICO")} PRATICO");

        AggiungiParagrafo(body, "");

        string data = string.IsNullOrWhiteSpace(studente.Data)
            ? "____________"
            : studente.Data;

        string docente = string.IsNullOrWhiteSpace(studente.Docente)
            ? "__________________________"
            : studente.Docente;

        AggiungiParagrafo(body, $"DATA: {data}        IL DOCENTE {docente}");

        SectionProperties sectionProperties = new SectionProperties(
            new PageMargin
            {
                Top = 720,
                Bottom = 720,
                Left = 720,
                Right = 720
            }
        );

        body.Append(sectionProperties);

        mainPart.Document.Append(body);
        mainPart.Document.Save();

        return path;
    }

    private static void TitoloSezione(Body body, string testo)
    {
        AggiungiParagrafo(body, testo, true, false, JustificationValues.Center, 24);
    }

    private static void AggiungiParagrafo(
        Body body,
        string testo,
        bool grassetto = false,
        bool sottolineato = false,
        JustificationValues? allineamento = null,
        int fontSize = 22)
    {
        RunProperties runProperties = new RunProperties();

        if (grassetto)
            runProperties.Append(new Bold());

        if (sottolineato)
            runProperties.Append(new Underline { Val = UnderlineValues.Single });

        runProperties.Append(new FontSize { Val = fontSize.ToString() });

        Run run = new Run();
        run.Append(runProperties);
        run.Append(new Text(testo) { Space = SpaceProcessingModeValues.Preserve });

        ParagraphProperties paragraphProperties = new ParagraphProperties();

        paragraphProperties.Append(
            new Justification
            {
                Val = allineamento ?? JustificationValues.Left
            }
        );

        Paragraph paragraph = new Paragraph();
        paragraph.Append(paragraphProperties);
        paragraph.Append(run);

        body.Append(paragraph);
    }

    private static string Check(Studente s, string voce)
    {
        return s.TipologieAccertamento.Any(t =>
            t.Equals(voce, StringComparison.OrdinalIgnoreCase))
            ? "☑"
            : "☐";
    }

    private static string Pulisci(string testo)
    {
        if (string.IsNullOrWhiteSpace(testo))
            return "ND";

        foreach (char c in Path.GetInvalidFileNameChars())
            testo = testo.Replace(c, '_');

        return testo.Trim().Replace(" ", "_").ToUpperInvariant();
    }
}
