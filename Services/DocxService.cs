using CarenzeGenerator.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CarenzeGenerator.Services;

public static class DocxService
{
    // Template Word vuoto con il solo header già impostato.
    // Salva il file in: Assets/template_header.docx
    private static readonly string TemplatePath = Path.Combine("Assets", "template_header.docx");

    // Firma opzionale come immagine.
    // Salva il file in: Assets/firma.png
    private static readonly string FirmaImagePath = Path.Combine("Assets", "firma.png");

    public static string GeneraScheda(Studente studente, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        if (!File.Exists(TemplatePath))
        {
            throw new FileNotFoundException(
                $"Template Word non trovato. Crea un file Word vuoto con il solo header e salvalo in: {TemplatePath}"
            );
        }

        string fileName = $"{Pulisci(studente.Classe)}_{Pulisci(studente.Cognome)}_{Pulisci(studente.Nome)}_carenze.docx";
        string path = Path.Combine(outputDir, fileName);

        File.Copy(TemplatePath, path, true);

        using WordprocessingDocument document =
            WordprocessingDocument.Open(path, true);

        MainDocumentPart mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("Il template non contiene un MainDocumentPart valido.");

        mainPart.Document ??= new Document();

        Body body = mainPart.Document.Body ?? new Body();

        if (mainPart.Document.Body == null)
            mainPart.Document.Append(body);

        PulisciCorpoTemplate(body);

        AggiungiParagrafo(body, "\n\n", false, false, null, 20);
        AggiungiParagrafo(body, "ANNO SCOLASTICO 2025 / 2026 - SCRUTINIO FINALE", true, false, JustificationValues.Center, 24);
        AggiungiParagrafo(body, "SCHEDA di SEGNALAZIONE delle CARENZE da COLMARE", true, true, JustificationValues.Center, 26);
        AggiungiParagrafo(body, "");
        AggiungiParagrafo(body, "\n\n", false, false, null, 20);

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
            "Prima di affrontare la prova, si consiglia di ripassare la parte teorica e ripetere lo svolgimento degli esercizi presenti sulla piattaforma Moodle."
        );

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "PROVA DI ACCERTAMENTO", true);
        //AggiungiParagrafo(body, studente.ProvaAccertamento ? "☑ SI    ☐ NO" : "☐ SI    ☑ NO");
        AggiungiParagrafo(body, studente.ProvaAccertamento ? "[X] SI    [ ] NO" : "[ ] SI    [X] NO");

        AggiungiParagrafo(body, "");

        AggiungiParagrafo(body, "TIPOLOGIA DELL’ACCERTAMENTO", true);
        AggiungiParagrafo(body, $"{Check(studente, "ORALE")} ORALE");
        AggiungiParagrafo(body, $"{Check(studente, "SCRITTO")} SCRITTO");
        AggiungiParagrafo(body, $"{Check(studente, "PRATICO")} PRATICO");

        AggiungiParagrafo(body, "");

        string data = string.IsNullOrWhiteSpace(studente.Data)
            ? "____________"
            : studente.Data;

        AggiungiParagrafo(body, $"DATA: {data}        IL DOCENTE");

        AggiungiImmagine(
            mainPart,
            body,
            FirmaImagePath,
            CmToEmu(6.5),
            CmToEmu(1.8),
            JustificationValues.Right
        );

        AssicuraSectionProperties(body);

        mainPart.Document.Save();

        return path;
    }

    private static void PulisciCorpoTemplate(Body body)
    {
        // Conserva SectionProperties perché contengono anche i riferimenti a header/footer.
        SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault()?.CloneNode(true) as SectionProperties;

        body.RemoveAllChildren();

        if (sectionProperties != null)
            body.Append(sectionProperties);
    }

    private static void AssicuraSectionProperties(Body body)
    {
        SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();

        if (sectionProperties == null)
        {
            sectionProperties = new SectionProperties();
            body.Append(sectionProperties);
        }

        if (!sectionProperties.Elements<PageMargin>().Any())
        {
            sectionProperties.Append(
                new PageMargin
                {
                    Top = 720,
                    Bottom = 720,
                    Left = 720,
                    Right = 720
                }
            );
        }
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

        InserisciPrimaDiSectionProperties(body, paragraph);
    }

    private static void AggiungiImmagine(
        MainDocumentPart mainPart,
        Body body,
        string imagePath,
        long widthEmu,
        long heightEmu,
        JustificationValues? allineamento = null)
    {
        if (!File.Exists(imagePath))
            return;

        PartTypeInfo imageType = GetImagePartType(imagePath);
        ImagePart imagePart = mainPart.AddImagePart(imageType);

        using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            imagePart.FeedData(stream);
        }

        string relationshipId = mainPart.GetIdOfPart(imagePart);

        uint imageId = (uint)new Random().Next(1, int.MaxValue);

        var drawing =
            new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = imageId,
                        Name = Path.GetFileName(imagePath)
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    {
                                        Id = imageId,
                                        Name = Path.GetFileName(imagePath)
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                }
            );

        Paragraph paragraph = new Paragraph(
            new ParagraphProperties(
                new Justification { Val = allineamento ?? JustificationValues.Center }
            ),
            new Run(drawing)
        );

        InserisciPrimaDiSectionProperties(body, paragraph);
    }

    private static void InserisciPrimaDiSectionProperties(Body body, OpenXmlElement element)
    {
        SectionProperties? sectionProperties = body.Elements<SectionProperties>().LastOrDefault();

        if (sectionProperties != null)
            body.InsertBefore(element, sectionProperties);
        else
            body.Append(element);
    }

    private static PartTypeInfo GetImagePartType(string imagePath)
    {
        string ext = Path.GetExtension(imagePath).ToLowerInvariant();

        return ext switch
        {
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };
    }

    private static long CmToEmu(double cm)
    {
        return (long)(cm * 360000);
    }

    /*
    private static string Check(Studente s, string voce)
    {
        return s.TipologieAccertamento.Any(t =>
            t.Equals(voce, StringComparison.OrdinalIgnoreCase))
            ? "☑"
            : "☐";
    }
    */
    
    private static string Check(Studente s, string voce)
    {
        return s.TipologieAccertamento.Any(t =>
            t.Equals(voce, StringComparison.OrdinalIgnoreCase))
            ? "[X]"
            : "[ ]";
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
