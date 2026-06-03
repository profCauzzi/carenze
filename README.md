# CarenzeGenerator OpenXML

Programma console C#/.NET 8 per generare schede di carenza in formato DOCX.

## Versione

Questa versione usa:

- .NET 8
- DocumentFormat.OpenXml 3.5.1
- nessuna dipendenza Xceed / DocX commerciale

## Funzioni

- caricamento argomenti da `Data/argomenti_per_anno_ottimizzato.json`
- inserimento studenti da console
- selezione carenze rilevate da elenco predefinito
- `T` per selezionare tutte le carenze
- selezione argomenti in base all'anno della classe
- salvataggio archivio studenti in `Data/studenti.json`
- generazione file `.docx` in `output/`
- formato output: `CLASSE_COGNOME_NOME_carenze.docx`

## Avvio

```bash
dotnet restore
dotnet run
```

## Build

```bash
dotnet clean
dotnet restore
dotnet build
```
