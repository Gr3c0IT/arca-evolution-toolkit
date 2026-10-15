# supplier-invoices-export

Export massivo delle fatture elettroniche passive archiviate in ARCA Evolution,
in formato XML e PDF.

## Problema

Le fatture passive sono archiviate nel DMS di ARCA (`DmsDocument`) come blob
compressi, spesso firmati (`.p7m`). Estrarle una per una per un controllo o una
richiesta di documentazione su un conto contabile richiede molto tempo, e non
tutte includono un PDF.

## Soluzione

`export_supplier_invoices.ps1` seleziona le registrazioni contabili
(`CGMovT`/`CGMovR`) su un conto e in un intervallo di date di registrazione, poi
per ogni documento DMS con tag `FTE`:

1. decomprime il blob di ARCA (record ZIP con dati Deflate);
2. se il file è firmato, anche in Base64, estrae l'XML dalla busta CAdES;
3. salva `<conto>_<CodiceFornitore>_<NumFattura>_<DataReg>.xml`;
4. salva il PDF allegato all'XML, se presente, altrimenti lo genera
   trasformando l'XML con un foglio di stile XSL e stampandolo con Microsoft
   Edge in modalità headless;
5. scrive `_riepilogo.csv` con l'esito di ogni documento: numero, data e totale
   dall'XML, coerenza tra numero in ARCA e numero nell'XML, origine del PDF,
   errori, registrazioni senza XML.

Requisiti: Windows con PowerShell 5.1 o successivo; Microsoft Edge, solo per
generare i PDF mancanti; un foglio di stile per la fattura elettronica salvato
come `FoglioStile.xsl` accanto allo script (ad esempio quello pubblicato da
FatturaPA); variabili `SERVER_DB` e `DB_ARCA` (ed eventualmente `DB_USER`,
`DB_PASSWORD`) nel `.env` alla radice del repository o nell'ambiente.

## Utilizzo

```powershell
.\export_supplier_invoices.ps1 -Conto 00000000000
.\export_supplier_invoices.ps1 -Conto 00000000000 -DataInizio 2026-01-01 -DataFine 2026-12-31
```

Senza date viene usato l'anno in corso. I file vengono salvati in
`Fatture_<conto>_<inizio>_<fine>` accanto allo script.
