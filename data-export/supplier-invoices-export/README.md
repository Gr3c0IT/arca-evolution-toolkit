# supplier-invoices-export — Export massivo delle fatture passive (XML + PDF)

`export_supplier_invoices.ps1` estrae dal DMS di ARCA Evolution
(`DmsDocument`) le fatture elettroniche passive registrate in contabilità su
un certo **conto** e in un intervallo di **date di registrazione**. Per ognuna
salva l'XML e il PDF.

## Come funziona

1. Cerca le registrazioni contabili (`CGMovT`/`CGMovR`) sul conto indicato e
   i documenti DMS collegati con tag `FTE`.
2. Decomprime il blob di ARCA (record ZIP + Deflate) e, se il file è firmato
   (`.p7m`, anche in Base64), estrae l'XML dalla busta CAdES.
3. Salva `<conto>_<CodiceFornitore>_<NumFattura>_<DataReg>.xml`.
4. Per il PDF:
   - se l'XML contiene un PDF allegato, salva quello;
   - altrimenti genera il PDF trasformando l'XML con un foglio di stile XSL
     (`FoglioStile.xsl`) e stampandolo con Microsoft Edge in modalità headless.
5. Scrive `_riepilogo.csv` con l'esito di ogni documento: numero e data
   dall'XML, coerenza tra numero ARCA e numero XML, origine del PDF, errori.
   Include anche le registrazioni senza XML.

## Requisiti

- Windows con PowerShell 5.1 o superiore.
- Microsoft Edge, solo per generare i PDF mancanti.
- Un foglio di stile per la fattura elettronica salvato come `FoglioStile.xsl`
  accanto allo script (ad esempio il foglio di stile pubblicato da
  FatturaPA/Agenzia delle Entrate).
- Variabili `SERVER_DB`, `DB_ARCA` (e facoltativamente `DB_USER`/`DB_PASSWORD`)
  nel file `.env` alla radice del repository o nell'ambiente.

## Utilizzo

```powershell
.\export_supplier_invoices.ps1 -Conto 00000000000
.\export_supplier_invoices.ps1 -Conto 00000000000 -DataInizio 2026-01-01 -DataFine 2026-12-31
```

Senza date si usa l'anno in corso. I file vengono salvati nella cartella
`Fatture_<conto>_<inizio>_<fine>` accanto allo script.
