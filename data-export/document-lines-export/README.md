# document-lines-export — Estrazione delle righe di un documento

`export_document_lines.py` è una GUI (tkinter) per estrarre in Excel le righe
di un documento di ARCA Evolution (ordine, DDT, fattura, ...) scegliendo
quali colonne includere.

## Utilizzo

```bash
python export_document_lines.py
```

1. Compilare **Anno esercizio**, **Codice documento** (es. `FTV`), **Numero
   documento** e **Codice cliente**. In alternativa compilare solo
   l'**Identificativo** (`Id_DOTes`), che ha la precedenza sugli altri filtri.
2. Spuntare le colonne da esportare: codice articolo, EAN, descrizione,
   quantità, fattore di conversione, prezzo unitario, aliquota IVA, prezzo
   lordo, netto e consigliato (dal listino con revisione 1).
3. Premere **Estrai Dati** e scegliere dove salvare il file Excel.

L'opzione di corrispondenza esatta applica il filtro sul numero documento
(`NumeroDoc`). Se i risultati provengono da più documenti diversi, lo script
mostra un avviso con l'elenco.

Richiede il file `.env` alla radice del repository (`SERVER_DB`, `DB_ARCA`,
`DB_USER`, `DB_PASSWORD`).
