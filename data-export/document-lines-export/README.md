# document-lines-export

Estrazione in Excel delle righe di un documento di ARCA Evolution.

## Problema

Clienti e colleghi chiedono spesso le righe di un ordine, DDT o fattura in un
formato modificabile, con colonne che le stampe standard non includono (EAN,
prezzo netto, prezzo consigliato).

## Soluzione

`export_document_lines.py` (GUI tkinter) costruisce dinamicamente la query su
`DORig` e `ARAlias` in base ai filtri e alle colonne scelte:

- filtri: anno, codice documento (es. `FTV`), numero documento e codice
  cliente, oppure il solo identificativo `Id_DOTes`, che ha la precedenza;
- colonne disponibili: codice articolo, EAN, descrizione, unità di misura,
  fattore di conversione, quantità, prezzo unitario, aliquota IVA, prezzo
  lordo, netto e consigliato (dal listino con revisione 1);
- l'opzione di corrispondenza esatta filtra sul numero documento (`NumeroDoc`);
  se i risultati provengono da più documenti viene mostrato un avviso con
  l'elenco.

## Utilizzo

```bash
python export_document_lines.py
```

Compilare i filtri, scegliere le colonne, premere Estrai Dati e indicare il
file Excel di destinazione. La connessione usa il `.env` alla radice del
repository.
