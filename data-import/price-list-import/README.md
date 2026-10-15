# price-list-import — Aggiornamento di un listino ARCA da Excel

`price_list_builder.py` è una piccola GUI (tkinter) che prende il modello di
listino esportato da ARCA Evolution e ci scrive i nuovi prezzi presi da un
Excel. Il file risultante si reimporta in ARCA.

## Preparazione dei file

1. **Listino originale da ARCA**: *Listini → Listini Standard → Manutenzione →
   Esporta Listino*, poi selezionare il listino ed esportarlo. Se il listino
   non esiste, va prima creato in ARCA.
2. **Listino con i nuovi prezzi**: un Excel con **un solo foglio**, **senza
   intestazione** e con due colonne: codice articolo e prezzo. Il prezzo può
   essere numerico o testo, ad esempio `€ 1,50`.

## Utilizzo

```bash
python price_list_builder.py
```

1. Selezionare il listino originale e il listino Excel.
2. Premere **Visualizza Anteprima**: lo script abbina i codici, ignorando gli
   spazi, e scrive i prezzi nella colonna E. Poi apre il file temporaneo
   (`temporaneo.xlsx`) con il programma predefinito, su Windows, macOS o Linux.
3. Se è tutto corretto, premere **Esegui la modifica**: viene creato
   `Listino_pronto.xlsx` nella cartella del listino originale, da importare in
   ARCA.

Non richiede connessione al database.
