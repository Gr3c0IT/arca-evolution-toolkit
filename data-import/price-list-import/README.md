# price-list-import

Aggiornamento dei prezzi di un listino ARCA Evolution a partire da un file
Excel.

## Problema

I nuovi prezzi vengono preparati in un Excel con codice articolo e prezzo, ma
ARCA importa i listini solo nel proprio modello di esportazione. Riportare i
prezzi a mano nel modello richiede tempo ed è soggetto a errori.

## Soluzione

`price_list_builder.py` (GUI tkinter) abbina i codici articolo dei due file,
ignorando gli spazi, e scrive i nuovi prezzi nella colonna E del modello
esportato da ARCA. Accetta prezzi numerici o testuali (es. `€ 1,50`). Mostra
un'anteprima delle modifiche e apre il file temporaneo con il programma
predefinito (Windows, macOS, Linux) prima di salvare il risultato. Non richiede
connessione al database.

## Utilizzo

1. In ARCA: Listini → Listini Standard → Manutenzione → Esporta Listino,
   selezionando il listino da aggiornare (deve già esistere).
2. Preparare l'Excel dei nuovi prezzi: un solo foglio, nessuna intestazione,
   colonne codice articolo e prezzo.
3. Avviare `python price_list_builder.py` e selezionare i due file.
4. Visualizza Anteprima genera `temporaneo.xlsx` e lo apre per il controllo.
5. Esegui la modifica salva `Listino_pronto.xlsx` nella cartella del listino
   originale, pronto per l'import in ARCA.
