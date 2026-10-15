# product-catalog-export

Generazione del file dati del catalogo prodotti unendo le informazioni di ARCA
Evolution con dati compilati manualmente.

## Problema

Il catalogo richiede sia dati presenti nel gestionale (descrizioni, EAN,
imballi, prezzi) sia informazioni di marketing che ARCA non contiene
(descrizione commerciale, flag come novità o senza glutine). Raccoglierli a
ogni nuova edizione del catalogo era un lavoro manuale ripetuto.

## Soluzione

`product_catalog.py` mantiene un archivio dei dati manuali
(`catalogo_unificato.xlsx`) e lo combina con i dati estratti da ARCA:

- da ARCA: descrizione, categoria, peso, EAN del pezzo e del cartone, pezzi per
  cartone, cartoni per strato e per pallet, immagine, prezzo di rivendita
  suggerito;
- dall'archivio: descrizione commerciale e flag `TOP_SELLER`, `NOVITA`,
  `TREND`, `GLUTEN_FREE`, `PROTEICO`, `VEGANO`, `HALAL`.

I flag accettano solo `TRUE`/`FALSE` (o `VERO`/`FALSO`); altri valori sono
considerati mancanti. Gli articoli con dati manuali mancanti vengono esportati
in `da_compilare.xlsx`.

L'output ha la struttura attesa da `automazioni-ufficio-python/pptx-catalog`,
che lo impagina in PowerPoint.

## Utilizzo

1. Generare il catalogo dai codici articolo (prima colonna del file):
   ```bash
   python product_catalog.py --genera codici.xlsx
   ```
   Produce `catalogo_output.xlsx` e `da_compilare.xlsx`.
2. Far compilare `da_compilare.xlsx` e aggiornare l'archivio:
   ```bash
   python product_catalog.py --aggiorna da_compilare.xlsx
   ```
   L'archivio viene aggiornato con i dati manuali e con quelli correnti di ARCA.
3. Rieseguire `--genera` per il catalogo definitivo. Con
   `--genera catalogo_unificato.xlsx` si rigenera il catalogo completo.

La connessione usa il `.env` alla radice del repository.
