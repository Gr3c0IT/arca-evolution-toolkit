# product-catalog-export — Catalogo prodotti da ARCA + dati manuali

`product_catalog.py` costruisce il file Excel del catalogo prodotti unendo:

- i **dati presenti in ARCA**: descrizione, categoria, peso, EAN del pezzo e del
  cartone, pezzi per cartone, cartoni per strato e per pallet, immagine,
  prezzo di rivendita suggerito;
- i **dati manuali** che il gestionale non contiene: descrizione marketing e
  flag `TOP_SELLER`, `NOVITA`, `TREND`, `GLUTEN_FREE`, `PROTEICO`, `VEGANO`,
  `HALAL`. Sono conservati nell'archivio `catalogo_unificato.xlsx`.

## File

| File | Descrizione |
|------|-------------|
| `product_catalog.py` | script principale |
| `catalogo_unificato.xlsx` | archivio storico dei prodotti con i dati manuali |
| file dei codici (es. `codici.xlsx`) | codici articolo del nuovo catalogo, nella prima colonna |
| `da_compilare.xlsx` | generato: articoli con dati manuali mancanti, da far compilare |
| `catalogo_output.xlsx` | generato: catalogo finale |

## Flusso di lavoro

1. **Genera il catalogo**
   ```bash
   python product_catalog.py --genera codici.xlsx
   ```
   Legge i codici, recupera i dati da ARCA, aggiunge i dati manuali
   dall'archivio e produce `catalogo_output.xlsx` e `da_compilare.xlsx`.
2. **Compilazione manuale**: chi si occupa del marketing compila
   `da_compilare.xlsx` e lo restituisce.
3. **Aggiorna l'archivio**
   ```bash
   python product_catalog.py --aggiorna da_compilare.xlsx
   ```
   Aggiorna `catalogo_unificato.xlsx` con i dati manuali e con quelli correnti
   di ARCA.
4. Rieseguire `--genera` per ottenere il catalogo definitivo. Con
   `--genera catalogo_unificato.xlsx` si rigenera il catalogo completo.

I campi booleani accettano solo `TRUE`/`FALSE` (o `VERO`/`FALSO`); qualsiasi
altro valore è considerato mancante.

L'output è pensato per `automazioni-ufficio-python/pptx-catalog`, che lo
impagina in PowerPoint.

Richiede il file `.env` alla radice del repository.
