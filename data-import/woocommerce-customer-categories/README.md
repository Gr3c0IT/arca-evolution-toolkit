# woocommerce-customer-categories — Categorie cliente da WooCommerce B2B ad ARCA

Questo processo confronta gli utenti del sito **WooCommerce B2B** con l'anagrafica clienti (`CF`)
sul database ARCA, individua le categorie commerciali mancanti/diverse/da
rimuovere, e produce un Excel di controllo + uno script SQL pronto da rivedere ed eseguire
manualmente. **Lo script Python è di sola lettura: non scrive mai sul database.**

---

## 1. Estrazione del file da WooCommerce B2B

1. Accedere al pannello di amministrazione del sito B2B (`/wp-admin`)
2. Nel menu laterale andare su **All Export → Manage Exports**
3. Cliccare sull'esportazione già configurata: (es. **"Export B2B Customers"**)
4. Lanciare l'export e scaricare il CSV → salvarlo come `woocommerce_ids.csv` nella cartella
   dello script.

### Se l'esportazione non esiste

Bisogna crearne una nuova (tipo **Customers**) con la stessa struttura di colonne, in questo
ordine:

| Colonna |
|---|
| ID |
| First Name |
| Last Name |
| Società di fatturazione |
| User Email |
| User Role |
| User Registered |
| Data primo ordine (con stato completato) |
| Referente commerciale |
| Email referente commerciale |

> Ai fini dello script servono realmente solo **ID** e **User Role** (più i dati anagrafici
> per il foglio "Utente Non Trovato"), ma è preferibile ricreare l'export identico per
> mantenere coerenza con le esportazioni future.

Il file va salvato in formato **CSV con separatore `;`** (è l'export standard di WooCommerce/
All Export in italiano) — lo script è già configurato per leggerlo così.

---

## 2. Cosa fa lo script

File: `compare_customer_categories.py`.

1. Legge `woocommerce_ids.csv`
2. Pulisce gli ID (rimuove spazi/caratteri indesiderati, tiene solo la parte numerica)
3. Mappa la colonna **User Role** del CSV nel codice categoria ARCA (`Cd_CFGruppo1`), secondo
   questa tabella:

   | User Role (WooCommerce) | Cd_CFGruppo1 |
   |---|---|
   | Ristorazione | RIS |
   | Negozio alimentari | ALI |
   | Vending | VEN |
   | Ingrosso | ING |
   | Vendita online | WEB |
   | Altro | OTH |

   La tabella è un esempio: ruoli e codici si configurano nel dizionario `ROLE_MAPPING`
   dello script e devono corrispondere ai codici `CFGruppo` presenti in ARCA.

   I ruoli **Agente/rappresentante, Amministratore, Cliente, Shop manager plus** sono ruoli
   di sistema WooCommerce (non tipologie di attività commerciale) e **non vengono mappati
   automaticamente** — finiscono nel foglio "Verifica Manuale".

4. Si connette al DB (parametri dal file `.env`) e legge dalla tabella `CF` i record
   il cui `xWooCommerce_ID_B2B` corrisponde a uno degli ID del CSV, prendendo `Cd_CF`,
   `Descrizione`, `Cd_CFGruppo1`, `Cd_CFGruppo2`, `Cd_CFGruppo3`.
5. Confronta CSV vs DB per ogni utente e assegna un'azione:

   | Azione | Quando |
   |---|---|
   | **AGGIUNGI** | DB senza categoria, CSV ha un ruolo mappato |
   | **AGGIORNA** | DB ha una categoria diversa da quella del CSV |
   | **RIMUOVI** | DB ha una categoria ma CSV non ha ruolo/ruolo vuoto |
   | **NESSUNA_AZIONE** | già allineati |
   | **DA_VERIFICARE_MANUALMENTE** | ruolo CSV di sistema o sconosciuto (non mappabile) |
   | **UTENTE_NON_TROVATO_IN_CF** | l'ID WooCommerce non esiste in `CF` |

   Quando si aggiunge/aggiorna una categoria, viene impostata **la stessa su tutte e 3** le
   colonne `Cd_CFGruppo1`, `Cd_CFGruppo2`, `Cd_CFGruppo3` (obbligatorie in ARCA, ma senza
   sottocategorie reali al momento usiamo lo stesso valore su tutte).

6. Produce due file di output:
   - **`confronto_categorie_CF.xlsx`** con 4 fogli:
     - `Confronto`: tutte le righe con l'azione risultante
     - `Riepilogo`: conteggio per tipo di azione
     - `Verifica Manuale`: utenti con ruolo CSV non mappabile, con motivo esplicito
     - `Utente Non Trovato`: utenti presenti nel CSV ma assenti in `CF`, con i dati
       anagrafici completi presi dal CSV
   - **`update_categorie_CF.sql`**: le query `UPDATE` pronte (una per ogni AGGIUNGI/
     AGGIORNA/RIMUOVI), commentate con nome cliente, ID WooCommerce e stato precedente

---

## 3. Come eseguire lo script

### Requisiti (una tantum)

```powershell
pip install -r requirements.txt   # dalla radice del repository
```

### Esecuzione

1. Copiare `woocommerce_ids.csv` nella stessa cartella dello script
2. Lanciare:

```powershell
python compare_customer_categories.py
```

3. Verranno generati nella stessa cartella:
   - `confronto_categorie_CF.xlsx`
   - `update_categorie_CF.sql`

### Cosa controllare prima di eseguire l'SQL

1. Aprire l'Excel e controllare in particolare:
   - **Verifica Manuale**: decidere manualmente se assegnare una categoria a questi utenti
     (es. "Shop manager plus" potrebbe in realtà essere un negozio vero, va valutato caso
     per caso) — lo script non genera SQL per queste righe.
   - **Utente Non Trovato**: sono clienti WooCommerce che non esistono ancora come `CF` in
     ARCA — non è oggetto di questo script (creazione nuovo cliente), va gestito a parte.
   - **Riepilogo**: numero di righe per azione, come controllo rapido di coerenza.
2. Aprire `update_categorie_CF.sql`, leggere i commenti riga per riga (nome cliente/WC ID/
   stato precedente) e verificare che abbia senso.
3. **Solo dopo revisione**, eseguire manualmente lo script SQL sul database (lo script Python
   non esegue mai scritture in autonomia).

---

## 4. Note tecniche

- Connessione DB: variabili `SERVER_DB`, `DB_ARCA`, `DB_USER`, `DB_PASSWORD` dal file `.env`
  (con `DB_USER` vuoto si usa l'autenticazione Windows integrata), driver `{SQL Server}`.
- Le colonne `Cd_CFGruppo1/2/3` sono `char(3)` nullable in `CF` — i valori letti dal DB
  vengono ripuliti dagli spazi di padding prima del confronto.
- `xWooCommerce_ID_B2B` è `int` in `CF`: gli ID letti dal CSV vengono convertiti a intero
  (estraendo la prima sequenza numerica dalla cella, per tollerare spazi/caratteri accidentali
  prima o dopo il numero).
- Query con `IN` su liste di ID gestite tramite `bindparam(..., expanding=True)` e suddivise
  in chunk da 500 per evitare limiti sui parametri ODBC.

## 5. Verifica dopo l'aggiornamento

`verify_customer_categories.sql` mostra tutte le colonne dell'anagrafica `CF` dei
clienti collegati a WooCommerce, affiancando ai codici `Cd_CFGruppo1/2/3` la
descrizione della categoria: utile per un controllo a campione dopo aver
eseguito gli UPDATE.

---

## Differenze con lo strumento gemello

`woocommerce-role-mapping/` e `woocommerce-customer-categories/` allineano entrambi le
categorie cliente (`Cd_CFGruppo1/2/3`) ai ruoli WooCommerce, ma nessuno dei due copre
tutto quello che fa l'altro:

| Funzione | role-mapping | customer-categories |
|---|:---:|:---:|
| Input | Excel (ID + ruolo) | CSV completo dell'export utenti |
| Connessione al database | no (genera solo SQL) | sì, in sola lettura |
| Mappatura ruoli | manuale + **fuzzy matching** | solo tabella esatta |
| Confronto con le categorie già presenti in ARCA | no (salta solo le righe già corrette, lato SQL) | **sì**: AGGIUNGI / AGGIORNA / NESSUNA_AZIONE |
| Rimozione della categoria se il ruolo non c'è più | no | **sì** (RIMUOVI) |
| Utenti WooCommerce non presenti in `CF` | no | **sì**, con dati anagrafici |
| Report | testo (`cf_update_report.txt`) | Excel a 4 fogli |
| Righe `CF` con lo stesso ID WooCommerce | **sì**: escluse e poi completate copiando la categoria | no |
| Disattivazione del trigger CLR `CF_atrg_brd` | **sì** | no |
| UPDATE set-based a blocchi da 1000 | **sì** | no (un UPDATE per cliente) |

In pratica: **customer-categories** serve per capire cosa è cambiato rispetto ad ARCA
e rivedere le differenze; **role-mapping** serve per applicare in blocco la
mappatura su molti clienti, anche con anagrafiche duplicate o ruoli scritti in modo
non uniforme.
