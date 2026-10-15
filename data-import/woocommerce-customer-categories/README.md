# woocommerce-customer-categories

Confronto tra le categorie dei clienti B2B su WooCommerce e le categorie
cliente in ARCA (`Cd_CFGruppo1/2/3`).

## Problema

La tipologia di attività di ogni cliente B2B è registrata come ruolo utente su
WooCommerce e deve essere riportata nelle categorie dell'anagrafica `CF` di
ARCA. Le due fonti tendono a disallinearsi: clienti senza categoria, categorie
cambiate, ruoli rimossi.

## Soluzione

`compare_customer_categories.py` legge l'export utenti di WooCommerce, lo
confronta in sola lettura con ARCA e genera un Excel di controllo e uno script
SQL da rivedere prima dell'esecuzione.

1. Legge il CSV, normalizza gli ID (solo la parte numerica) e converte il ruolo
   nel codice categoria secondo `ROLE_MAPPING`. Esempio di configurazione:

   | Ruolo WooCommerce  | Codice |
   |--------------------|--------|
   | Ristorazione       | RIS    |
   | Negozio alimentari | ALI    |
   | Vending            | VEN    |
   | Ingrosso           | ING    |
   | Vendita online     | WEB    |
   | Altro              | OTH    |

   I codici devono esistere tra i `CFGruppo` di ARCA. I ruoli di sistema
   (amministratore, cliente, agente, shop manager) non vengono mappati.
2. Legge da `CF` i clienti con `xWooCommerce_ID_B2B` corrispondente, a blocchi
   di 500 ID (`bindparam(..., expanding=True)`).
3. Assegna a ogni utente un'azione:

   | Azione | Condizione |
   |---|---|
   | `AGGIUNGI` | categoria assente in ARCA, ruolo mappato nel CSV |
   | `AGGIORNA` | categoria in ARCA diversa da quella del CSV |
   | `RIMUOVI` | categoria presente in ARCA, ruolo assente nel CSV |
   | `NESSUNA_AZIONE` | dati allineati |
   | `DA_VERIFICARE_MANUALMENTE` | ruolo di sistema o non previsto |
   | `UTENTE_NON_TROVATO_IN_CF` | ID WooCommerce non presente in `CF` |

   La stessa categoria viene scritta su tutte e tre le colonne `Cd_CFGruppo1/2/3`.
4. Produce:
   - `confronto_categorie_CF.xlsx`, con i fogli Confronto, Riepilogo, Verifica
     Manuale e Utente Non Trovato (con i dati anagrafici dal CSV);
   - `update_categorie_CF.sql`, con un `UPDATE` commentato per ogni riga da
     aggiungere, aggiornare o rimuovere.

`verify_customer_categories.sql` mostra, dopo l'aggiornamento, tutte le colonne
di `CF` dei clienti collegati a WooCommerce con la descrizione delle tre
categorie.

## Utilizzo

1. Da WooCommerce (plugin All Export) esportare gli utenti B2B in CSV con
   separatore `;` e salvarlo come `woocommerce_ids.csv` nella cartella dello
   script. Colonne attese: `ID`, `First Name`, `Last Name`,
   `Società di fatturazione`, `User Email`, `User Role`, `User Registered`,
   `Referente commerciale`, `Email referente commerciale`. Sono indispensabili
   solo `ID` e `User Role`.
2. Eseguire:
   ```bash
   python compare_customer_categories.py
   ```
3. Controllare l'Excel: le righe in Verifica Manuale vanno valutate a mano, i
   clienti in Utente Non Trovato vanno creati in ARCA separatamente.
4. Rivedere ed eseguire `update_categorie_CF.sql` in SSMS.

La connessione usa le variabili `SERVER_DB`, `DB_ARCA`, `DB_USER`, `DB_PASSWORD`
del file `.env` alla radice del repository.

## Confronto con woocommerce-role-mapping

I due strumenti risolvono lo stesso problema con approcci complementari.

| Funzione | role-mapping | customer-categories |
|---|---|---|
| Input | Excel (ID, ruolo) | CSV dell'export utenti |
| Accesso al database | no, genera solo SQL | sì, in sola lettura |
| Mappatura dei ruoli | manuale con fallback fuzzy | tabella esatta |
| Confronto con le categorie in ARCA | no | sì |
| Rimozione delle categorie non più valide | no | sì |
| Segnalazione utenti assenti in `CF` | no | sì |
| Report | testo | Excel |
| Gestione di più righe `CF` con lo stesso ID WooCommerce | sì | no |
| Disattivazione del trigger `CF_atrg_brd` | sì | no |
| `UPDATE` set-based a blocchi | sì | no |

customer-categories è adatto a verificare e rivedere le differenze rispetto ad
ARCA; role-mapping ad applicare la mappatura in blocco, anche con anagrafiche
duplicate o ruoli scritti in modo non uniforme.
