# woocommerce-role-mapping — Ruoli WooCommerce → categorie cliente ARCA

`generate_customer_group_updates.py` legge un Excel esportato da WooCommerce con
gli utenti B2B e il loro ruolo, associa ogni ruolo a una categoria cliente di
ARCA (`CFGruppo`) e genera lo script SQL che aggiorna `Cd_CFGruppo1/2/3` nella
tabella `CF`, usando come chiave `xWooCommerce_ID_B2B`.

Lo script **non si collega al database**: produce due file da controllare.

| File generato          | Contenuto |
|------------------------|-----------|
| `cf_update_report.txt` | report pre-esecuzione, riga per riga, con l'esito di ogni match |
| `cf_update.sql`        | script idempotente da eseguire in SSMS |

## Utilizzo

1. Da WooCommerce (plugin *All Export*) esportare gli utenti B2B con le colonne
   **ID** e **Role**, rinominandole in `xWooCommerce_ID_B2B` e `Ruolo`.
2. Eseguire:
   ```bash
   python generate_customer_group_updates.py percorso/del/file.xlsx
   ```
   In alternativa, impostare `EXCEL_FILE` in testa allo script.
3. Leggere `cf_update_report.txt`:
   - ✅ **OK**: match trovato (`[MANUALE]` dalla mappatura, `[FUZZY x%]` per similarità);
   - ➖ **ESCLUSO**: ruolo di sistema ignorato (amministratore, cliente, ...);
   - ⚠️ **DUPLICATO**: ID presente più volte nell'Excel;
   - ❌ **NO MATCH**: ruolo non riconosciuto, che richiede una mappatura manuale.
4. Eseguire `cf_update.sql` in SSMS.

## Come è fatto lo script SQL

1. Crea `#id_duplicati` con gli ID WooCommerce presenti su più righe di `CF`.
2. Disabilita il trigger CLR `CF_atrg_brd` di ARCA, aggiorna a blocchi di 1000
   solo le righe con categoria diversa o NULL e riabilita il trigger.
3. Elenca i clienti con ID duplicato.
4. Per i duplicati, copia la categoria dalla riga già valorizzata alle righe
   "sorelle" che ne sono prive.
5. Mostra i duplicati ancora senza categoria, da sistemare a mano in ARCA.

## Parametri configurabili

`EXCEL_FILE`, `SHEET_NAME`, `COL_ID`, `COL_RUOLO`, `OUTPUT_SQL`, `OUTPUT_REPORT`,
`FUZZY_THRESHOLD` (default 0.60), `RUOLI_ESCLUSI`, `CF_GRUPPI` (codice →
descrizione delle categorie ARCA), `MAPPING_MANUALE` (ruolo normalizzato →
codice). La mappatura manuale ha sempre la priorità sul fuzzy matching
(`difflib.SequenceMatcher`).

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
