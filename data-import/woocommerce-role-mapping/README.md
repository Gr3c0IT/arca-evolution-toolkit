# woocommerce-role-mapping

Generazione degli aggiornamenti delle categorie cliente in ARCA a partire dai
ruoli utente di WooCommerce.

## Problema

I ruoli degli utenti B2B su WooCommerce vanno riportati come categorie
(`Cd_CFGruppo1/2/3`) sull'anagrafica `CF` di ARCA. I nomi dei ruoli non sempre
coincidono con le descrizioni delle categorie e in `CF` possono esistere più
righe con lo stesso ID WooCommerce.

## Soluzione

`generate_customer_group_updates.py` legge un Excel con ID e ruolo, associa ogni
ruolo a una categoria e genera uno script SQL idempotente. Non si collega al
database.

- Associazione dei ruoli: prima il dizionario `MAPPING_MANUALE` (confronto
  esatto dopo normalizzazione: minuscole, separatori `/ - – _` convertiti in
  spazio), poi, in mancanza, un fuzzy matching con `difflib.SequenceMatcher`
  sulle descrizioni in `CF_GRUPPI`, con soglia `FUZZY_THRESHOLD` (default 0,60).
- Ruoli esclusi: quelli in `RUOLI_ESCLUSI` (amministratore, cliente, ecc.).
- Output:
  - `cf_update_report.txt`: esito di ogni riga (OK con metodo manuale o fuzzy,
    ESCLUSO, DUPLICATO, NO MATCH) e riepilogo;
  - `cf_update.sql`, strutturato in cinque passi:
    1. tabella temporanea `#id_duplicati` con gli ID WooCommerce presenti su
       più righe di `CF`;
    2. disattivazione del trigger CLR `CF_atrg_brd`, `UPDATE` a blocchi di
       1000 righe limitati alle categorie diverse o nulle, riattivazione del
       trigger;
    3. elenco dei clienti con ID duplicato;
    4. per i duplicati, copia della categoria dalla riga già valorizzata alle
       altre righe con lo stesso ID;
    5. elenco dei duplicati ancora senza categoria.

Parametri configurabili in testa allo script: `EXCEL_FILE`, `SHEET_NAME`,
`COL_ID`, `COL_RUOLO`, `OUTPUT_SQL`, `OUTPUT_REPORT`, `FUZZY_THRESHOLD`,
`RUOLI_ESCLUSI`, `CF_GRUPPI`, `MAPPING_MANUALE`.

## Utilizzo

1. Esportare da WooCommerce gli utenti B2B con le colonne ID e ruolo,
   rinominate in `xWooCommerce_ID_B2B` e `Ruolo`.
2. Eseguire:
   ```bash
   python generate_customer_group_updates.py percorso/del/file.xlsx
   ```
3. Controllare in `cf_update_report.txt` le righe NO MATCH; per correggerle,
   aggiungere il ruolo a `MAPPING_MANUALE`.
4. Eseguire `cf_update.sql` in SSMS.

Il confronto con lo strumento `woocommerce-customer-categories` è descritto nel
README di quella cartella.
