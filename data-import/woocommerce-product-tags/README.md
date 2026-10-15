# woocommerce-product-tags

Import dei tag prodotto WooCommerce nella tabella personalizzata
`xARARTags_WooCommerce` di ARCA.

## Problema

L'associazione tra articoli e tag WooCommerce viene preparata in Excel e va
caricata in ARCA, dove l'inserimento manuale di centinaia di righe non è
pratico.

## Soluzione

`generate_product_tags_sql.py` legge l'Excel e genera lo script di
inserimento, senza collegarsi al database:

- inserimento in blocchi da 900 righe, sotto il limite di 1000 righe per
  `VALUES` di SQL Server;
- transazione con `TRY/CATCH` e rollback in caso di errore;
- conteggio in testa al file delle righe inserite, escluse (tag vuoto) e
  anomale (tag non ammesso).

Formato dell'Excel: intestazioni nella riga 1, codice articolo (`Cd_AR`) in
colonna A, ID del tag (`Id_xARTags_WooCommerce`) in colonna D. Gli ID ammessi
si configurano in `VALORI_AMMESSI`.

## Utilizzo

```bash
python generate_product_tags_sql.py input.xlsx output.sql
```

Eseguire poi `output.sql` in SSMS.
