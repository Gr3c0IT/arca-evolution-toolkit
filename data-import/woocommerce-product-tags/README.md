# woocommerce-product-tags — Import dei tag prodotto

`generate_product_tags_sql.py` genera lo script SQL che popola la tabella
personalizzata `xARARTags_WooCommerce` (articolo ↔ tag WooCommerce) partendo da
un file Excel.

Lo script **non si collega al database**.

## Formato dell'Excel

- Riga 1: intestazioni.
- Colonna **A**: codice articolo (`Cd_AR`).
- Colonna **D**: ID del tag WooCommerce (`Id_xARTags_WooCommerce`), oppure vuota.

Gli ID ammessi si configurano in `VALORI_AMMESSI` in testa allo script.

## Utilizzo

```bash
python generate_product_tags_sql.py input.xlsx output.sql
```

Il file SQL generato:
- inserisce le righe valide in blocchi da 900, sotto il limite di 1000 righe
  per `VALUES` di SQL Server;
- lavora in una transazione con `TRY/CATCH`, con rollback automatico in caso
  di errore;
- riporta in testa il conteggio delle righe inserite, escluse (colonna D
  vuota) e anomale (valore non ammesso).
