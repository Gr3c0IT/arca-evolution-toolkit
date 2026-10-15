# queries

Query T-SQL da eseguire in SQL Server Management Studio sul database di ARCA
Evolution. I parametri (date, codici) sono dichiarati con `DECLARE` all'inizio
di ogni file.

| File | Cosa restituisce | Parametri |
|------|------------------|-----------|
| `article_sales_traceability.sql` | Ogni movimento di vendita di un articolo con documento, cliente, contatti, quantità (pezzi e cartoni) e indirizzo di consegna, anche dalla destinazione diversa o dai dati di spedizione. Utile per richiami e richieste di tracciabilità. | `@Cd_AR`, `@DallaData` |
| `converted_orders.sql` | Ordini (B2C, B2B, ...) e documento in cui sono stati evasi (DDT, fattura), con le date. | `@DataInizio`, `@DataFine` |
| `price_lists_pivot.sql` | Una riga per articolo con marca, EAN e una colonna per ciascun listino (revisione in vigore), costruita con SQL dinamico. In testa si sceglie se includere tutti i listini commerciali o un elenco a scelta. | lista `@Listini` |
| `shipments_report.sql` | DDT del periodo con cliente, categoria, ID ordine WooCommerce, spese di spedizione, presenza di bevande, confezioni totali, peso lordo e imponibile. | `@DataInizio`, `@DataFine` |

Tabelle usate: `MGMov`, `DOTes`, `DORig`, `DOTotali`, `DoRigSpesa`, `CF`,
`CFDest`, `CFContatto`, `CFGruppo1`, `AR`, `ARAlias`, `ARMarca`, `LS`,
`LSRevisione`, `LSArticolo`, `Provincia`, `Regione`.
