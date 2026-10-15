# queries

## Problema

Alcune estrazioni ricorrenti (tracciabilità di un articolo, ordini evasi,
confronto tra listini, report spedizioni) non sono disponibili come stampe
standard in ARCA Evolution.

## Soluzione

Query T-SQL sulle tabelle standard di ARCA, con i parametri dichiarati in testa
a ogni file.

| File | Risultato | Parametri |
|------|-----------|-----------|
| `article_sales_traceability.sql` | Movimenti di vendita di un articolo con documento, cliente, contatti, quantità (pezzi e cartoni) e indirizzo di consegna, preso dalla destinazione diversa o dai dati di spedizione quando presenti. | `@Cd_AR`, `@DallaData` |
| `converted_orders.sql` | Ordini (B2C, B2B, ...) con il documento in cui sono stati evasi e le relative date. | `@DataInizio`, `@DataFine` |
| `price_lists_pivot.sql` | Una riga per articolo con marca, EAN e una colonna per ogni listino (revisione in vigore), tramite SQL dinamico. Si possono includere tutti i listini commerciali o un elenco scelto. | lista `@Listini` |
| `shipments_report.sql` | DDT del periodo con cliente, categoria, ID ordine WooCommerce, spese di spedizione, presenza di bevande, confezioni, peso lordo e imponibile. | `@DataInizio`, `@DataFine` |

Tabelle usate: `MGMov`, `DOTes`, `DORig`, `DOTotali`, `DoRigSpesa`, `CF`,
`CFDest`, `CFContatto`, `CFGruppo1`, `AR`, `ARAlias`, `ARMarca`, `LS`,
`LSRevisione`, `LSArticolo`, `Provincia`, `Regione`.

## Utilizzo

Aprire il file in SQL Server Management Studio sul database di ARCA, impostare
i parametri ed eseguire.
