# arca-evolution-toolkit

Raccolta di query SQL e script Python/PowerShell per leggere e scrivere dati sul
database del gestionale **ARCA Evolution** (SQL Server): estrazioni, import da
WooCommerce, aggiornamento di anagrafiche, listini e ubicazioni, export di
documenti e fatture.

## Che problema risolve

ARCA Evolution copre bene la gestione quotidiana, ma molte operazioni massive
(allineare centinaia di anagrafiche con il sito B2B, aggiornare un listino da un
Excel, estrarre in blocco le fatture passive con i PDF, preparare i dati per un
catalogo) richiederebbero ore di lavoro manuale dall'interfaccia. Questi
strumenti le automatizzano lavorando direttamente sulle tabelle standard di
ARCA (`CF`, `AR`, `DOTes`, `DORig`, `LSArticolo`, `MGUbicazione`, ...).

## Come funziona

- **Le query** in `queries/` si eseguono in SSMS e hanno i parametri dichiarati
  in testa (`DECLARE`).
- **Gli script di import** non scrivono mai direttamente sul database:
  leggono i dati sorgente (Excel o CSV), li confrontano con ARCA e **generano un
  file `.sql`** da rivedere ed eseguire manualmente in SSMS.
- **Gli script di export** leggono da ARCA e producono file Excel, XML o PDF.

## Struttura delle cartelle

```
arca-evolution-toolkit/
├── queries/                              query SQL pronte per SSMS
│   ├── article_sales_traceability.sql    tracciabilità delle vendite di un articolo
│   ├── converted_orders.sql              ordini e documenti in cui sono stati evasi
│   ├── price_lists_pivot.sql             tutti i listini affiancati per articolo
│   └── shipments_report.sql              report spedizioni (DDT) con pesi e importi
├── data-import/
│   ├── woocommerce-customer-categories/  confronto categorie clienti WooCommerce ↔ ARCA (CSV)
│   ├── woocommerce-role-mapping/         mappatura fuzzy dei ruoli WooCommerce su CFGruppo (Excel)
│   ├── woocommerce-product-tags/         import dei tag prodotto WooCommerce
│   ├── warehouse-locations/              associazione articoli ↔ ubicazioni di magazzino
│   └── price-list-import/                aggiornamento di un listino ARCA da Excel (GUI)
└── data-export/
    ├── supplier-invoices-export/         export massivo delle fatture passive XML + PDF (PowerShell)
    ├── document-lines-export/            estrazione delle righe di un documento in Excel (GUI)
    └── product-catalog-export/           generazione del catalogo prodotti in Excel
```

Ogni sottocartella ha un proprio `README.md` con i dettagli d'uso.

## Tecnologie usate

- ARCA Evolution su Microsoft SQL Server (T-SQL, query dinamiche con `sp_executesql`)
- Python 3.10+: pandas, openpyxl, pyodbc, SQLAlchemy, tkinter, python-dotenv
- PowerShell 5.1 (ADO.NET, `System.Security.Cryptography.Pkcs` per i file `.p7m`)

## Configurazione

1. Installare le dipendenze Python:
   ```bash
   pip install -r requirements.txt
   ```
2. Installare il driver ODBC per SQL Server.
3. Copiare `.env.example` in `.env` e compilare i valori:

   | Variabile     | Significato                                                  |
   |---------------|--------------------------------------------------------------|
   | `SERVER_DB`   | server e istanza SQL Server (es. `server\ISTANZA`)           |
   | `DB_ARCA`     | nome del database di ARCA                                    |
   | `DB_USER`     | utente SQL Server; se vuoto si usa l'autenticazione Windows  |
   | `DB_PASSWORD` | password dell'utente SQL Server                              |

   Tutti gli script Python leggono il `.env` dalla radice del repository. Lo
   script PowerShell lo legge da solo, se il file è presente.

> ⚠️ Il file `.env` e tutti i file di dati (Excel, CSV, XML, PDF) sono esclusi
> da `.gitignore`: non vanno mai pubblicati.
