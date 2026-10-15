# arca-evolution-toolkit

Query SQL e script Python/PowerShell per leggere e scrivere dati sul database
del gestionale ARCA Evolution (SQL Server).

## Contesto

Strumenti sviluppati per le esigenze di un'azienda di distribuzione che usa
ARCA Evolution, pubblicati in forma anonimizzata con l'autorizzazione
dell'azienda.

## Problema

Diverse operazioni ricorrenti sul gestionale, se fatte dall'interfaccia,
richiedono molto lavoro manuale: allineare le anagrafiche clienti con il sito
B2B, aggiornare un listino a partire da un Excel, esportare in blocco le
fatture passive con i relativi PDF, preparare i dati di un catalogo prodotti.

## Soluzione

Gli strumenti lavorano direttamente sulle tabelle standard di ARCA (`CF`, `AR`,
`DOTes`, `DORig`, `LSArticolo`, `MGUbicazione`, ...) e seguono tre schemi:

- query in `queries/`, da eseguire in SSMS, con i parametri dichiarati in testa;
- script di import che leggono Excel o CSV, li confrontano con ARCA e generano
  un file `.sql` da rivedere ed eseguire manualmente, senza scrivere
  direttamente sul database;
- script di export che leggono da ARCA e producono Excel, XML o PDF.

## Struttura

```
arca-evolution-toolkit/
├── queries/                              query SQL per SSMS
│   ├── article_sales_traceability.sql    tracciabilità delle vendite di un articolo
│   ├── converted_orders.sql              ordini e documenti in cui sono stati evasi
│   ├── price_lists_pivot.sql             listini affiancati per articolo
│   └── shipments_report.sql              spedizioni (DDT) con pesi e importi
├── data-import/
│   ├── woocommerce-customer-categories/  confronto categorie clienti WooCommerce / ARCA
│   ├── woocommerce-role-mapping/         mappatura dei ruoli WooCommerce su CFGruppo
│   ├── woocommerce-product-tags/         import dei tag prodotto WooCommerce
│   ├── warehouse-locations/              associazione articoli / ubicazioni
│   └── price-list-import/                aggiornamento di un listino da Excel
└── data-export/
    ├── supplier-invoices-export/         export delle fatture passive XML + PDF
    ├── document-lines-export/            righe di un documento in Excel
    └── product-catalog-export/           catalogo prodotti in Excel
```

Ogni sottocartella ha un README con i dettagli d'uso.

## Tecnologie

- ARCA Evolution su Microsoft SQL Server (T-SQL, SQL dinamico con `sp_executesql`)
- Python 3.10+: pandas, openpyxl, pyodbc, SQLAlchemy, tkinter, python-dotenv
- PowerShell 5.1: ADO.NET, `System.Security.Cryptography.Pkcs` per i file `.p7m`

## Configurazione

1. Installare il driver ODBC per SQL Server e le dipendenze:
   ```bash
   pip install -r requirements.txt
   ```
2. Copiare `.env.example` in `.env` e impostare:

   | Variabile     | Significato                                                 |
   |---------------|-------------------------------------------------------------|
   | `SERVER_DB`   | server e istanza SQL Server                                 |
   | `DB_ARCA`     | nome del database di ARCA                                   |
   | `DB_USER`     | utente SQL Server; se vuoto si usa l'autenticazione Windows |
   | `DB_PASSWORD` | password dell'utente SQL Server                             |

Gli script Python leggono il `.env` dalla radice del repository; lo script
PowerShell lo legge da solo, se presente. Il file `.env` non è versionato.
