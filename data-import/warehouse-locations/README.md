# warehouse-locations — Ubicazioni di magazzino

## Panoramica

Il sistema di ubicazioni magazzino è composto da due tabelle del database di ARCA:

- **`MGUbicazione`** — anagrafica delle ubicazioni (il "catalogo" di tutte le ubicazioni esistenti)
- **`ARMGUbicazione`** — associazione articolo ↔ ubicazione

---

## Formato delle ubicazioni

Le ubicazioni usano il formato `NNN.S` con zero-padding obbligatorio sul numero:

- ✅ `001.1`, `012.3`, `105.2`
- ❌ `1.1`, `12.3`

Lo zero-padding è fondamentale perché la colonna è un `varchar` e l'ordinamento è alfabetico. Senza padding, l'ordine risultante sarebbe `1.1 → 10.1 → 100.1 → 11.1 → 2.1...` invece di `001.1 → 002.1 → ... → NNN.S`.

### Range

Le ubicazioni vanno da `001` a `NNN` (numero della scaffalatura) con un suffisso
`.1`…`.S` (livello). Se le scaffalature hanno un numero di livelli diverso si
definiscono più range, ad esempio:

| Range | Suffissi | Totale |
|-------|----------|--------|
| 001 → NNN | .1 … .S1 | NNN × S1 |
| (NNN+1) → MMM | .1 … .S2 | (MMM − NNN) × S2 |

---

## Tabelle coinvolte e FK

`MGUbicazione` è referenziata da molte tabelle figlie tramite FK:

| Tabella | Colonna | Note |
|---------|---------|------|
| `ARMGUbicazione` | `Cd_MGUbicazione` | Associazione articolo ↔ ubicazione |
| `DORig` | `Cd_MGUbicazione_P`, `Cd_MGUbicazione_A` | Righe documenti |
| `MGMov` | `Cd_MGUbicazione` | Movimenti magazzino |
| `MGOrdImp` | `Cd_MGUbicazione` | Ordini impegnati |

> ⚠️ `MGMov` e `DORig` hanno trigger ARCA (`asp_MgMovDocIns`, `asp_MGMovDocUpd`) che bloccano gli UPDATE diretti. Usare sempre `DISABLE TRIGGER ALL` / `ENABLE TRIGGER ALL`.

---

## Procedura: ricreare le ubicazioni da zero

### 1. Backup BCP

```cmd
bcp %DB_ARCA%.dbo.ARMGUbicazione out "C:\Backup\ARMGUbicazione_backup_YYYYMMDD.bcp" -S %SERVER_DB% -T -n
bcp %DB_ARCA%.dbo.MGUbicazione out "C:\Backup\MGUbicazione_backup_YYYYMMDD.bcp" -S %SERVER_DB% -T -n
```

### 2. Azzerare i riferimenti nelle tabelle figlie

```sql
DISABLE TRIGGER ALL ON [dbo].[DORig];
UPDATE [dbo].[DORig]
SET Cd_MGUbicazione_P = NULL, Cd_MGUbicazione_A = NULL
WHERE Cd_MGUbicazione_P IS NOT NULL OR Cd_MGUbicazione_A IS NOT NULL;
ENABLE TRIGGER ALL ON [dbo].[DORig];

DISABLE TRIGGER ALL ON [dbo].[MGMov];
UPDATE [dbo].[MGMov] SET Cd_MGUbicazione = NULL WHERE Cd_MGUbicazione IS NOT NULL;
ENABLE TRIGGER ALL ON [dbo].[MGMov];

DISABLE TRIGGER ALL ON [dbo].[MGOrdImp];
UPDATE [dbo].[MGOrdImp] SET Cd_MGUbicazione = NULL WHERE Cd_MGUbicazione IS NOT NULL;
ENABLE TRIGGER ALL ON [dbo].[MGOrdImp];
```

### 3. Svuotare ARMGUbicazione e MGUbicazione

```sql
BEGIN TRANSACTION;

ALTER TABLE [dbo].[MGMov]          NOCHECK CONSTRAINT FK_MGMov_MGUbicazione;
ALTER TABLE [dbo].[MGOrdImp]        NOCHECK CONSTRAINT FK_MGOrdImp_MGUbicazione;
ALTER TABLE [dbo].[DORig]           NOCHECK CONSTRAINT FK_DORig_MGUbicazione_P;
ALTER TABLE [dbo].[DORig]           NOCHECK CONSTRAINT FK_DORig_MGUbicazione_A;
ALTER TABLE [dbo].[ARMGUbicazione]  NOCHECK CONSTRAINT FK_ARMGUbicazione_MGUbicazione;

DELETE FROM [dbo].[ARMGUbicazione];
DELETE FROM [dbo].[MGUbicazione];

ALTER TABLE [dbo].[MGMov]          CHECK CONSTRAINT FK_MGMov_MGUbicazione;
ALTER TABLE [dbo].[MGOrdImp]        CHECK CONSTRAINT FK_MGOrdImp_MGUbicazione;
ALTER TABLE [dbo].[DORig]           CHECK CONSTRAINT FK_DORig_MGUbicazione_P;
ALTER TABLE [dbo].[DORig]           CHECK CONSTRAINT FK_DORig_MGUbicazione_A;
ALTER TABLE [dbo].[ARMGUbicazione]  CHECK CONSTRAINT FK_ARMGUbicazione_MGUbicazione;

COMMIT TRANSACTION;
```

### 4. Inserire le nuove ubicazioni in MGUbicazione

Eseguire una volta per ogni range, impostando i parametri:

```sql
DECLARE @Da      INT     = 1;        -- primo numero del range (001)
DECLARE @A       INT     = 50;       -- ultimo numero del range (NNN)
DECLARE @Livelli INT     = 3;        -- numero di suffissi (.1 ... .S)
DECLARE @Cd_MG   CHAR(5) = '00002';  -- magazzino

INSERT INTO [dbo].[MGUbicazione] (Cd_MGUbicazione, Cd_MG, Descrizione)
SELECT
    RIGHT('000' + CAST(n.number AS VARCHAR(3)), 3) + '.' + CAST(s.number AS VARCHAR(1)),
    @Cd_MG,
    RIGHT('000' + CAST(n.number AS VARCHAR(3)), 3) + '.' + CAST(s.number AS VARCHAR(1))
FROM master..spt_values n
CROSS JOIN master..spt_values s
WHERE n.type = 'P' AND n.number BETWEEN @Da AND @A
  AND s.type = 'P' AND s.number BETWEEN 1 AND @Livelli;
```

---

## Script Python — generate_locations_import_sql.py

Lo script legge un file Excel con le assegnazioni articolo ↔ ubicazione e genera un file `.sql` da eseguire su SSMS.

### Struttura del file Excel atteso

| Cd_AR | Descrizione | Categoria | Ubicazione_Preferita | Ubicazione_1 | ... | Ubicazione_8 |
|-------|-------------|-----------|----------------------|--------------|-----|--------------|

> Le colonne `Descrizione` e `Categoria` sono facoltative e vengono ignorate dallo script.

### Regole di compilazione del file Excel

- `Ubicazione_Preferita` è **obbligatoria** — senza di essa la riga viene scartata
- Se l'articolo ha una sola ubicazione → mettila solo in `Ubicazione_Preferita`
- Se l'articolo ha più ubicazioni → una in `Ubicazione_Preferita`, le altre da `Ubicazione_1` in avanti
- Formattare le colonne come **testo** in Excel per evitare che gli zeri vengano eliminati
- Usare il **punto** come separatore (es. `012.3`), non la virgola

### Come usare lo script

1. Configura il file `.env` nella radice del repository (`SERVER_DB`, `DB_ARCA`, ...)
2. Metti il file Excel nella stessa cartella dello script
3. Verifica che `EXCEL_FILE` nello script corrisponda al nome del file Excel
4. Lancia: `python generate_locations_import_sql.py`
5. Lo script stampa a console gli articoli skippati (ubicazione non valida o mancante)
6. Viene generato `importa_ubicazioni.sql` nella stessa cartella
7. Esegui il file `.sql` su SSMS

### Comportamento dello script

- Si connette al database e carica le ubicazioni valide da `MGUbicazione`
- Applica automaticamente lo zero-padding (`1.1` → `001.1`) e sostituisce la virgola col punto
- Skippa le righe senza `Ubicazione_Preferita`
- Skippa le ubicazioni che non esistono in `MGUbicazione` (le stampa a console)
- Usa `WHERE NOT EXISTS` → rilancio sicuro, nessun duplicato

---

## Query di verifica utili

```sql
-- Conteggio ubicazioni per numero di livelli (suffisso massimo) e range
SELECT
    MaxLivello,
    MIN(Numero) AS Da,
    MAX(Numero) AS A,
    SUM(Righe)  AS Righe
FROM (
    SELECT LEFT(Cd_MGUbicazione, 3)                       AS Numero,
           MAX(CAST(SUBSTRING(Cd_MGUbicazione, 5, 1) AS INT)) AS MaxLivello,
           COUNT(*)                                       AS Righe
    FROM [dbo].[MGUbicazione]
    GROUP BY LEFT(Cd_MGUbicazione, 3)
) x
GROUP BY MaxLivello;

-- Verifica riferimenti orfani nelle tabelle figlie (devono essere tutti 0)
SELECT 'DORig_P'        AS Tabella, COUNT(*) AS Righe FROM [dbo].[DORig]        WHERE Cd_MGUbicazione_P IS NOT NULL
UNION ALL
SELECT 'DORig_A',                   COUNT(*) FROM [dbo].[DORig]                 WHERE Cd_MGUbicazione_A IS NOT NULL
UNION ALL
SELECT 'MGMov',                     COUNT(*) FROM [dbo].[MGMov]                 WHERE Cd_MGUbicazione   IS NOT NULL
UNION ALL
SELECT 'MGOrdImp',                  COUNT(*) FROM [dbo].[MGOrdImp]              WHERE Cd_MGUbicazione   IS NOT NULL
UNION ALL
SELECT 'ARMGUbicazione',            COUNT(*) FROM [dbo].[ARMGUbicazione]        WHERE Cd_MGUbicazione   IS NOT NULL;
```
