# warehouse-locations

Gestione delle ubicazioni di magazzino in ARCA: struttura delle tabelle,
procedura di ricreazione e import delle associazioni articolo/ubicazione da
Excel.

## Problema

Le ubicazioni sono definite in `MGUbicazione` e associate agli articoli in
`ARMGUbicazione`. Riorganizzare il magazzino significa ricreare l'anagrafica
delle ubicazioni e reinserire centinaia di associazioni, operazione lenta
dall'interfaccia e complicata da vincoli e trigger.

## Soluzione

### Formato delle ubicazioni

Formato `NNN.S`: numero della scaffalatura con zero-padding a tre cifre e
livello come suffisso (es. `001.1`, `012.3`, `105.2`). Il padding è necessario
perché la colonna è `varchar` e l'ordinamento è alfabetico: senza padding si
otterrebbe `1.1, 10.1, 100.1, 11.1, 2.1`.

Le ubicazioni vanno da `001` a `NNN` con suffissi `.1`…`.S`. Scaffalature con un
numero di livelli diverso si gestiscono come range distinti:

| Range | Suffissi | Totale |
|-------|----------|--------|
| 001 → NNN | .1 … .S1 | NNN × S1 |
| (NNN+1) → MMM | .1 … .S2 | (MMM − NNN) × S2 |

### Tabelle e vincoli

`MGUbicazione` è referenziata tramite chiave esterna da:

| Tabella | Colonna |
|---------|---------|
| `ARMGUbicazione` | `Cd_MGUbicazione` |
| `DORig` | `Cd_MGUbicazione_P`, `Cd_MGUbicazione_A` |
| `MGMov` | `Cd_MGUbicazione` |
| `MGOrdImp` | `Cd_MGUbicazione` |

`MGMov` e `DORig` hanno trigger di ARCA (`asp_MgMovDocIns`, `asp_MGMovDocUpd`)
che bloccano gli `UPDATE` diretti; le modifiche vanno racchiuse tra
`DISABLE TRIGGER ALL` e `ENABLE TRIGGER ALL`.

### Script di import

`generate_locations_import_sql.py` legge un Excel con le associazioni
articolo/ubicazione e genera `importa_ubicazioni.sql`:

- carica dal database le ubicazioni valide di `MGUbicazione`;
- applica lo zero-padding (`1.1` → `001.1`) e converte la virgola in punto;
- scarta le righe senza `Ubicazione_Preferita` e le ubicazioni inesistenti,
  segnalandole a console;
- usa `IF NOT EXISTS`, quindi lo script generato può essere rieseguito senza
  creare duplicati.

Colonne dell'Excel: `Cd_AR`, `Ubicazione_Preferita` (obbligatoria) e
`Ubicazione_1` … `Ubicazione_8` per le ubicazioni aggiuntive. Eventuali colonne
`Descrizione` e `Categoria` vengono ignorate. Le colonne delle ubicazioni vanno
formattate come testo.

## Utilizzo

### Ricreare le ubicazioni

1. Backup:
   ```cmd
   bcp %DB_ARCA%.dbo.ARMGUbicazione out "C:\Backup\ARMGUbicazione_backup_YYYYMMDD.bcp" -S %SERVER_DB% -T -n
   bcp %DB_ARCA%.dbo.MGUbicazione out "C:\Backup\MGUbicazione_backup_YYYYMMDD.bcp" -S %SERVER_DB% -T -n
   ```
2. Azzerare i riferimenti nelle tabelle figlie:
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
3. Svuotare `ARMGUbicazione` e `MGUbicazione`:
   ```sql
   BEGIN TRANSACTION;

   ALTER TABLE [dbo].[MGMov]          NOCHECK CONSTRAINT FK_MGMov_MGUbicazione;
   ALTER TABLE [dbo].[MGOrdImp]       NOCHECK CONSTRAINT FK_MGOrdImp_MGUbicazione;
   ALTER TABLE [dbo].[DORig]          NOCHECK CONSTRAINT FK_DORig_MGUbicazione_P;
   ALTER TABLE [dbo].[DORig]          NOCHECK CONSTRAINT FK_DORig_MGUbicazione_A;
   ALTER TABLE [dbo].[ARMGUbicazione] NOCHECK CONSTRAINT FK_ARMGUbicazione_MGUbicazione;

   DELETE FROM [dbo].[ARMGUbicazione];
   DELETE FROM [dbo].[MGUbicazione];

   ALTER TABLE [dbo].[MGMov]          CHECK CONSTRAINT FK_MGMov_MGUbicazione;
   ALTER TABLE [dbo].[MGOrdImp]       CHECK CONSTRAINT FK_MGOrdImp_MGUbicazione;
   ALTER TABLE [dbo].[DORig]          CHECK CONSTRAINT FK_DORig_MGUbicazione_P;
   ALTER TABLE [dbo].[DORig]          CHECK CONSTRAINT FK_DORig_MGUbicazione_A;
   ALTER TABLE [dbo].[ARMGUbicazione] CHECK CONSTRAINT FK_ARMGUbicazione_MGUbicazione;

   COMMIT TRANSACTION;
   ```
4. Inserire le nuove ubicazioni, una volta per ogni range:
   ```sql
   DECLARE @Da      INT     = 1;        -- primo numero del range
   DECLARE @A       INT     = 50;       -- ultimo numero del range
   DECLARE @Livelli INT     = 3;        -- numero di suffissi
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

### Importare le associazioni articolo/ubicazione

1. Configurare il `.env` alla radice del repository.
2. Copiare l'Excel nella cartella dello script e verificare che il nome
   corrisponda a `EXCEL_FILE`.
3. Eseguire `python generate_locations_import_sql.py`.
4. Eseguire `importa_ubicazioni.sql` in SSMS.

### Query di verifica

```sql
-- Ubicazioni per numero di livelli e range
SELECT
    MaxLivello,
    MIN(Numero) AS Da,
    MAX(Numero) AS A,
    SUM(Righe)  AS Righe
FROM (
    SELECT LEFT(Cd_MGUbicazione, 3)                           AS Numero,
           MAX(CAST(SUBSTRING(Cd_MGUbicazione, 5, 1) AS INT)) AS MaxLivello,
           COUNT(*)                                           AS Righe
    FROM [dbo].[MGUbicazione]
    GROUP BY LEFT(Cd_MGUbicazione, 3)
) x
GROUP BY MaxLivello;

-- Riferimenti residui nelle tabelle figlie (attesi tutti 0 dopo l'azzeramento)
SELECT 'DORig_P'        AS Tabella, COUNT(*) AS Righe FROM [dbo].[DORig]          WHERE Cd_MGUbicazione_P IS NOT NULL
UNION ALL
SELECT 'DORig_A',                   COUNT(*)          FROM [dbo].[DORig]          WHERE Cd_MGUbicazione_A IS NOT NULL
UNION ALL
SELECT 'MGMov',                     COUNT(*)          FROM [dbo].[MGMov]          WHERE Cd_MGUbicazione   IS NOT NULL
UNION ALL
SELECT 'MGOrdImp',                  COUNT(*)          FROM [dbo].[MGOrdImp]       WHERE Cd_MGUbicazione   IS NOT NULL
UNION ALL
SELECT 'ARMGUbicazione',            COUNT(*)          FROM [dbo].[ARMGUbicazione] WHERE Cd_MGUbicazione   IS NOT NULL;
```
