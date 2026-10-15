-- Prezzi di tutti i listini affiancati per articolo (pivot dinamico sulla
-- revisione corrente di ciascun listino), con marca ed EAN.
DECLARE @Listini TABLE (Ord int IDENTITY(1,1), Cd_LS varchar(7));

-- Opzione 1: TUTTI i listini commerciali (esclude i 4 automatici ultima vendita/acquisto)
INSERT INTO @Listini (Cd_LS)
SELECT RTRIM(Cd_LS) FROM LS WHERE TipoLS = 1 ORDER BY Cd_LS;

-- Opzione 2 (in alternativa): scegli tu quali e in che ordine — commenta l'INSERT sopra
-- INSERT INTO @Listini (Cd_LS) VALUES ('B2C'), ('B2B'), ('LISTINO_A'), ('LISTINO_B');

DECLARE @cols nvarchar(max) = N'', @inlist nvarchar(max) = N'';
SELECT @cols = @cols + N',
    CAST(MAX(CASE WHEN p.Cd_LS = ' + QUOTENAME(RTRIM(Cd_LS), '''') +
    N' THEN p.Prezzo END) AS decimal(18,2)) AS ' + QUOTENAME(RTRIM(Cd_LS)),
       @inlist = @inlist + N',' + QUOTENAME(RTRIM(Cd_LS), '''')
FROM @Listini ORDER BY Ord;
SET @inlist = STUFF(@inlist, 1, 1, N'');

DECLARE @sql nvarchar(max) = N'
;WITH RevCorrente AS (
    SELECT r.Id_LSRevisione, RTRIM(r.Cd_LS) AS Cd_LS,
           ROW_NUMBER() OVER (PARTITION BY r.Cd_LS
                              ORDER BY r.DaData DESC, r.Id_LSRevisione DESC) AS rn
    FROM LSRevisione r
    WHERE r.DaData <= GETDATE()
),
Prezzi AS (
    SELECT a.Cd_AR, rc.Cd_LS, a.Prezzo
    FROM LSArticolo a
    JOIN RevCorrente rc ON rc.Id_LSRevisione = a.Id_LSRevisione AND rc.rn = 1
    WHERE rc.Cd_LS IN (' + @inlist + N')
),
Ean AS (
    SELECT Cd_AR, LTRIM(RTRIM(Alias)) AS EAN,
           ROW_NUMBER() OVER (PARTITION BY Cd_AR ORDER BY Riga, Id_ARAlias) AS rn
    FROM ARAlias
    WHERE TipoAlias LIKE ''%[Ee]%''
)
SELECT
    AR.Cd_AR,
    mk.Descrizione AS Marca,
    AR.Descrizione,
    e.EAN' + @cols + N'
FROM AR
LEFT JOIN ARMarca mk ON mk.Cd_ARMarca = AR.Cd_ARMarca
LEFT JOIN Ean e      ON e.Cd_AR = AR.Cd_AR AND e.rn = 1
LEFT JOIN Prezzi p   ON p.Cd_AR = AR.Cd_AR
GROUP BY AR.Cd_AR, mk.Descrizione, AR.Descrizione, e.EAN
ORDER BY AR.Cd_AR;';

EXEC sys.sp_executesql @sql;
