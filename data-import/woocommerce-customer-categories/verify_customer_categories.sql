-- Verifica post-importazione: tutte le colonne dell'anagrafica CF dei clienti
-- collegati a WooCommerce, con la descrizione delle tre categorie.
DECLARE @Colonne NVARCHAR(MAX);
DECLARE @Sql     NVARCHAR(MAX);

SELECT @Colonne =
    STRING_AGG(
        CAST(
            CASE
                WHEN C.name = 'Cd_CFGruppo1' THEN
                    'CF.' + QUOTENAME(C.name) +
                    ', G1.Descrizione AS Categoria1'

                WHEN C.name = 'Cd_CFGruppo2' THEN
                    'CF.' + QUOTENAME(C.name) +
                    ', G2.Descrizione AS Categoria2'

                WHEN C.name = 'Cd_CFGruppo3' THEN
                    'CF.' + QUOTENAME(C.name) +
                    ', G3.Descrizione AS Categoria3'

                ELSE
                    'CF.' + QUOTENAME(C.name)
            END
        AS NVARCHAR(MAX)),
        ',' + CHAR(13) + CHAR(10) + '    '
    ) WITHIN GROUP (ORDER BY C.column_id)
FROM sys.columns AS C
WHERE C.object_id = OBJECT_ID('dbo.CF');

SET @Sql = N'
SELECT
    ' + @Colonne + '
FROM dbo.CF AS CF

LEFT JOIN dbo.CFGruppo1 AS G1
    ON G1.Cd_CFGruppo1 = CF.Cd_CFGruppo1

LEFT JOIN dbo.CFGruppo2 AS G2
    ON G2.Cd_CFGruppo2 = CF.Cd_CFGruppo2

LEFT JOIN dbo.CFGruppo3 AS G3
    ON G3.Cd_CFGruppo3 = CF.Cd_CFGruppo3

WHERE CF.xWooCommerce_ID_B2B IS NOT NULL
  AND CF.Cd_CFGruppo1 IS NOT NULL;
';

EXEC sys.sp_executesql @Sql;
