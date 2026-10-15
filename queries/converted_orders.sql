-- Ordini evasi: per ogni ordine (B2C, B2B, ...) mostra il documento in cui
-- è stato convertito (DDT, fattura, ...) e la data di conversione.
DECLARE @DataInizio DATE = '2026-01-01';
DECLARE @DataFine   DATE = '2026-02-01';  -- esclusa

SELECT DISTINCT
    t_ord.Id_DOTes,
    t_ord.NumeroDoc              AS [Numero Ordine],
    t_ord.Cd_DO                  AS [Tipo Documento Ordine],
    CAST(t_ord.DataDoc AS DATE)  AS [Data Ordine],
    cf.Descrizione               AS [Ragione Sociale],
    t_conv.Cd_DO                 AS [Tipo Documento Convertito],
    CAST(t_conv.DataDoc AS DATE) AS [Data Conversione]
FROM dbo.DORig r_ord
JOIN dbo.DOTes  t_ord  ON t_ord.Id_DOTes        = r_ord.Id_DOTes
JOIN dbo.CF     cf     ON cf.Cd_CF              = t_ord.Cd_CF
JOIN dbo.DORig  r_conv ON r_conv.Id_DORig_Evade = r_ord.Id_DORig
JOIN dbo.DOTes  t_conv ON t_conv.Id_DOTes       = r_conv.Id_DOTes
WHERE t_ord.Cd_DO IN ('B2C', 'OAM', 'B2B', 'OW')
    AND t_ord.DataDoc >= @DataInizio AND t_ord.DataDoc < @DataFine
ORDER BY [Data Ordine], [Numero Ordine];
