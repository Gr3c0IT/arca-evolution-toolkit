-- Tracciabilità delle vendite di un articolo: per ogni movimento di scarico
-- restituisce documento, cliente, contatti, quantità e indirizzo di consegna.
-- Parametri: codice articolo e data di inizio.
DECLARE @Cd_AR     VARCHAR(20) = 'ART001';
DECLARE @DallaData DATE        = '2025-01-01';

SELECT
    CAST(mg.DataMov  as Date)                     AS [DataMov],
    dr.Id_DOTes,
    dr.NumeroDoc,
    dr.Cd_DO AS [Tipo Documento],
    cf.Descrizione                  AS [Ragione Sociale],
    cf.Indirizzo                    AS [Via/Località],
    cf.Localita                     AS [Comune],
    cf.Cd_Provincia                 AS [Provincia],
    reg.Descrizione                 AS [Regione],
    cf.Cd_Nazione                   AS [Nazione],
    cfc.Telefono                    AS [Tel/Fax],
    cfc.Email                       AS [E-Mail],
    'PZ'                            AS [Tipologia Confezioni],
    CAST(mg.Quantita AS INT)        AS [Quantitativo Venduto],
    'CF'                            AS [Tipologia Confezioni 2],
    CAST(mg.Quantita / 12  AS INT)              AS [Nr. Confezioni 2],
CASE 
    WHEN dot.xDatiSHP IS NOT NULL AND dot.xDatiSHP <> '' THEN 
        REPLACE(REPLACE(TRIM(dot.xDatiSHP), CHAR(13), ' '), CHAR(10), ' ')
    WHEN cfd.Descrizione IS NOT NULL THEN CONCAT(
        cfd.Descrizione,  ' - ',
        cfd.Indirizzo,    ', ',
        cfd.Cap,          ' ',
        cfd.Localita,     ' (',
        cfd.Cd_Provincia, ') ',
        cfd.Cd_Nazione
    )
    ELSE CONCAT(
        cf.Descrizione,  ' - ',
        cf.Indirizzo,    ', ',
        cf.Cap,          ' ',
        cf.Localita,     ' (',
        cf.Cd_Provincia, ') ',
        cf.Cd_Nazione
    )
END AS [Indirizzo di Consegna],
    CAST(dr.DataConsegna_R AS DATE) AS [Data Consegna]
FROM
    dbo.MGMov mg
    INNER JOIN dbo.DORig dr     ON dr.Id_DoRig    = mg.Id_DoRig
    INNER JOIN dbo.DOTes dot    ON dot.Id_DOTes   = dr.Id_DOTes
    INNER JOIN dbo.CF cf        ON cf.Cd_CF        = dr.Cd_CF
    LEFT  JOIN dbo.CFDest cfd   ON cfd.Cd_CF       = dot.Cd_CF
                                AND cfd.Cd_CFDest   = dot.Cd_CFDest
    LEFT  JOIN dbo.Provincia pv ON pv.Cd_Provincia = cf.Cd_Provincia
                                AND pv.Cd_Nazione  = cf.Cd_Nazione
    LEFT  JOIN dbo.Regione reg  ON reg.Cd_Regione  = pv.Cd_Regione
                                AND reg.Cd_Nazione  = pv.Cd_Nazione
    OUTER APPLY ( 
        SELECT TOP 1 Telefono, Email
        FROM dbo.CFContatto
        WHERE Cd_CF   = cf.Cd_CF
          AND Obsoleto = 0
        ORDER BY Sequenza
    ) cfc
    
WHERE
    mg.Cd_AR   = @Cd_AR
    AND mg.DataMov >= @DallaData;
