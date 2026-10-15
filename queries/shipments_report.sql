-- Report spedizioni (DDT): cliente, categoria, ordine WooCommerce, spese di
-- spedizione, presenza di bevande, confezioni, peso e imponibile.
-- Parametri: intervallo di date del movimento di magazzino.
DECLARE @DataInizio DATE = '2026-01-01';
DECLARE @DataFine   DATE = '2026-01-31';

SELECT
    m.DataMov,
    t.Id_DOTes,
    t.Cd_CF                                          AS [Codice Cliente],
    cf.Descrizione                                   AS [Intestazione],
    cgr.Descrizione                                   AS [Categoria Venditore],
    t.xWooCommerce_ID_B2B                            AS [Identificativo Ordine WooCommerce],
    s.Descrizione                                    AS [Spedizione],
    CAST(s.ImportoV AS DECIMAL(18, 2))               AS [Importo Spedizione],
    CASE WHEN bev.Id_DOTes IS NOT NULL
         THEN 'Sì' ELSE 'No' END                     AS [Bevande],
    CAST(conf.ConfezioniTotali AS INT)    AS [Confezioni Totali],
    CAST(t.PesoLordo AS DECIMAL(18, 2))              AS [Peso Lordo],
    CAST(tot.TotImponibileE AS DECIMAL(18, 2))       AS [Totale Imponibile]
FROM (
    -- Una sola riga per DDT con DataMov aggregata
    SELECT r.Id_DOTes, CAST(MIN(m.DataMov) AS DATE) AS DataMov
    FROM MGMov m
    INNER JOIN DORig r ON r.Id_DORig = m.Id_DORig
    WHERE m.Cd_MGCausale = 'DDT'
    GROUP BY r.Id_DOTes
) m
INNER JOIN DOTes      t   ON t.Id_DOTes  = m.Id_DOTes
INNER JOIN CF         cf  ON cf.Cd_CF    = t.Cd_CF
LEFT JOIN CFGruppo1  cgr  ON cgr.Cd_CFGruppo1 = cf.Cd_CFGruppo1
LEFT  JOIN DoRigSpesa s   ON s.Id_DOTes  = t.Id_DOTes

-- Bevande: Sì se almeno un articolo del DDT ha Cd_ARGruppo1 = '001'
LEFT JOIN (
    SELECT DISTINCT r.Id_DOTes
    FROM DORig r
    INNER JOIN AR ar ON ar.Cd_AR = r.Cd_AR
    WHERE ar.Cd_ARGruppo1 = '001'
) bev ON bev.Id_DOTes = t.Id_DOTes

-- Confezioni totali: somma Qta * FattoreToUM1 per articoli con Cd_AR valorizzato
LEFT JOIN (
    SELECT r.Id_DOTes, SUM(r.Qta * r.FattoreToUM1) AS ConfezioniTotali
    FROM DORig r
    WHERE r.Cd_AR IS NOT NULL AND r.Cd_AR <> ''
    GROUP BY r.Id_DOTes
) conf ON conf.Id_DOTes = t.Id_DOTes

-- Totale imponibile da DOTotali (pre-aggregato nel caso ci siano più aliquote IVA)
LEFT JOIN (
    SELECT Id_DOTes, SUM(TotImponibileE) AS TotImponibileE
    FROM DOTotali
    GROUP BY Id_DOTes
) tot ON tot.Id_DOTes = t.Id_DOTes

-- WHERE cf.Cd_CFGruppo1 = 'VEN'   -- es. solo una categoria cliente
WHERE CAST(m.DataMov AS DATE) BETWEEN @DataInizio AND @DataFine AND t.Cd_DO = 'DDT'

ORDER BY
    m.DataMov DESC,
    t.Id_DOTes,
    s.Descrizione;
