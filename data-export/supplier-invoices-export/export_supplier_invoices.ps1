# =====================================================================
#  Export massivo fatture passive da ARCA (DmsDocument) -> XML + PDF
#  Filtro: conto contabile + intervallo data registrazione
#  Nome file: <conto>_<CdCF>_<NumFattura>_<DataReg>.xml / .pdf
#  Compatibile PowerShell 5.1
#
#  Connessione: variabili d'ambiente SERVER_DB, DB_ARCA, DB_USER, DB_PASSWORD
#  (lette anche dal file .env nella radice del repository).
#
#  Uso:
#    .\export_supplier_invoices.ps1 -Conto 00000000000
#    .\export_supplier_invoices.ps1 -Conto 00000000000 -DataInizio 2026-01-01 -DataFine 2026-12-31
# =====================================================================
param(
    [Parameter(Mandatory = $true)]
    [string]  $Conto,
    [datetime]$DataInizio = (Get-Date -Year (Get-Date).Year -Month 1 -Day 1).Date,
    [datetime]$DataFine   = (Get-Date -Year (Get-Date).Year -Month 12 -Day 31).Date,
    [string]  $Server     = "",
    [string]  $Db         = "",
    [string]  $OutRoot    = $PSScriptRoot,
    [string]  $Xsl        = (Join-Path $PSScriptRoot "FoglioStile.xsl"),
    [string]  $FmtData    = "yyyyMMdd"
)

$ErrorActionPreference = 'Stop'

# Carica le variabili dal file .env del repository (se presente)
$envFile = Join-Path $PSScriptRoot "..\..\.env"
if (Test-Path $envFile) {
    foreach ($line in Get-Content $envFile) {
        if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$' -and -not [Environment]::GetEnvironmentVariable($Matches[1])) {
            [Environment]::SetEnvironmentVariable($Matches[1], $Matches[2].Trim('"'))
        }
    }
}
if (-not $Server) { $Server = $env:SERVER_DB }
if (-not $Db)     { $Db     = $env:DB_ARCA }
if (-not $Server -or -not $Db) { throw "Imposta SERVER_DB e DB_ARCA (file .env o variabili d'ambiente)." }
$auth = if ($env:DB_USER) { "User ID=$($env:DB_USER);Password=$($env:DB_PASSWORD)" } else { "Integrated Security=True" }
Add-Type -AssemblyName System.Security

$out = Join-Path $OutRoot ("Fatture_{0}_{1}_{2}" -f $Conto, $DataInizio.ToString("yyyyMMdd"), $DataFine.ToString("yyyyMMdd"))
New-Item -ItemType Directory -Path $out -Force | Out-Null

# ------------------------------ FUNZIONI -----------------------------
function Get-NomeSicuro([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return (($s.Trim()) -replace '[\\/:*?"<>|\s]+', '-').Trim('-')
}

function Expand-ArcaBlob([byte[]]$b) {
    # Formato ARCA: record "central directory" ZIP (46 byte + nome + extra + commento) + dati Deflate
    if ($b.Length -gt 46 -and $b[0] -eq 0x50 -and $b[1] -eq 0x4B -and $b[2] -eq 0x01 -and $b[3] -eq 0x02) {
        $start = 46 + [BitConverter]::ToUInt16($b, 28) + [BitConverter]::ToUInt16($b, 30) + [BitConverter]::ToUInt16($b, 32)
        $ms = New-Object IO.MemoryStream -ArgumentList @($b, $start, ($b.Length - $start))
        $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
        $o  = New-Object IO.MemoryStream
        $ds.CopyTo($o); $ds.Dispose()
        return ,$o.ToArray()
    }
    return ,$b
}

function Get-XmlDaP7m([byte[]]$p) {
    if ($p[0] -ne 0x30) {   # p7m in Base64
        $p = [Convert]::FromBase64String(([Text.Encoding]::ASCII.GetString($p) -replace '\s', ''))
    }
    $cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
    $cms.Decode($p)
    return ,$cms.ContentInfo.Content
}

function Test-IsPdf([byte[]]$b) {
    return ($b.Length -gt 4 -and $b[0] -eq 0x25 -and $b[1] -eq 0x50 -and $b[2] -eq 0x44 -and $b[3] -eq 0x46)  # %PDF
}

function Get-EdgePath {
    foreach ($p in @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                     "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe")) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

$script:xslt = $null
function New-PdfDaXml([string]$xmlPath, [string]$pdfPath) {
    if (-not $script:edge) { return $false }
    if (-not $script:xslt) {
        $settings = New-Object System.Xml.Xsl.XsltSettings($true, $false)   # document() si, script no
        $script:xslt = New-Object System.Xml.Xsl.XslCompiledTransform
        $script:xslt.Load($Xsl, $settings, (New-Object System.Xml.XmlUrlResolver))
    }
    $html = [IO.Path]::ChangeExtension($pdfPath, ".html")
    $script:xslt.Transform($xmlPath, $html)

    $profilo = Join-Path $env:TEMP ("edge_pdf_" + [guid]::NewGuid().ToString("N"))
    Start-Process -FilePath $script:edge -Wait -WindowStyle Hidden -ArgumentList @(
        "--headless", "--disable-gpu", "--no-pdf-header-footer",
        "--user-data-dir=`"$profilo`"",
        "--print-to-pdf=`"$pdfPath`"",
        "`"file:///$($html -replace '\\','/')`""
    )
    Remove-Item $html -ErrorAction SilentlyContinue
    Remove-Item $profilo -Recurse -Force -ErrorAction SilentlyContinue
    return (Test-Path $pdfPath)
}

function Get-NomeUnivoco([string]$base) {
    # evita sovrascritture se due documenti producono lo stesso nome
    $nome = $base; $i = 2
    while ($script:nomiUsati.Contains($nome)) { $nome = "{0}_{1}" -f $base, $i; $i++ }
    [void]$script:nomiUsati.Add($nome)
    return $nome
}

# --------------------------- CONTROLLI INIZIALI ----------------------
$script:edge = Get-EdgePath
$puoGenerarePdf = $true
if (-not (Test-Path $Xsl))  { Write-Host "ATTENZIONE: foglio di stile non trovato ($Xsl). I PDF mancanti non saranno generati." -ForegroundColor Yellow; $puoGenerarePdf = $false }
if (-not $script:edge)      { Write-Host "ATTENZIONE: msedge.exe non trovato. I PDF mancanti non saranno generati." -ForegroundColor Yellow; $puoGenerarePdf = $false }

Write-Host ("Conto {0} - registrazioni dal {1:dd/MM/yyyy} al {2:dd/MM/yyyy}" -f $Conto, $DataInizio, $DataFine)
Write-Host "Cartella: $out`n"

# ----------------------------- QUERY ---------------------------------
$filtroMov = @"
FROM dbo.CGMovT t
LEFT JOIN dbo.CF cf ON cf.Cd_CF = t.Cd_CF
WHERE t.DtReg >= @Inizio
  AND t.DtReg <  @Fine
  AND EXISTS (SELECT 1 FROM dbo.CGMovR r
              WHERE r.Id_CGMovT = t.Id_CGMovT
                AND r.Cd_CGConto = @Conto)
"@

$sqlDoc = @"
SELECT t.Id_CGMovT, t.DtReg, t.NumProt, t.NumRif, t.Cd_CF, cf.Descrizione AS Fornitore,
       d.Id_DmsDocument, d.FileName, d.Content
FROM dbo.CGMovT t
JOIN dbo.DmsDocument d ON d.EntityTable = 'CGMovT' AND d.EntityId = t.Id_CGMovT AND d.Tag = 'FTE'
LEFT JOIN dbo.CF cf ON cf.Cd_CF = t.Cd_CF
WHERE t.DtReg >= @Inizio
  AND t.DtReg <  @Fine
  AND EXISTS (SELECT 1 FROM dbo.CGMovR r
              WHERE r.Id_CGMovT = t.Id_CGMovT
                AND r.Cd_CGConto = @Conto)
ORDER BY t.DtReg, t.Id_CGMovT, d.Id_DmsDocument
"@

$sqlSenzaXml = @"
SELECT t.Id_CGMovT, t.DtReg, t.NumProt, t.NumRif, t.Cd_CF, cf.Descrizione AS Fornitore
$filtroMov
  AND NOT EXISTS (SELECT 1 FROM dbo.DmsDocument d
                  WHERE d.EntityTable = 'CGMovT' AND d.EntityId = t.Id_CGMovT AND d.Tag = 'FTE')
ORDER BY t.DtReg, t.Id_CGMovT
"@

function Invoke-Query([string]$sql) {
    $cn = New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$Db;$auth"
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $sql
        $cmd.CommandTimeout = 300
        [void]$cmd.Parameters.Add("@Conto",  [Data.SqlDbType]::Char, 20)
        $cmd.Parameters["@Conto"].Value = $Conto
        [void]$cmd.Parameters.AddWithValue("@Inizio", $DataInizio.Date)
        [void]$cmd.Parameters.AddWithValue("@Fine",   $DataFine.Date.AddDays(1))
        $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dt = New-Object System.Data.DataTable
        [void]$da.Fill($dt)
        return ,$dt
    } finally { $cn.Close() }
}

$docs      = Invoke-Query $sqlDoc
$senzaXml  = Invoke-Query $sqlSenzaXml
Write-Host ("Documenti FTE trovati: {0} - Registrazioni senza XML: {1}`n" -f $docs.Rows.Count, $senzaXml.Rows.Count)

# --------------------------- ELABORAZIONE ----------------------------
$script:nomiUsati = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$log = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($r in $docs.Rows) {
    $i++
    $riga = [ordered]@{
        Id_CGMovT      = $r.Id_CGMovT
        Id_DmsDocument = $r.Id_DmsDocument
        DataReg        = ([datetime]$r.DtReg).ToString("dd/MM/yyyy")
        Protocollo     = ("$($r.NumProt)").Trim()
        Cd_CF          = ("$($r.Cd_CF)").Trim()
        Fornitore      = "$($r.Fornitore)"
        NumRif         = ("$($r.NumRif)").Trim()
        NumeroXml      = ""
        DataXml        = ""
        TotaleXml      = ""
        NumCoerente    = ""
        FileXml        = ""
        Pdf            = ""
        Esito          = ""
        Errore         = ""
    }
    try {
        $numRif = Get-NomeSicuro $riga.NumRif
        if (-not $numRif) { $numRif = "MOV$($r.Id_CGMovT)" }
        $base = Get-NomeUnivoco ("{0}_{1}_{2}_{3}" -f (Get-NomeSicuro $Conto), (Get-NomeSicuro $riga.Cd_CF), $numRif, ([datetime]$r.DtReg).ToString($FmtData))
        $xmlPath = Join-Path $out "$base.xml"
        $pdfPath = Join-Path $out "$base.pdf"

        Write-Host ("[{0}/{1}] {2}" -f $i, $docs.Rows.Count, $base)

        # Decompressione + sbustamento
        $b = Expand-ArcaBlob ([byte[]]$r.Content)
        if ("$($r.FileName)" -match '\.p7m$') { $xmlBytes = Get-XmlDaP7m $b } else { $xmlBytes = $b }
        [IO.File]::WriteAllBytes($xmlPath, $xmlBytes)
        $riga.FileXml = "$base.xml"

        # Dati di controllo dall'XML
        [xml]$x = [Text.Encoding]::UTF8.GetString($xmlBytes).TrimStart([char]0xFEFF)
        $dg = $x.SelectSingleNode("//*[local-name()='DatiGeneraliDocumento']")
        if ($dg) {
            $riga.NumeroXml = "$($dg.Numero)"
            $riga.DataXml   = "$($dg.Data)"
            $riga.TotaleXml = "$($dg.ImportoTotaleDocumento)"
            $riga.NumCoerente = if ($riga.NumeroXml.Trim() -eq $riga.NumRif) { "SI" } else { "NO" }
        }

        # Allegati: primo PDF -> <base>.pdf, altri con suffisso
        $n = 0
        foreach ($a in $x.SelectNodes("//*[local-name()='Allegati']")) {
            $n++
            $nomeAll = "$($a.SelectSingleNode("*[local-name()='NomeAttachment']").InnerText)"
            $dati = [Convert]::FromBase64String(($a.SelectSingleNode("*[local-name()='Attachment']").InnerText -replace '\s', ''))
            if ((Test-IsPdf $dati) -and -not $riga.Pdf) {
                [IO.File]::WriteAllBytes($pdfPath, $dati)
                $riga.Pdf = "ALLEGATO"
            } else {
                [IO.File]::WriteAllBytes((Join-Path $out ("{0}_allegato{1}_{2}" -f $base, $n, (Get-NomeSicuro $nomeAll))), $dati)
            }
        }

        # PDF mancante -> generato dal foglio di stile
        if (-not $riga.Pdf) {
            if ($puoGenerarePdf -and (New-PdfDaXml $xmlPath $pdfPath)) { $riga.Pdf = "GENERATO" }
            else { $riga.Pdf = "MANCANTE" }
        }

        $riga.Esito = "OK"
        $col = if ($riga.Pdf -eq "MANCANTE") { "Yellow" } else { "Green" }
        Write-Host ("    XML ok - PDF {0}" -f $riga.Pdf) -ForegroundColor $col
    } catch {
        $riga.Esito  = "ERRORE"
        $riga.Errore = $_.Exception.Message
        Write-Host ("    ERRORE: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
    $log.Add([pscustomobject]$riga)
}

# Registrazioni senza XML nel DMS
foreach ($r in $senzaXml.Rows) {
    $log.Add([pscustomobject][ordered]@{
        Id_CGMovT = $r.Id_CGMovT; Id_DmsDocument = ""
        DataReg = ([datetime]$r.DtReg).ToString("dd/MM/yyyy")
        Protocollo = ("$($r.NumProt)").Trim(); Cd_CF = ("$($r.Cd_CF)").Trim(); Fornitore = "$($r.Fornitore)"
        NumRif = ("$($r.NumRif)").Trim(); NumeroXml = ""; DataXml = ""; TotaleXml = ""; NumCoerente = ""
        FileXml = ""; Pdf = ""; Esito = "NESSUN XML"; Errore = ""
    })
}

# ------------------------------ RIEPILOGO ----------------------------
$csv = Join-Path $out "_riepilogo.csv"
$log | Export-Csv -Path $csv -Delimiter ';' -NoTypeInformation -Encoding UTF8

$ok       = @($log | Where-Object Esito -eq 'OK').Count
$err      = @($log | Where-Object Esito -eq 'ERRORE').Count
$pdfAll   = @($log | Where-Object Pdf -eq 'ALLEGATO').Count
$pdfGen   = @($log | Where-Object Pdf -eq 'GENERATO').Count
$pdfMan   = @($log | Where-Object Pdf -eq 'MANCANTE').Count
$numNo    = @($log | Where-Object NumCoerente -eq 'NO').Count

Write-Host "`n================ RIEPILOGO ================"
Write-Host "XML esportati       : $ok"
Write-Host "Errori              : $err"   -ForegroundColor $(if ($err) { 'Red' } else { 'Gray' })
Write-Host "PDF allegati        : $pdfAll"
Write-Host "PDF generati        : $pdfGen"
Write-Host "PDF mancanti        : $pdfMan" -ForegroundColor $(if ($pdfMan) { 'Yellow' } else { 'Gray' })
Write-Host "Registrazioni senza XML: $($senzaXml.Rows.Count)" -ForegroundColor $(if ($senzaXml.Rows.Count) { 'Yellow' } else { 'Gray' })
Write-Host "Numero ARCA <> XML  : $numNo"  -ForegroundColor $(if ($numNo) { 'Yellow' } else { 'Gray' })
Write-Host "Log: $csv"
