$origem  = 'W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Arquivos Brutos\202603'
$destino = 'W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Arquivos PDFA\202603'

$gsExe      = 'C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe'
$iccProfile = 'C:\WINDOWS\System32\spool\drivers\color\sRGB Color Space Profile.icm'

New-Item -ItemType Directory -Force -Path $destino | Out-Null

if (-not (Test-Path $gsExe)) {
    throw "Ghostscript nao encontrado em: $gsExe"
}

if (-not (Test-Path $iccProfile)) {
    throw "Perfil ICC nao encontrado em: $iccProfile"
}

Write-Host "ICC usado:" $iccProfile

# Ghostscript/PS prefere barras normais no caminho dentro do arquivo .ps
$iccProfilePS = $iccProfile -replace '\\','/'

# Cria um PDFA_def.ps temporario, sem depender do arquivo instalado do Ghostscript
$pdfaDefTemp = Join-Path $env:TEMP 'PDFA_def_custom.ps'

$pdfaDefContent = @"
%!
/ICCProfile ($iccProfilePS) def

[ /Title (Title)
  /DOCINFO pdfmark

[/_objdef {icc_PDFA} /type /stream /OBJ pdfmark
[{icc_PDFA} <</N systemdict /ProcessColorModel get /DeviceGray eq {1} {systemdict /ProcessColorModel get /DeviceRGB eq {3} {4} ifelse} ifelse >> /PUT pdfmark
[{icc_PDFA} ICCProfile (r) file /PUT pdfmark

[/_objdef {OutputIntent_PDFA} /type /dict /OBJ pdfmark
[{OutputIntent_PDFA} <<
  /Type /OutputIntent
  /S /GTS_PDFA1
  /DestOutputProfile {icc_PDFA}
  /OutputConditionIdentifier (sRGB)
>> /PUT pdfmark

[{Catalog} <</OutputIntents [ {OutputIntent_PDFA} ]>> /PUT pdfmark
"@

Set-Content -Path $pdfaDefTemp -Value $pdfaDefContent -Encoding ASCII

function Get-UniqueOutputPath {
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Extension = '.pdf'
    )

    $candidate = Join-Path $Folder ($BaseName + $Extension)
    $i = 2

    while (Test-Path $candidate) {
        $candidate = Join-Path $Folder ("{0}_{1}{2}" -f $BaseName, $i, $Extension)
        $i++
    }

    return $candidate
}

$ok = 0
$falha = 0

Get-ChildItem -Path $origem -Recurse -File -Filter *.pdf |
    Where-Object { $_.FullName -notlike "$destino*" } |
    ForEach-Object {

        $entrada = $_.FullName
        $saida   = Get-UniqueOutputPath -Folder $destino -BaseName $_.BaseName

        Write-Host ""
        Write-Host "Convertendo:" $entrada
        Write-Host "Saida.....:" $saida

        & $gsExe `
            -dPDFA=2 `
            -dBATCH `
            -dNOPAUSE `
            -sDEVICE=pdfwrite `
            -sColorConversionStrategy=RGB `
            -sBlendConversionStrategy=Simple `
            -dPDFACompatibilityPolicy=1 `
            "--permit-file-read=$iccProfile" `
            "-sOutputFile=$saida" `
            $pdfaDefTemp `
            $entrada

        if ($LASTEXITCODE -eq 0) {
            $ok++
        }
        else {
            $falha++
            Write-Warning "Falhou: $entrada"
        }
    }

Write-Host ""
Write-Host "Concluido."
Write-Host "Convertidos com sucesso:" $ok
Write-Host "Falhas:" $falha