param(
    [string]$PythonExe = "W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Venv_Leonardo\Venv_leov1\Scripts\python.exe",
    [string]$ScriptPy = "",
    [string]$Pattern = "https://www.nfse.gov.br/*",
    [switch]$SomenteListar
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ScriptPy)) {
    $ScriptPy = Join-Path $PSScriptRoot "nfse_nacional_url_concorrente.py"
}

function Parse-Dn {
    param([string]$Dn)
    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Dn)) { return $map }

    $parts = $Dn -split ",\s*"
    foreach ($p in $parts) {
        if ($p -match "^\s*([^=]+)=(.+)\s*$") {
            $k = $matches[1].Trim().ToUpperInvariant()
            $v = $matches[2].Trim()
            if (-not [string]::IsNullOrWhiteSpace($k) -and -not [string]::IsNullOrWhiteSpace($v)) {
                # Se houver chave repetida (ex.: OU), mantem a primeira.
                if (-not $map.ContainsKey($k)) {
                    $map[$k] = $v
                }
            }
        }
    }
    return $map
}

function Test-EcnpjCandidate {
    param([System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert)

    if ($null -eq $Cert) { return $false }
    if ($Cert.NotAfter -le (Get-Date)) { return $false }
    if (-not $Cert.HasPrivateKey) { return $false }

    $subject = [string]$Cert.Subject
    if ([string]::IsNullOrWhiteSpace($subject)) { return $false }
    if ($subject -notmatch "ICP-Brasil") { return $false }
    if ($subject -match "e-CPF") { return $false }
    if ($subject -match "e-CNPJ") { return $true }

    # Fallback: extrai CNPJ do CN apos ':' aceitando pontuacao.
    # Ex.: "CN=EMPRESA XYZ:07.053.914/0001-60, ..."
    if ($subject -match "CN=([^,]+)") {
        $cn = $matches[1]
        if ($cn -match ":(.+)$") {
            $raw = $matches[1]
            $digits = ($raw -replace "\D", "")
            if ($digits.Length -eq 14) { return $true }
        }
    }

    return $false
}

function Build-PolicyJson {
    param(
        [string]$PatternValue,
        [hashtable]$SubjectMap,
        [hashtable]$IssuerMap
    )

    $subject = @{}
    foreach ($k in @("CN", "O", "C", "S", "L")) {
        if ($SubjectMap.ContainsKey($k)) { $subject[$k] = $SubjectMap[$k] }
    }

    $issuer = @{}
    foreach ($k in @("CN", "O", "C")) {
        if ($IssuerMap.ContainsKey($k)) { $issuer[$k] = $IssuerMap[$k] }
    }

    $obj = @{
        pattern = $PatternValue
        filter  = @{
            SUBJECT = $subject
            ISSUER  = $issuer
        }
    }

    return ($obj | ConvertTo-Json -Compress -Depth 10)
}

function Set-ChromePolicySingle {
    param([string]$PolicyValue)

    $regUser = "HKCU:\SOFTWARE\Policies\Google\Chrome\AutoSelectCertificateForUrls"
    $regMachine = "HKLM:\SOFTWARE\Policies\Google\Chrome\AutoSelectCertificateForUrls"

    foreach ($reg in @($regUser, $regMachine)) {
        try {
            if (Test-Path $reg) {
                Remove-Item -Path $reg -Recurse -Force
            }
            New-Item -Path $reg -Force | Out-Null
            Set-ItemProperty -Path $reg -Name "1" -Type String -Value $PolicyValue
        } catch {
            Write-Host ("[AVISO] Nao foi possivel atualizar policy em {0}: {1}" -f $reg, $_.Exception.Message)
        }
    }
}

function Close-Browsers {
    foreach ($name in @("chrome", "msedge")) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path $PythonExe)) {
    throw "Python nao encontrado em: $PythonExe"
}
if (-not (Test-Path $ScriptPy)) {
    throw "Script Python nao encontrado em: $ScriptPy"
}

$certsCurrentUser = @(
    Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { Test-EcnpjCandidate $_ }
)

$certsLocalMachine = @(
    Get-ChildItem "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue | Where-Object { Test-EcnpjCandidate $_ }
)

$certs = @(
    $certsCurrentUser + $certsLocalMachine |
        Sort-Object Thumbprint -Unique |
        Sort-Object Subject
)

if (-not $certs -or $certs.Count -eq 0) {
    throw "Nenhum certificado e-CNPJ valido encontrado em Cert:\CurrentUser\My"
}

Write-Host "Certificados encontrados: $($certs.Count)"
Write-Host " - CurrentUser\\My: $($certsCurrentUser.Count)"
Write-Host " - LocalMachine\\My: $($certsLocalMachine.Count)"
foreach ($c in $certs) {
    $storeHint = if ($certsCurrentUser.Thumbprint -contains $c.Thumbprint) { "CurrentUser" } else { "LocalMachine" }
    Write-Host (" - [{0}] {1} | vence {2}" -f $storeHint, $c.Subject, $c.NotAfter.ToString("dd/MM/yyyy HH:mm:ss"))
}

if ($SomenteListar) {
    Write-Host "Modo SomenteListar ativado. Nenhuma execucao foi iniciada."
    exit 0
}

$results = @()
foreach ($cert in $certs) {
    $sub = Parse-Dn -Dn $cert.Subject
    $iss = Parse-Dn -Dn $cert.Issuer
    $policy = Build-PolicyJson -PatternValue $Pattern -SubjectMap $sub -IssuerMap $iss

    Write-Host ""
    Write-Host "============================================================"
    Write-Host ("Executando certificado: {0}" -f ($sub["CN"]))
    Write-Host ("Thumbprint: {0}" -f $cert.Thumbprint)
    Write-Host "Aplicando policy do Chrome..."
    Set-ChromePolicySingle -PolicyValue $policy

    Write-Host "Fechando Chrome/Edge..."
    Close-Browsers

    Write-Host "Iniciando script Python..."
    $env:NFSE_CERT_CN = $sub["CN"]
    & $PythonExe $ScriptPy
    $exitCode = $LASTEXITCODE
    Remove-Item Env:NFSE_CERT_CN -ErrorAction SilentlyContinue

    $results += [PSCustomObject]@{
        CertCN    = $sub["CN"]
        Thumbprint = $cert.Thumbprint
        ExitCode  = $exitCode
    }

    Write-Host ("Finalizado com ExitCode={0}" -f $exitCode)
}

Write-Host ""
Write-Host "======================= RESUMO ======================="
$results | Format-Table -AutoSize
